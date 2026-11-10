"""Windows IMAP/POP3 mail notifier with bounded message previews."""
import email
from email.header import decode_header
from email.parser import BytesHeaderParser
from email.utils import parsedate_to_datetime, parseaddr
from datetime import datetime
from html import unescape
from html.parser import HTMLParser
import re
import imaplib
import json
import os
import sys
import subprocess
import socket
import uuid
import winreg
from pathlib import Path
import poplib
import queue
import time
import threading
import webbrowser

import keyring
import base64
import ctypes
from ctypes import wintypes
from winotify import Notification, audio

APP = '메일 알리미'
START_IN_TRAY = '--tray' in sys.argv
SERVICE = 'MailTitleNotifier'
# Portable storage next to the executable (not in PyInstaller's temporary _MEIPASS).
DATA = Path(sys.executable).resolve().parent if getattr(sys, 'frozen', False) else Path(__file__).resolve().parent
CONFIG = DATA / 'settings.json'
SECRETS = DATA / 'credentials.json'
LEGACY_DATA = Path(os.environ.get('LOCALAPPDATA', str(Path.home()))) / 'MailTitleNotifier'
LEGACY_CONFIG = LEGACY_DATA / 'settings.json'
MIGRATION_MARKER = DATA / '.legacy_migration_complete'
RUN_KEY = r'Software\Microsoft\Windows\CurrentVersion\Run'
RUN_NAME = 'MailTitleNotifier'
URI_SCHEME = 'mailtitlenotifier'
IPC_PORT = 49327
NOTIFY_ICON = DATA / 'notification_icon.png'

# Official mail home pages. This opens the mailbox, not an unreliable message deep link.
MAIL_SITES = {
    'daum': ('다음 메일 열기', 'https://mail.daum.net/'),
    'naver': ('네이버 메일 열기', 'https://mail.naver.com/'),
    'gmail': ('지메일 열기', 'https://mail.google.com/'),
}

def mailbox_url(account):
    """Identify provider by configured login/server, not a user-defined nickname."""
    user = str(account.get('username', '')).strip().lower()
    host = str(account.get('host', '')).strip().lower().rstrip('.')
    domain = user.rsplit('@', 1)[-1] if '@' in user else ''
    if domain in ('daum.net', 'hanmail.net') or host == 'daum.net' or host.endswith('.daum.net'):
        return MAIL_SITES['daum'][1]
    if domain == 'naver.com' or host == 'naver.com' or host.endswith('.naver.com'):
        return MAIL_SITES['naver'][1]
    if domain in ('gmail.com', 'googlemail.com') or host in ('imap.gmail.com', 'pop.gmail.com'):
        return MAIL_SITES['gmail'][1]
    return None




class _DataBlob(ctypes.Structure):
    _fields_ = [('cbData', wintypes.DWORD), ('pbData', ctypes.POINTER(ctypes.c_byte))]


def _blob(data):
    buf = ctypes.create_string_buffer(data)
    return _DataBlob(len(data), ctypes.cast(buf, ctypes.POINTER(ctypes.c_byte))), buf


def _dpapi(data, protect):
    """Windows DPAPI encrypt/decrypt tied to current Windows user."""
    inp, keeper = _blob(data)
    out = _DataBlob()
    fn = ctypes.windll.crypt32.CryptProtectData if protect else ctypes.windll.crypt32.CryptUnprotectData
    if protect:
        ok = fn(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out))
    else:
        ok = fn(ctypes.byref(inp), None, None, None, None, 0, ctypes.byref(out))
    if not ok:
        raise ctypes.WinError()
    try:
        return ctypes.string_at(out.pbData, out.cbData)
    finally:
        free = ctypes.windll.kernel32.LocalFree
        free.argtypes = [ctypes.c_void_p]
        free.restype = ctypes.c_void_p
        free(ctypes.cast(out.pbData, ctypes.c_void_p))


def _read_secrets():
    try:
        return json.loads(SECRETS.read_text(encoding='utf-8'))
    except FileNotFoundError:
        return {}


def _write_json(path, value):
    path.parent.mkdir(parents=True, exist_ok=True)
    temp = path.with_name(path.name + '.tmp')
    temp.write_text(json.dumps(value, ensure_ascii=False, indent=2), encoding='utf-8')
    temp.replace(path)


def save_secret(key, password):
    secrets = _read_secrets()
    secrets[key] = base64.b64encode(_dpapi(password.encode('utf-8'), True)).decode('ascii')
    _write_json(SECRETS, secrets)


def load_secret(key):
    stored = _read_secrets().get(key)
    if stored:
        return _dpapi(base64.b64decode(stored), False).decode('utf-8')
    # Upgrade earlier versions: migrate the Windows Credential Manager entry on use.
    legacy = keyring.get_password(SERVICE, key)
    if legacy:
        save_secret(key, legacy)
        try:
            keyring.delete_password(SERVICE, key)
        except Exception:
            pass
    return legacy


def delete_unreferenced_secrets(accounts):
    active = {a['credential_key'] for a in accounts}
    secrets = _read_secrets()
    remaining = {k: v for k, v in secrets.items() if k in active}
    if remaining != secrets:
        _write_json(SECRETS, remaining)


def short_received(value):
    # Normalized IMAP INTERNALDATE / email Date use local time; omit year and seconds.
    try:
        return datetime.strptime(value, '%Y-%m-%d %H:%M:%S').strftime('%m-%d %H:%M')
    except (ValueError, TypeError):
        return value or '(시간 정보 없음)'


def decoded_subject(value):
    if not value:
        return '(제목 없음)'
    try:
        segments = decode_header(value)
        pieces = []
        for part, charset in segments:
            if isinstance(part, bytes):
                pieces.append(part.decode(charset or 'utf-8', errors='replace'))
            else:
                pieces.append(part)
        return ''.join(pieces).strip() or '(제목 없음)'
    except (LookupError, ValueError, UnicodeError):
        return str(value)


class HTMLText(HTMLParser):
    def __init__(self):
        super().__init__(convert_charrefs=True)
        self.parts = []
        self.skip = 0

    def handle_starttag(self, tag, attrs):
        if tag in ('script', 'style'):
            self.skip += 1
        elif tag in ('p', 'br', 'div', 'li'):
            self.parts.append('\n')

    def handle_endtag(self, tag):
        if tag in ('script', 'style'):
            self.skip = max(0, self.skip - 1)
        elif tag in ('p', 'div', 'li'):
            self.parts.append('\n')

    def handle_data(self, data):
        if not self.skip:
            self.parts.append(data)


def short_preview(raw, html=False):
    if html:
        parser = HTMLText()
        parser.feed(raw)
        raw = unescape(''.join(parser.parts))
    lines = [re.sub(r'\s+', ' ', x).strip() for x in raw.splitlines()]
    lines = [x for x in lines if x]
    out = []
    for line in lines:
        while len(line) > 120 and len(out) < 3:
            out.append(line[:120])
            line = line[120:]
        if line:
            out.append(line)
        if len(out) >= 3:
            break
    return '\n'.join(out[:3]) or '(미리보기 없음)'


def decode_part(part):
    payload = part.get_payload(decode=True)
    if payload is None:
        payload = part.get_payload()
        if isinstance(payload, str):
            return payload
        return ''
    for charset in (part.get_content_charset(), 'utf-8', 'cp949', 'latin-1'):
        if not charset:
            continue
        try:
            return payload.decode(charset)
        except (UnicodeError, LookupError):
            continue
    return payload.decode('utf-8', errors='replace')


def read_message(raw, received=None):
    msg = email.message_from_bytes(raw)
    raw_sender = decoded_subject(msg.get('From')) if msg.get('From') else ''
    display_name, address = parseaddr(raw_sender)
    # Prefer the sender's display name, falling back to the email address if absent.
    sender = display_name.strip().strip('\"') or address or '(발신자 정보 없음)'
    subject = decoded_subject(msg.get('Subject'))
    when = received
    if not when:
        try:
            when = parsedate_to_datetime(msg.get('Date')).astimezone().strftime('%Y-%m-%d %H:%M:%S')
        except (TypeError, ValueError, OverflowError, AttributeError):
            when = '(시간 정보 없음)'
    plain = html = None
    for part in msg.walk():
        if part.is_multipart() or part.get_content_disposition() == 'attachment':
            continue
        typ = part.get_content_type()
        if typ == 'text/plain' and plain is None:
            plain = decode_part(part)
        elif typ == 'text/html' and html is None:
            html = decode_part(part)
    preview = short_preview(plain) if plain is not None else short_preview(html, True) if html is not None else '(미리보기 없음)'
    return {'subject': subject, 'sender': sender, 'received': when, 'preview': preview}


def get_imap(config, password, old):
    host, port = config['host'], int(config['port'])
    with imaplib.IMAP4_SSL(host, port, timeout=20) as conn:
        conn.login(config['username'], password)
        status, _ = conn.select('INBOX', readonly=True)
        if status != 'OK':
            raise RuntimeError('INBOX 폴더를 열 수 없습니다.')
        valid = conn.response('UIDVALIDITY')[1]
        uidvalidity = (valid[0].decode('ascii', 'replace') if valid and valid[0] else '')
        previous = old if old.get('type') == 'IMAP' and old.get('uidvalidity') == uidvalidity else {}
        last = int(previous.get('last_uid', 0))
        criteria = f'{last + 1}:*' if previous else 'ALL'
        status, parts = conn.uid('search', None, criteria)
        if status != 'OK':
            raise RuntimeError('IMAP UID 검색에 실패했습니다.')
        uids = [int(x) for x in (parts[0] if parts else b'').split()]
        uids = [uid for uid in uids if uid > last]
        latest = max(uids, default=last)
        messages = []
        if previous:
            for uid in uids[-50:]:
                # Read at most 32 KiB per new message; PEEK never marks it read.
                typ, response = conn.uid('fetch', str(uid), '(INTERNALDATE BODY.PEEK[]<0.32768>)')
                if typ != 'OK':
                    continue
                raw = b''.join(part[1] for part in response if isinstance(part, tuple) and isinstance(part[1], bytes))
                received = None
                for part in response:
                    if isinstance(part, tuple):
                        match = re.search(rb'INTERNALDATE "([^"]+)"', part[0])
                        if match:
                            try:
                                received = datetime.strptime(match.group(1).decode('ascii'), '%d-%b-%Y %H:%M:%S %z').astimezone().strftime('%Y-%m-%d %H:%M:%S')
                            except (ValueError, UnicodeError):
                                pass
                if raw:
                    messages.append(read_message(raw, received))
        return {'type': 'IMAP', 'uidvalidity': uidvalidity, 'last_uid': latest}, messages, len(uids)


def get_pop3(config, password, old):
    with poplib.POP3_SSL(config['host'], int(config['port']), timeout=20) as conn:
        conn.user(config['username'])
        conn.pass_(password)
        _, lines, _ = conn.uidl()
        pairs = []
        for line in lines:
            idx, uid = line.decode('ascii', errors='replace').split(maxsplit=1)
            pairs.append((int(idx), uid))
        seen = set(old.get('seen', [])) if old.get('type') == 'POP3' else set()
        new = [(idx, uid) for idx, uid in pairs if uid not in seen]
        messages = []
        if old.get('type') == 'POP3':
            for idx, _ in new[-50:]:
                # TOP requests message headers and at most 60 lines of body.
                try:
                    _, headers_and_body, _ = conn.top(idx, 60)
                except poplib.error_proto as exc:
                    raise RuntimeError('POP3 TOP 명령을 지원하지 않는 서버입니다.') from exc
                raw = b'\r\n'.join(headers_and_body) + b'\r\n'
                messages.append(read_message(raw))
        state = {'type': 'POP3', 'seen': [uid for _, uid in pairs][-5000:]}
        return state, messages, len(new)



from PySide6.QtCore import Qt, QTimer, Signal, QObject, QSize
from PySide6.QtGui import QIcon, QAction, QFont
from PySide6.QtWidgets import (QApplication, QMainWindow, QWidget, QVBoxLayout, QHBoxLayout,
    QLabel, QPushButton, QLineEdit, QComboBox, QSpinBox, QCheckBox, QListWidget,
    QTableWidget, QTableWidgetItem, QHeaderView, QGroupBox, QGridLayout, QFrame,
    QMessageBox, QSystemTrayIcon, QMenu, QAbstractItemView, QScrollArea, QSizePolicy)

ICON_PATH = Path(getattr(sys, '_MEIPASS', Path(__file__).resolve().parent)) / 'mail_icon.ico'

STYLE = """
QWidget { background: #f3f7fd; color: #213454; font-family: 'Malgun Gothic'; font-size: 10pt; }
QMainWindow { background: #f3f7fd; }
QLabel#headline { font-size: 20pt; font-weight: 700; color: #1c416e; }
QLabel#subtitle { color: #71819a; font-size: 9pt; }
QLabel#section { font-size: 11pt; font-weight: 700; color: #254a80; }
QLabel#footnote { color: #8c99aa; font-size: 9pt; }
QFrame#card { background: white; border: 1px solid #e2ebf7; border-radius: 12px; }
QFrame#card QLabel { background: transparent; }
QLineEdit, QComboBox, QSpinBox { border: 1px solid #d4dfef; background: #fff; border-radius: 5px; padding: 7px; min-height: 20px; color: #233552; }
QLineEdit:focus, QComboBox:focus, QSpinBox:focus { border: 1px solid #3579d3; }
QPushButton { background: #eaf2fc; color: #245992; border: 1px solid #bad3f0; border-radius: 7px; padding: 9px 14px; min-height: 24px; font-weight: 600; }
QPushButton:hover { background: #d9eafd; border-color: #8ab9f1; }
QPushButton:pressed { background: #c9dff9; }
QPushButton:focus { border: 1px solid #3579d3; }
QPushButton#primary { background: #186bca; color: #ffffff; border: 1px solid #186bca; font-weight: 700; }
QPushButton#primary:hover { background: #1c65b8; color: #ffffff; }
QPushButton#primary:pressed { background: #14549a; color: #ffffff; }
QPushButton#primary:disabled { background: #dce7f4; border-color: #dce7f4; color: #697d99; }
QPushButton#danger { color: #ad3346; background: #fff2f3; border-color: #edc6cd; }
QPushButton#danger:hover { background: #ffe4e8; border-color: #e6aeb8; }
QListWidget, QTableWidget { background: white; border: 1px solid #e0e9f5; border-radius: 5px; outline: none; alternate-background-color: #f6f9fe; }
QListWidget::item { padding: 11px; }
QListWidget::item:selected { color: #1c5c9c; background: #e7f1ff; }
QHeaderView::section { background: #f0f5fc; color: #446084; padding: 9px; border: 0px; font-weight: 700; }
QTableWidget::item { padding: 6px; }
QCheckBox { spacing: 8px; }
"""

class Popup(QWidget):
    def __init__(self, owner, account, msg, url):
        super().__init__(None, Qt.Tool | Qt.FramelessWindowHint | Qt.WindowStaysOnTopHint)
        self.owner, self.url = owner, url
        self.setFixedWidth(410)
        self.setObjectName('popup')
        self.setStyleSheet('QWidget#popup { background: white; border: 1px solid #9cbddd; border-radius: 9px; } QLabel {background: transparent;}')
        outer = QVBoxLayout(self)
        outer.setContentsMargins(15, 12, 15, 13)
        header = QHBoxLayout()
        title = QLabel(f'[{account}] {msg["subject"]}')
        title.setFont(QFont('Malgun Gothic', 11, QFont.Bold))
        title.setWordWrap(True)
        close = QPushButton('×')
        close.setFixedSize(32, 30)
        close.clicked.connect(self.dismiss)
        header.addWidget(title, 1); header.addWidget(close)
        outer.addLayout(header)
        info = QLabel(f'발신자: {msg["sender"]}  ·  수신: {short_received(msg["received"])}')
        info.setStyleSheet('color: #637993;')
        outer.addWidget(info)
        preview = QLabel(msg.get('preview') or '(본문 미리보기 없음)')
        preview.setWordWrap(True)
        outer.addWidget(preview)
        self.adjustSize()
        QTimer.singleShot(owner.seconds() * 1000, self.dismiss)

    def mousePressEvent(self, event):
        if event.button() == Qt.LeftButton:
            if self.url:
                webbrowser.open(self.url)
            else:
                self.owner.show_window()
            self.dismiss()
        super().mousePressEvent(event)

    def dismiss(self):
        if self in self.owner.popups:
            self.owner.popups.remove(self)
        self.close()
        self.owner.position_popups()


class Window(QMainWindow):
    def __init__(self):
        super().__init__()
        self.closed = False
        self.busy = False
        self.events = queue.Queue()
        # Windows toast APIs may block; keep them off the Qt GUI thread.
        self.toast_events = queue.Queue(maxsize=250)
        self.toast_thread = threading.Thread(target=self._toast_worker, daemon=True)
        self.toast_thread.start()
        self._mail_state_dirty = False
        self.config = self.load_config()
        self.accounts = self.config.get('accounts', [])
        self.selected_id = None
        self.popups = []
        self.setWindowTitle(APP)
        self.setWindowIcon(QIcon(str(ICON_PATH)))
        self.resize(1120, 760)
        self.setMinimumSize(970, 680)
        self.build_ui()
        self.setup_tray()
        self.start_ipc()
        self.tick = QTimer(self); self.tick.timeout.connect(self.process_events); self.tick.start(200)
        self.schedule = QTimer(self); self.schedule.timeout.connect(self.periodic); self.schedule.start(1000)
        if not START_IN_TRAY and not self.accounts:
            self.show()
        if self.accounts:
            QTimer.singleShot(800, self.poll)

    @staticmethod
    def load_config():
        try:
            if not CONFIG.exists() and not MIGRATION_MARKER.exists() and LEGACY_CONFIG.is_file():
                _write_json(CONFIG, json.loads(LEGACY_CONFIG.read_text(encoding='utf-8')))
                MIGRATION_MARKER.touch()
            cfg = json.loads(CONFIG.read_text(encoding='utf-8'))
            if 'accounts' not in cfg and cfg.get('host'):
                cfg = {'accounts': [dict(cfg, id=uuid.uuid4().hex, label=cfg.get('username', '계정'))]}
            return cfg
        except (OSError, ValueError):
            return {'accounts': []}

    def card(self, title, parent_layout):
        frame = QFrame(); frame.setObjectName('card')
        box = QVBoxLayout(frame); box.setContentsMargins(20, 17, 20, 18); box.setSpacing(12)
        heading = QLabel(title); heading.setObjectName('section'); box.addWidget(heading)
        parent_layout.addWidget(frame)
        return box

    def build_ui(self):
        root = QWidget(); self.setCentralWidget(root)
        full = QVBoxLayout(root); full.setContentsMargins(22, 20, 22, 12); full.setSpacing(16)
        head = QHBoxLayout()
        title_col = QVBoxLayout(); title_col.setSpacing(2)
        h = QLabel('✉  메일 알리미'); h.setObjectName('headline'); title_col.addWidget(h)
        s = QLabel('메일 계정과 수신 알림을 한곳에서 관리합니다.'); s.setObjectName('subtitle'); title_col.addWidget(s)
        head.addLayout(title_col); head.addStretch()
        self.state = QLabel('●  대기 중'); self.state.setStyleSheet('color: #2975bd; font-weight: 600;')
        head.addWidget(self.state); full.addLayout(head)
        columns = QHBoxLayout(); columns.setSpacing(16); full.addLayout(columns, 1)
        # The account panel can be taller than the available window height.
        # Scroll that column instead of allowing Qt to squeeze/overlap rows.
        left_scroll = QScrollArea()
        left_scroll.setWidgetResizable(True)
        left_scroll.setFrameShape(QFrame.NoFrame)
        left_scroll.setHorizontalScrollBarPolicy(Qt.ScrollBarAlwaysOff)
        left_scroll.setStyleSheet('QScrollArea { background: transparent; border: none; }')
        left_content = QWidget()
        left_content.setStyleSheet('background: transparent;')
        left = QVBoxLayout(left_content)
        left.setContentsMargins(0, 0, 8, 0)
        left.setSpacing(14)
        left_scroll.setWidget(left_content)
        columns.addWidget(left_scroll, 4)
        right = QVBoxLayout(); right.setSpacing(14); columns.addLayout(right, 6)

        ac = self.card('메일 계정 목록', left)
        self.accounts_list = QListWidget()
        self.accounts_list.setMinimumHeight(120)
        self.accounts_list.setMaximumHeight(190)
        self.accounts_list.setSpacing(3)
        self.accounts_list.currentRowChanged.connect(self.choose_account)
        ac.addWidget(self.accounts_list)
        actions = QHBoxLayout()
        new = QPushButton('＋  새 계정'); new.clicked.connect(self.new_account)
        delete = QPushButton('계정 삭제'); delete.setObjectName('danger'); delete.clicked.connect(self.delete_account)
        actions.addWidget(new); actions.addWidget(delete); ac.addLayout(actions)

        form = self.card('메일 계정 설정', left)
        grid = QGridLayout(); grid.setHorizontalSpacing(12); grid.setVerticalSpacing(13)
        self.label = QLineEdit(); self.protocol = QComboBox(); self.protocol.addItems(['IMAP', 'POP3'])
        self.host = QLineEdit(); self.port = QSpinBox(); self.port.setRange(1, 65535)
        self.username = QLineEdit(); self.password = QLineEdit(); self.password.setEchoMode(QLineEdit.Password)
        self.interval = QSpinBox(); self.interval.setRange(15, 86400); self.interval.setSuffix(' 초')
        # Explicit row/control heights avoid font scaling + style padding clipping.
        for control in (self.label, self.protocol, self.host, self.port, self.username, self.password, self.interval):
            control.setMinimumHeight(39)
            control.setSizePolicy(QSizePolicy.Expanding, QSizePolicy.Fixed)
        self.port.setMinimumWidth(110)
        self.interval.setMinimumWidth(110)
        grid.setColumnStretch(1, 1)
        grid.setColumnMinimumWidth(0, 77)
        self.protocol.currentTextChanged.connect(self.protocol_changed)
        fields = [('계정 이름', self.label), ('메일 방식', self.protocol), ('서버 주소', self.host), ('포트', self.port), ('아이디', self.username), ('비밀번호', self.password)]
        for i, (label, widget) in enumerate(fields):
            grid.addWidget(QLabel(label), i, 0); grid.addWidget(widget, i, 1)
        grid.addWidget(QLabel('확인 간격'), 6, 0)
        intervalrow = QHBoxLayout(); intervalrow.setSpacing(9); intervalrow.addWidget(self.interval, 1)
        check = QPushButton('↻  지금 확인'); check.setMinimumHeight(39); check.clicked.connect(self.poll); intervalrow.addWidget(check)
        grid.addLayout(intervalrow, 6, 1)
        form.addLayout(grid)
        save = QPushButton('계정 저장'); save.setObjectName('primary'); save.setMinimumHeight(42); save.clicked.connect(self.save_account)
        form.addWidget(save)
        reset = QPushButton('모든 계정 및 저장 정보 초기화'); reset.setObjectName('danger'); reset.setMinimumHeight(42); reset.clicked.connect(self.reset_all)
        form.addWidget(reset)
        left.addStretch(1)

        opts = self.card('프로그램 설정', right)
        prefs = QGridLayout(); prefs.setVerticalSpacing(13); prefs.setHorizontalSpacing(12)
        self.mode = QComboBox(); self.mode.addItems(['Windows 기본 알림', '프로그램 팝업'])
        self.mode.setCurrentIndex(1 if self.config.get('notification_mode') == 'popup' else 0)
        seconds = self.config.get('notification_seconds', self.config.get('toast_duration', 25))
        seconds = 25 if seconds == 'long' else 5 if seconds == 'short' else seconds
        self.duration = QSpinBox(); self.duration.setRange(1, 300); self.duration.setSuffix(' 초')
        try: self.duration.setValue(int(seconds))
        except (ValueError, TypeError): self.duration.setValue(25)
        self.position = QComboBox()
        self.positions = {'오른쪽 아래':'bottom_right', '왼쪽 아래':'bottom_left', '오른쪽 위':'top_right', '왼쪽 위':'top_left'}
        self.position.addItems(list(self.positions))
        selected = self.config.get('popup_position','bottom_right')
        self.position.setCurrentIndex(list(self.positions.values()).index(selected) if selected in self.positions.values() else 0)
        self.autostart = QCheckBox('Windows 로그인 시 자동 실행 (트레이에서 시작)')
        self.autostart.setChecked(self.startup_enabled())
        for i, (label, widget) in enumerate([('알림 방식',self.mode),('표시 시간',self.duration),('팝업 위치',self.position)]):
            prefs.addWidget(QLabel(label),i,0); prefs.addWidget(widget,i,1)
        opts.addLayout(prefs); opts.addWidget(self.autostart)
        save_prefs = QPushButton('프로그램 설정 저장'); save_prefs.setObjectName('primary'); save_prefs.setMinimumHeight(42)
        save_prefs.clicked.connect(self.save_program_settings); opts.addWidget(save_prefs)

        recent = self.card('최근 수신 메일', right)
        info = QLabel('최근 확인한 메일의 제목을 표시합니다.'); info.setObjectName('subtitle'); recent.addWidget(info)
        self.subjects = QTableWidget(0, 3); self.subjects.setHorizontalHeaderLabels(['수신 시간', '발신자', '메일 제목'])
        self.subjects.setEditTriggers(QAbstractItemView.NoEditTriggers)
        self.subjects.setSelectionBehavior(QAbstractItemView.SelectRows)
        self.subjects.setAlternatingRowColors(True)
        self.subjects.verticalHeader().hide()
        self.subjects.horizontalHeader().setSectionResizeMode(0, QHeaderView.ResizeToContents)
        self.subjects.horizontalHeader().setSectionResizeMode(1, QHeaderView.ResizeToContents)
        self.subjects.horizontalHeader().setSectionResizeMode(2, QHeaderView.Stretch)
        self.subjects.cellDoubleClicked.connect(self.open_selected_mail)
        self.subjects.setMinimumHeight(260)
        recent.addWidget(self.subjects, 1)
        right.setStretchFactor(recent.parentWidget(), 1)
        footer = QHBoxLayout()
        self.status = QLabel('계정을 추가하거나 선택하세요.'); self.status.setObjectName('footnote')
        self.status.setWordWrap(True); footer.addWidget(self.status, 1)
        credit = QLabel('배포처: tokis.tistory.com'); credit.setObjectName('footnote'); footer.addWidget(credit)
        full.addLayout(footer)
        self.refresh_accounts(); self.new_account()

    def refresh_accounts(self):
        self.accounts_list.blockSignals(True)
        self.accounts_list.clear()
        for a in self.accounts:
            self.accounts_list.addItem(a.get('label') or a['username'])
        self.accounts_list.blockSignals(False)

    def choose_account(self, index):
        if index < 0 or index >= len(self.accounts): return
        a = self.accounts[index]; self.selected_id = a['id']
        self.label.setText(a.get('label','')); self.protocol.setCurrentText(a.get('type','IMAP'))
        self.host.setText(a.get('host','')); self.port.setValue(int(a.get('port',993)))
        self.username.setText(a.get('username','')); self.password.clear()
        self.interval.setValue(int(a.get('interval',60)))

    def new_account(self):
        self.selected_id = None
        self.accounts_list.blockSignals(True); self.accounts_list.clearSelection(); self.accounts_list.setCurrentRow(-1); self.accounts_list.blockSignals(False)
        for item in (self.label,self.host,self.username,self.password): item.clear()
        self.protocol.setCurrentText('IMAP'); self.port.setValue(993); self.interval.setValue(60)

    def protocol_changed(self, text):
        if self.port.value() in (993,995): self.port.setValue(993 if text == 'IMAP' else 995)

    def save_account(self):
        host, user = self.host.text().strip(), self.username.text().strip()
        port, interval = self.port.value(), self.interval.value()
        if not host or not user:
            QMessageBox.warning(self,APP,'서버 주소와 아이디를 입력하세요.'); return
        protocol = self.protocol.currentText(); key = f'{protocol}:{host}:{port}:{user}'
        existing = next((a for a in self.accounts if a['id'] == self.selected_id),None)
        pwd = self.password.text()
        try:
            if pwd: save_secret(key,pwd)
            elif not load_secret(key):
                QMessageBox.warning(self,APP,'새 계정 또는 로그인 정보 변경 시 비밀번호를 입력하세요.'); return
        except Exception as exc:
            QMessageBox.critical(self,APP,f'비밀번호 저장 실패: {exc}'); return
        a = {'id':existing['id'] if existing else uuid.uuid4().hex,
             'label':self.label.text().strip() or user,'type':protocol,'host':host,'port':port,
             'username':user,'interval':interval,'credential_key':key,
             'mail_state':existing.get('mail_state',{}) if existing and existing.get('credential_key')==key else {},
             'last_check':0}
        if existing: self.accounts[self.accounts.index(existing)] = a
        else: self.accounts.append(a)
        self.selected_id = a['id']; self.password.clear()
        try:
            delete_unreferenced_secrets(self.accounts); self.save_config()
        except OSError as exc:
            QMessageBox.critical(self,APP,f'설정 저장 실패: {exc}'); return
        self.refresh_accounts(); self.accounts_list.setCurrentRow(self.accounts.index(a))
        self.status.setText(f'저장됨: {a["label"]}'); self.poll()

    def delete_account(self):
        a = next((a for a in self.accounts if a['id']==self.selected_id),None)
        if not a: return
        if QMessageBox.question(self,APP,f'{a["label"]} 계정을 삭제할까요?') != QMessageBox.Yes: return
        self.accounts.remove(a)
        try: delete_unreferenced_secrets(self.accounts); self.save_config()
        except OSError as exc: QMessageBox.critical(self,APP,str(exc)); return
        self.refresh_accounts(); self.new_account()

    def reset_all(self):
        if QMessageBox.warning(self,APP,'모든 계정, 비밀번호, 확인 기록과 알림 설정을 삭제하고 자동 실행을 해제할까요?\n이 작업은 되돌릴 수 없습니다.',
                               QMessageBox.Yes | QMessageBox.No, QMessageBox.No) != QMessageBox.Yes: return
        previous = list(self.accounts)
        self.accounts = []; self.selected_id = None
        self.subjects.setRowCount(0)
        for popup in list(self.popups): popup.dismiss()
        self.autostart.setChecked(False)
        self.set_startup()
        MIGRATION_MARKER.touch()
        problems = []
        for path in (CONFIG,SECRETS):
            try: path.unlink(missing_ok=True)
            except OSError as exc: problems.append(str(exc))
        for key in {a.get('credential_key') for a in previous if a.get('credential_key')}:
            try: keyring.delete_password(SERVICE,key)
            except keyring.errors.PasswordDeleteError: pass
            except Exception: pass
        try: LEGACY_CONFIG.unlink(missing_ok=True)
        except OSError: pass
        self.mode.setCurrentIndex(0); self.duration.setValue(25); self.position.setCurrentIndex(0)
        self._mode='windows'; self._seconds=25; self._position='bottom_right'
        self.config={'accounts': [], 'notification_mode':'windows', 'notification_seconds':25, 'popup_position':'bottom_right'}
        self.refresh_accounts(); self.new_account()
        self.status.setText('모든 계정과 저장 정보를 초기화했습니다.')
        if problems: QMessageBox.warning(self,APP,'일부 파일 삭제 실패: '+'; '.join(problems))
        else: QMessageBox.information(self,APP,'모든 계정과 저장 정보를 초기화했습니다.')

    def seconds(self): return self._seconds if hasattr(self,'_seconds') else self.duration.value()

    def save_program_settings(self):
        self._mode = 'popup' if self.mode.currentIndex() == 1 else 'windows'
        self._seconds = self.duration.value()
        self._position = self.positions[self.position.currentText()]
        try:
            self.save_config(); self.set_startup()
        except OSError as exc:
            QMessageBox.critical(self,APP,f'프로그램 설정 저장 실패: {exc}'); return
        self.config.update(notification_mode=self._mode, notification_seconds=self._seconds, popup_position=self._position)
        self.position_popups(); self.status.setText('프로그램 설정을 저장했습니다.')
        QMessageBox.information(self,APP,'프로그램 설정을 저장했습니다.')

    def save_config(self):
        _write_json(CONFIG, {'accounts':self.accounts,'notification_mode':getattr(self,'_mode',self.config.get('notification_mode','windows')),
                            'notification_seconds':self.seconds(),
                            'popup_position':getattr(self,'_position',self.config.get('popup_position','bottom_right'))})

    @staticmethod
    def startup_enabled():
        try:
            with winreg.OpenKey(winreg.HKEY_CURRENT_USER,RUN_KEY) as key:
                return bool(winreg.QueryValueEx(key,RUN_NAME)[0])
        except OSError: return False

    def set_startup(self):
        with winreg.CreateKey(winreg.HKEY_CURRENT_USER,RUN_KEY) as key:
            if self.autostart.isChecked():
                if getattr(sys,'frozen',False): args=[sys.executable,'--tray']
                else: args=[sys.executable,str(Path(__file__).resolve()),'--tray']
                winreg.SetValueEx(key,RUN_NAME,0,winreg.REG_SZ,subprocess.list2cmdline(args))
            else:
                try: winreg.DeleteValue(key,RUN_NAME)
                except FileNotFoundError: pass

    def start_ipc(self):
        def server():
            try:
                with socket.socket() as sock:
                    sock.setsockopt(socket.SOL_SOCKET,socket.SO_REUSEADDR,1)
                    sock.bind(('127.0.0.1',IPC_PORT)); sock.listen(4); sock.settimeout(1)
                    while not self.closed:
                        try: conn,_=sock.accept()
                        except socket.timeout: continue
                        with conn:
                            conn.settimeout(2)
                            if conn.recv(32) == b'SHOW': self.events.put(('show',None))
            except OSError: pass
        threading.Thread(target=server,daemon=True).start()

    def register_protocol(self):
        try:
            path=rf'Software\Classes\{URI_SCHEME}'
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER,path) as key:
                winreg.SetValueEx(key,'',0,winreg.REG_SZ,APP)
                winreg.SetValueEx(key,'URL Protocol',0,winreg.REG_SZ,'')
            if getattr(sys,'frozen',False): args=[sys.executable,'%1']
            else: args=[sys.executable,str(Path(__file__).resolve()),'%1']
            with winreg.CreateKey(winreg.HKEY_CURRENT_USER,path+r'\shell\open\command') as key:
                winreg.SetValueEx(key,'',0,winreg.REG_SZ,subprocess.list2cmdline(args))
        except OSError as exc: self.status.setText(f'알림 클릭 연결 등록 실패: {exc}')

    def setup_tray(self):
        tray=QSystemTrayIcon(QIcon(str(ICON_PATH)),self)
        menu=QMenu()
        def add(label,callback):
            item=QAction(label,menu); item.triggered.connect(callback); menu.addAction(item)
        add('받은 제목 보기',self.show_window)
        add('지금 확인',self.poll)
        menu.addSeparator()
        for key in ('daum','naver','gmail'):
            label,url=MAIL_SITES[key]; add(label,lambda checked=False,u=url: webbrowser.open(u))
        menu.addSeparator(); add('종료',self.quit_app)
        tray.setContextMenu(menu)
        tray.activated.connect(lambda reason: self.show_window() if reason == QSystemTrayIcon.DoubleClick else None)
        tray.show(); self.tray = tray; self.tray_menu = menu

    def closeEvent(self,event):
        if not self.closed:
            event.ignore(); self.hide()
        else: event.accept()

    def show_window(self):
        self.showNormal(); self.raise_(); self.activateWindow()

    def periodic(self):
        if self.closed or self.busy: return
        if any(time.monotonic() >= a.get('last_check',0) + int(a.get('interval',60)) for a in self.accounts):
            self.poll()

    def poll(self):
        if self.closed or self.busy or not self.accounts: return
        due=[dict(a) for a in self.accounts if time.monotonic() >= a.get('last_check',0) + int(a.get('interval',60))]
        if not due: due=[dict(a) for a in self.accounts]
        now=time.monotonic()
        for a in self.accounts:
            if any(x['id']==a['id'] for x in due): a['last_check']=now
        self.busy=True; self.status.setText(f'{len(due)}개 계정 확인 중…'); self.state.setText('●  확인 중')
        def worker():
            for a in due:
                try:
                    password=load_secret(a['credential_key'])
                    if not password: raise RuntimeError('저장된 비밀번호가 없습니다.')
                    fetch=get_imap if a['type']=='IMAP' else get_pop3
                    state,messages,count=fetch(a,password,a.get('mail_state',{}))
                    self.events.put(('result',(a['id'],a['credential_key'],state,len(messages))))
                    for msg in messages:
                        self.events.put(('message',(a['id'],a['credential_key'],msg)))
                except Exception as exc: self.events.put(('error',(a['label'],str(exc))))
            self.events.put(('done',None))
        threading.Thread(target=worker,daemon=True).start()

    def process_events(self):
        # Bound work per timer tick to preserve responsive painting and input.
        # Large mail batches are processed over several 200-ms ticks.
        for _ in range(8):
            try:
                kind, data = self.events.get_nowait()
            except queue.Empty:
                break
            if kind == 'show':
                self.show_window()
            elif kind == 'poll':
                self.poll()
            elif kind == 'quit':
                self.quit_app()
                return
            elif kind == 'result':
                aid, key, state, count = data
                account = next((a for a in self.accounts if a['id'] == aid and a['credential_key'] == key), None)
                if account is None:
                    continue
                account['mail_state'] = state
                self._mail_state_dirty = True
                self.status.setText(f'{account["label"]}: 새 메일 {count}개')
                self.state.setText('●  연결됨')
            elif kind == 'message':
                aid, key, msg = data
                account = next((a for a in self.accounts if a['id'] == aid and a['credential_key'] == key), None)
                if account is None:
                    continue
                self.subjects.insertRow(0)
                for col, value in enumerate((short_received(msg['received']), msg['sender'], msg['subject'])):
                    self.subjects.setItem(0, col, QTableWidgetItem(value))
                self.subjects.item(0, 0).setData(Qt.UserRole, mailbox_url(account))
                if self.subjects.rowCount() > 150:
                    self.subjects.removeRow(self.subjects.rowCount() - 1)
                self.notify(account, msg)
            elif kind == 'error':
                self.status.setText(f'{data[0]} 연결 오류: {data[1]}')
                self.state.setText('●  연결 오류')
            elif kind == 'toast_error':
                self.status.setText(f'Windows 알림 오류: {data}')
            elif kind == 'done':
                self.busy = False
                if self._mail_state_dirty:
                    self._mail_state_dirty = False
                    try:
                        self.save_config()
                    except OSError as exc:
                        self.status.setText(f'설정 저장 오류: {exc}')

    def _toast_worker(self):
        # One serial worker prevents the OS notification API from blocking Qt.
        while True:
            job = self.toast_events.get()
            if job is None:
                return
            try:
                title, body, duration, url = job
                toast = Notification(app_id=APP, title=title, msg=body,
                                     duration=duration, launch=url)
                toast.set_audio(audio.Mail, loop=False)
                toast.show()
            except Exception as exc:
                self.events.put(('toast_error', str(exc)))
            finally:
                self.toast_events.task_done()

    def open_selected_mail(self,row,col):
        item=self.subjects.item(row,0)
        url=item.data(Qt.UserRole) if item else None
        if url: webbrowser.open(url)

    def notify(self, account, msg):
        url=mailbox_url(account)
        if getattr(self,'_mode',self.config.get('notification_mode','windows'))=='popup':
            popup=Popup(self,account['label'],msg,url); self.popups.append(popup)
            self.position_popups(); popup.show(); return
        try:
            self.toast_events.put_nowait((
                f'[{account["label"]}] {msg["subject"]}',
                f'발신자: {msg["sender"]}  ·  수신: {short_received(msg["received"])}\n{msg["preview"]}',
                'long' if self.seconds() > 10 else 'short',
                url or f'{URI_SCHEME}://show',
            ))
        except queue.Full:
            self.status.setText('알림이 너무 많아 일부 Windows 알림을 생략했습니다.')

    def position_popups(self):
        screen=QApplication.primaryScreen()
        if not screen: return
        area=screen.availableGeometry(); margin=16; offset=0
        pos=getattr(self,'_position',self.config.get('popup_position','bottom_right'))
        for popup in reversed(self.popups):
            popup.adjustSize(); width=popup.width(); height=popup.height()
            x=area.right()-margin-width if pos.endswith('right') else area.left()+margin
            y=area.bottom()-margin-height-offset if pos.startswith('bottom') else area.top()+margin+offset
            popup.move(max(area.left(),x),max(area.top(),y))
            offset += height+8

    def quit_app(self):
        try:
            self.toast_events.put_nowait(None)
        except queue.Full:
            pass
        self.closed=True
        self.tray.hide()
        QApplication.instance().quit()


def forward_to_running(show=True):
    try:
        with socket.create_connection(('127.0.0.1',IPC_PORT),timeout=.5) as sock:
            sock.sendall(b'SHOW' if show else b'PING')
        return True
    except OSError: return False

if __name__=='__main__':
    if forward_to_running(show=not START_IN_TRAY): sys.exit(0)
    app=QApplication(sys.argv)
    app.setQuitOnLastWindowClosed(False)
    app.setWindowIcon(QIcon(str(ICON_PATH)))
    app.setStyleSheet(STYLE)
    window=Window(); window.register_protocol()
    sys.exit(app.exec())
