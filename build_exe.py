"""Build the Windows one-file executable from the existing app."""
from pathlib import Path
import subprocess
import sys

root = Path(__file__).resolve().parent
source = root / 'mail_title_notifier.py'
icon = root / 'mail_icon.ico'
for needed in (source, icon, root / 'requirements.txt'):
    if not needed.is_file():
        raise SystemExit(f'Missing: {needed}')

cmd = [sys.executable, '-m', 'PyInstaller', '--noconfirm', '--clean', '--onefile', '--windowed',
       '--icon', str(icon), '--add-data', f'{icon}{__import__("os").pathsep}.', '--name', '메일 알리미', str(source)]
print('Building:', source.name, flush=True)
subprocess.run(cmd, cwd=root, check=True)
result = root / 'dist' / '메일 알리미.exe'
if not result.is_file():
    raise SystemExit(f'Build returned success but EXE not found: {result}')
print(f'BUILD SUCCEEDED: {result}', flush=True)
