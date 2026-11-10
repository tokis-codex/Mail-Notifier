$ErrorActionPreference = 'SilentlyContinue'
$root = $PSScriptRoot
$result = Join-Path $root '.build_python_path.txt'
Remove-Item -LiteralPath $result -Force -ErrorAction SilentlyContinue
$candidates = New-Object System.Collections.Generic.List[string]
function Add-Python([string]$p) {
    if (-not [string]::IsNullOrWhiteSpace($p)) {
        $p = $p.Trim('"').Trim()
        if (-not $candidates.Contains($p)) { [void]$candidates.Add($p) }
    }
}
if ($env:PYTHON_EXE) { Add-Python $env:PYTHON_EXE }
$launch = Get-Command py.exe -ErrorAction SilentlyContinue
if ($launch) {
    foreach ($ver in @('3.13','3.12','3.11','3')) {
        $found = & $launch.Source "-$ver" -c 'import sys; print(sys.executable)' 2>$null
        if ($LASTEXITCODE -eq 0 -and $found) { Add-Python ($found | Select-Object -First 1) }
    }
}
foreach ($cmd in @('python.exe','python3.exe')) {
  foreach ($item in @(Get-Command $cmd -All -ErrorAction SilentlyContinue)) {
    if ($item.Source) { Add-Python $item.Source }
  }
}
foreach ($base in @($env:LOCALAPPDATA, $env:ProgramFiles, ${env:ProgramFiles(x86)}, 'C:\Python311', 'C:\Python312', 'C:\Python313')) {
    if (-not $base) { continue }
    foreach ($ver in @('Python313','Python312','Python311')) {
        Add-Python (Join-Path $base "Programs\Python\$ver\python.exe")
        Add-Python (Join-Path $base "Python\$ver\python.exe")
        Add-Python (Join-Path $base "$ver\python.exe")
    }
    Add-Python (Join-Path $base 'python.exe')
}
foreach ($key in @('HKCU:\Software\Python\PythonCore','HKLM:\Software\Python\PythonCore','HKLM:\Software\WOW6432Node\Python\PythonCore')) {
    if (-not (Test-Path $key)) { continue }
    foreach ($v in @(Get-ChildItem $key)) {
        $install = Join-Path $v.PSPath 'InstallPath'
        Add-Python (Get-ItemPropertyValue -Path $install -Name ExecutablePath)
        $path = (Get-Item $install).GetValue('')
        if ($path) { Add-Python (Join-Path $path 'python.exe') }
    }
}
foreach ($p in $candidates) {
    if (-not (Test-Path -LiteralPath $p -PathType Leaf)) { continue }
    & $p -c 'import sys;sys.exit(0 if (3,11)<=sys.version_info[:2]<(3,14) else 1)' 2>$null
    if ($LASTEXITCODE -eq 0) {
        [System.IO.File]::WriteAllText($result,$p,[System.Text.Encoding]::ASCII)
        Write-Host "Found supported Python: $p"
        exit 0
    }
}
Write-Host 'No supported Python 3.11 - 3.13 discovered for this Windows user.'
Write-Host 'Try running: py -0p  or  where python  or  python --version'
Write-Host 'If Python is installed elsewhere, set PYTHON_EXE to its full python.exe path.'
exit 1
