# WinPython portable base zip (small) from GitHub releases.
# You can override with: powershell -ExecutionPolicy Bypass -File .\build_windows_exe.ps1 -WinPythonUrl "<url>"
param(
  [string]$WinPythonUrl = "https://github.com/winpython/winpython/releases/download/17.2.20260225/WinPython64-3.13.12.0dotb3.zip"
)

$ErrorActionPreference = "Stop"

function Write-Log($msg) {
  $ts = Get-Date -Format "yyyy-MM-dd HH:mm:ss"
  Write-Host "[$ts] $msg"
}

$Root = Split-Path -Parent $MyInvocation.MyCommand.Path
Set-Location $Root

$LogPath = Join-Path $Root "build_windows_exe_ps.log"
"[build] started at $(Get-Date)" | Out-File -Encoding UTF8 $LogPath

function Run($cmd, $args) {
  throw "Internal error: old Run() signature still in file"
}

function Run($cmd, $argList) {
  $argText = if ($null -eq $argList) { "" } else { [string]$argList }
  Write-Log "$cmd $argText"
  "`n> $cmd $argText" | Add-Content -Encoding UTF8 $LogPath

  $tmpOut = Join-Path $Root ("._tmp_out_" + ([guid]::NewGuid().ToString()) + ".log")
  $tmpErr = Join-Path $Root ("._tmp_err_" + ([guid]::NewGuid().ToString()) + ".log")
  New-Item -ItemType File -Path $tmpOut -Force | Out-Null
  New-Item -ItemType File -Path $tmpErr -Force | Out-Null

  try {
    if ($null -eq $argList -or $argText.Trim().Length -eq 0) {
      $p = Start-Process -FilePath $cmd -NoNewWindow -PassThru -Wait -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr
    } else {
      $p = Start-Process -FilePath $cmd -ArgumentList $argList -NoNewWindow -PassThru -Wait -RedirectStandardOutput $tmpOut -RedirectStandardError $tmpErr
    }
  } finally {
    if (Test-Path $tmpOut) { Get-Content $tmpOut -ErrorAction SilentlyContinue | Add-Content -Encoding UTF8 $LogPath }
    if (Test-Path $tmpErr) { Get-Content $tmpErr -ErrorAction SilentlyContinue | Add-Content -Encoding UTF8 $LogPath }
    Remove-Item -Force -ErrorAction SilentlyContinue $tmpOut, $tmpErr | Out-Null
  }

  if ($p.ExitCode -ne 0) {
    throw "Command failed with exit code $($p.ExitCode). See log: $LogPath"
  }
}

if (!(Test-Path (Join-Path $Root "app_tk.py"))) {
  throw "app_tk.py not found in $Root"
}

$ToolsDir = Join-Path $Root ".tools"
$WpZip = Join-Path $ToolsDir "winpython.zip"
$WpDir = Join-Path $ToolsDir "winpython"

New-Item -ItemType Directory -Force -Path $ToolsDir | Out-Null

if (!(Test-Path $WpDir)) {
  New-Item -ItemType Directory -Force -Path $WpDir | Out-Null
}

if (!(Test-Path $WpZip)) {
  Write-Log "Downloading WinPython..."
  "Downloading: $WinPythonUrl" | Add-Content -Encoding UTF8 $LogPath
  Invoke-WebRequest -Uri $WinPythonUrl -OutFile $WpZip
}

if (!(Test-Path (Join-Path $WpDir "_extracted.ok"))) {
  Write-Log "Extracting WinPython..."
  Get-ChildItem -Path $WpDir -Force | Remove-Item -Recurse -Force -ErrorAction SilentlyContinue
  Expand-Archive -Path $WpZip -DestinationPath $WpDir -Force
  New-Item -ItemType File -Path (Join-Path $WpDir "_extracted.ok") | Out-Null
}

Write-Log "Locating python.exe..."
$pythonExe = Get-ChildItem -Path $WpDir -Recurse -Filter "python.exe" -ErrorAction SilentlyContinue |
  Where-Object { $_.FullName -notmatch "\\tcl\\|\\DLLs\\|\\Lib\\|\\include\\|\\site-packages\\" } |
  Select-Object -First 1

if (-not $pythonExe) {
  # fallback: pick first python.exe
  $pythonExe = Get-ChildItem -Path $WpDir -Recurse -Filter "python.exe" -ErrorAction SilentlyContinue | Select-Object -First 1
}

if (-not $pythonExe) {
  throw "python.exe not found inside extracted WinPython. See log: $LogPath"
}

$PY = $pythonExe.FullName
Write-Log "Using Python: $PY"
& $PY --version | Tee-Object -FilePath $LogPath -Append | Out-Null

Write-Log "Ensuring pip..."
try {
  Run $PY "-m pip --version"
} catch {
  Run $PY "-m ensurepip --upgrade"
  Run $PY "-m pip install --upgrade pip"
}

Write-Log "Installing dependencies..."
Run $PY "-m pip install -r requirements.txt"
Run $PY "-m pip install pyinstaller"

Write-Log "Building exe..."
Run $PY "-m PyInstaller --noconfirm --clean --windowed --name MailNotifier app_tk.py"

Write-Log "DONE"
Write-Host ""
Write-Host "Your exe is here:"
Write-Host "  $Root\dist\MailNotifier\MailNotifier.exe"
Write-Host ""
Write-Host "Log:"
Write-Host "  $LogPath"

