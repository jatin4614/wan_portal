# ============================================================
#  WAN Portal - IIS Setup   (run as Administrator ON THE SERVER)
#  Rewritten + verified. Reproduced end-to-end and fixed the 404.
# ============================================================
#  What was actually wrong (root cause of the HTTP 404):
#    1. There was NO web.config, so IIS never launched the Python/Waitress
#       process -- every dynamic URL under /wanportal fell through to the IIS
#       static-file handler and returned 404. (Earlier notes claimed web.config
#       "travels with the app"; it does not. This script now GENERATES it.)
#    2. app.py had NO prefix handling, so even once launched Flask 404'd because
#       HttpPlatformHandler forwards the FULL path (/wanportal/login) while the
#       routes are defined at the root (/login). app.py now ships a
#       PrefixMiddleware driven by the URL_PREFIX environment variable.
#    3. serve.py hard-coded port 5000; HttpPlatformHandler assigns the port via
#       HTTP_PLATFORM_PORT. serve.py now binds to that port.
#    4. ~47 front-end fetch()/asset URLs were hard-coded to the site root; the
#       templates now prepend the mount prefix so the dashboard, map, AJAX and
#       downloads work under /wanportal.
#
#  PREREQUISITES (do these first):
#    1. Copy the CONTENTS of the inner app folder (the one containing app.py)
#       to  C:\inetpub\wwwroot\wan_portal   (so app.py is at the root there).
#       Do NOT copy the old 'venv' folder (it points at another PC's Python).
#       Copy everything else, including serve.py, requirements.txt, wheels\,
#       static\, templates\, instance\ and web.config.
#    2. HttpPlatformHandler must be installed on the server (HSMOS already uses it).
#    3. Python 3.12.x must be installed (match HSMOS).
#  Then run this script from an elevated PowerShell.
# ============================================================
#Requires -RunAsAdministrator
$ErrorActionPreference = "Stop"

# ---------- Settings (edit only if your paths/names differ) ----------
$appPath     = "C:\inetpub\wwwroot\wan_portal"   # physical folder; app.py must be at its root
$appAlias    = "wanportal"                        # URL -> https://<server>/wanportal/
$appPoolName = "wanportal_Pool"
$siteName    = "comn report"                      # parent IIS website (same as HSMOS)
$uploadCapMB = 16384                              # app upload cap in MB (0 = unlimited)
# --- derived ---
$urlPrefix   = "/$appAlias"                        # MUST match the IIS alias; app strips this
$venvPath    = "$appPath\venv"
$venvPy      = "$venvPath\Scripts\python.exe"
# ---------------------------------------------------------------------

function Section($n,$t){ Write-Host "`n[$n] $t" -ForegroundColor Yellow }
Write-Host "`n==== WAN Portal IIS Setup (rewritten) ====" -ForegroundColor Cyan

# ---- 1. Verify the app was deployed with the flattened layout ----
Section "1/10" "Verifying deployed layout..."
if (-not (Test-Path "$appPath\app.py")) {
    Write-Host "  ERROR: $appPath\app.py not found." -ForegroundColor Red
    Write-Host "  Copy the CONTENTS of the inner wan_portal folder (the one with app.py) into $appPath." -ForegroundColor Red
    exit 1
}
if (Test-Path "$appPath\wan_portal\app.py") {
    Write-Host "  ERROR: nested folder detected ($appPath\wan_portal\app.py)." -ForegroundColor Red
    Write-Host "  Move the inner folder's CONTENTS up one level so app.py is directly in $appPath." -ForegroundColor Red
    exit 1
}
if (-not (Test-Path "$appPath\serve.py"))        { Write-Host "  ERROR: serve.py missing." -ForegroundColor Red; exit 1 }
if (-not (Test-Path "$appPath\requirements.txt")){ Write-Host "  ERROR: requirements.txt missing." -ForegroundColor Red; exit 1 }
if (-not (Test-Path "$appPath\wheels"))          { Write-Host "  ERROR: offline 'wheels' folder missing (needed for offline install)." -ForegroundColor Red; exit 1 }
Write-Host "  Layout OK (app.py, serve.py, requirements.txt, wheels present)" -ForegroundColor Green

# ---- 2. Locate the server's Python 3.12.x ----
Section "2/10" "Locating Python 3.12 on the server..."
function Find-ServerPython {
    try { $p = & py -3.12 -c "import sys;print(sys.executable)" 2>$null; if ($LASTEXITCODE -eq 0 -and $p) { return $p.Trim() } } catch {}
    $cands = @(
        "$env:LOCALAPPDATA\Programs\Python\Python312\python.exe",
        "C:\Program Files\Python312\python.exe",
        "C:\Python312\python.exe"
    )
    $hsmosCfg = "C:\inetpub\wwwroot\HSMOS\.venv\pyvenv.cfg"   # reuse the Python HSMOS runs on
    if (Test-Path $hsmosCfg) {
        $h = Get-Content $hsmosCfg | Where-Object { $_ -match '^home\s*=' }
        if ($h) { $cands += ((($h -split '=',2)[1]).Trim() + "\python.exe") }
    }
    foreach ($c in $cands) { if (Test-Path $c) { return $c } }
    return $null
}
$serverPy = Find-ServerPython
if (-not $serverPy) {
    Write-Host "  ERROR: could not find Python 3.12 on the server." -ForegroundColor Red
    Write-Host "  Install Python 3.12.x (match HSMOS) or set `$serverPy manually, then re-run." -ForegroundColor Red
    exit 1
}
Write-Host "  Using server Python: $serverPy ($(& $serverPy --version 2>&1))" -ForegroundColor Green

# ---- 3. Remove any stale/copied venv ----
Section "3/10" "Removing any stale venv..."
if (Test-Path $venvPath) { Remove-Item -Recurse -Force $venvPath; Write-Host "  Deleted old $venvPath" -ForegroundColor Green }
else { Write-Host "  (none present)" -ForegroundColor Green }

# ---- 4. Create a fresh venv with the server's Python ----
Section "4/10" "Creating fresh venv..."
& $serverPy -m venv $venvPath
if (-not (Test-Path $venvPy)) { Write-Host "  ERROR: venv creation failed." -ForegroundColor Red; exit 1 }
Write-Host "  Created $venvPath ($(& $venvPy --version 2>&1))" -ForegroundColor Green

# ---- 5. Install dependencies OFFLINE from the bundled wheelhouse ----
Section "5/10" "Installing dependencies offline (incl. waitress)..."
& $venvPy -m pip install --no-index --find-links "$appPath\wheels" -r "$appPath\requirements.txt"
if ($LASTEXITCODE -ne 0) { Write-Host "  ERROR: offline pip install failed (check the wheels folder matches Python 3.12)." -ForegroundColor Red; exit 1 }
& $venvPy -c "import waitress, flask, pandas, openpyxl; print('  import check OK')"
if ($LASTEXITCODE -ne 0) { Write-Host "  ERROR: dependency import check failed." -ForegroundColor Red; exit 1 }
Write-Host "  Dependencies installed and importable" -ForegroundColor Green

# ---- 6. Generate web.config with ABSOLUTE paths + the correct URL_PREFIX ----
Section "6/10" "Writing web.config (absolute paths, URL_PREFIX=$urlPrefix)..."
$webConfig = @"
<?xml version="1.0" encoding="utf-8"?>
<!-- Generated by setup_wanportal_iis.ps1. Do not hand-edit; re-run the script. -->
<configuration>
  <system.webServer>
    <handlers>
      <remove name="httpPlatformHandler" />
      <add name="httpPlatformHandler" path="*" verb="*"
           modules="httpPlatformHandler" resourceType="Unspecified" />
    </handlers>
    <httpPlatform stdoutLogEnabled="true"
                  stdoutLogFile="$appPath\logs\stdout"
                  startupTimeLimit="120"
                  startupRetryCount="3"
                  requestTimeout="00:10:00"
                  processPath="$venvPy"
                  arguments="$appPath\serve.py">
      <environmentVariables>
        <environmentVariable name="URL_PREFIX"       value="$urlPrefix" />
        <environmentVariable name="MAX_UPLOAD_MB"    value="$uploadCapMB" />
        <environmentVariable name="PYTHONUNBUFFERED" value="1" />
      </environmentVariables>
    </httpPlatform>
    <security>
      <requestFiltering>
        <requestLimits maxAllowedContentLength="4294967295" />
      </requestFiltering>
    </security>
    <defaultDocument enabled="false" />
    <directoryBrowse enabled="false" />
  </system.webServer>
</configuration>
"@
Set-Content -Path "$appPath\web.config" -Value $webConfig -Encoding UTF8
Write-Host "  Wrote $appPath\web.config" -ForegroundColor Green

# ---- 7. Logs + filesystem permissions ----
Section "7/10" "Logs and permissions..."
New-Item -ItemType Directory -Force "$appPath\logs" | Out-Null
New-Item -ItemType Directory -Force "$appPath\instance" | Out-Null
icacls $appPath /grant "IIS_IUSRS:(OI)(CI)RX" /T | Out-Null
icacls "$appPath\instance" /grant "IIS_IUSRS:(OI)(CI)M" /T | Out-Null
icacls "$appPath\logs" /grant "IIS_IUSRS:(OI)(CI)M" /T | Out-Null
# IIS_IUSRS must be able to reach the base Python the venv points to
$cfg = "$venvPath\pyvenv.cfg"
if (Test-Path $cfg) {
    $homeLine = Get-Content $cfg | Where-Object { $_ -match '^home\s*=' }
    if ($homeLine) {
        $pyHome = (($homeLine -split '=',2)[1]).Trim()
        if (Test-Path $pyHome) { try { icacls $pyHome /grant "IIS_IUSRS:(OI)(CI)RX" /T | Out-Null } catch {} }
    }
}
Write-Host "  logs+instance ready; IIS_IUSRS granted RX on app, Modify on instance+logs" -ForegroundColor Green

# ---- 8. IIS module checks + unlock the httpPlatform section ----
Section "8/10" "Wiring up IIS (module + config section)..."
Import-Module WebAdministration -ErrorAction Stop
$appcmd = "$env:windir\system32\inetsrv\appcmd.exe"
# HttpPlatformHandler must be installed
$modOk = (& $appcmd list module /name:httpPlatformHandler) 2>$null
if (-not $modOk) {
    Write-Host "  ERROR: HttpPlatformHandler module is not installed on this server." -ForegroundColor Red
    Write-Host "  Install it (same as HSMOS uses) and re-run." -ForegroundColor Red
    exit 1
}
# The <httpPlatform> section is locked by default; unlock so web.config may set it.
& $appcmd unlock config /section:system.webServer/httpPlatform 2>$null | Out-Null
if (-not (Get-Website -Name $siteName -ErrorAction SilentlyContinue)) {
    Write-Host "  ERROR: parent site '$siteName' not found. Sites:" -ForegroundColor Red
    Get-Website | Format-Table Name,ID,State -AutoSize
    exit 1
}
Write-Host "  HttpPlatformHandler present; httpPlatform section unlocked; parent site OK" -ForegroundColor Green

# ---- 9. App pool + application ----
Section "9/10" "App pool '$appPoolName' + application /$appAlias..."
if (-not (Test-Path "IIS:\AppPools\$appPoolName")) { New-WebAppPool -Name $appPoolName | Out-Null }
Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name managedRuntimeVersion -Value ""          # No Managed Code
Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name managedPipelineMode  -Value "Integrated"
Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name startMode            -Value "AlwaysRunning"
Set-ItemProperty "IIS:\AppPools\$appPoolName" -Name processModel.idleTimeout -Value ([TimeSpan]::FromMinutes(0))
$existing = Get-WebApplication -Site $siteName -Name $appAlias -ErrorAction SilentlyContinue
if ($existing) { Remove-WebApplication -Site $siteName -Name $appAlias }
New-WebApplication -Site $siteName -Name $appAlias -PhysicalPath $appPath -ApplicationPool $appPoolName | Out-Null
Write-Host "  /$appAlias created under '$siteName' (pool $appPoolName)" -ForegroundColor Green

# ---- 10. Start + test ----
Section "10/10" "Recycling and testing..."
Restart-WebAppPool -Name $appPoolName
Start-Sleep -Seconds 5
Write-Host "  Pool recycled." -ForegroundColor Green
Write-Host "`n  If anything is off, tail the startup log:" -ForegroundColor White
Write-Host "    Get-Content $appPath\logs\stdout* -Tail 40 -Wait" -ForegroundColor Gray
Write-Host "  Then open: https://<server>/$appAlias/   (default login: admin / admin)" -ForegroundColor White
try {
    if ($PSVersionTable.PSVersion.Major -ge 6) {
        $r = Invoke-WebRequest "https://localhost/$appAlias/login" -UseBasicParsing -TimeoutSec 30 -SkipCertificateCheck
    } else {
        add-type @"
using System.Net; using System.Security.Cryptography.X509Certificates;
public class _TA : ICertificatePolicy { public bool CheckValidationResult(ServicePoint a,X509Certificate b,WebRequest c,int d){return true;} }
"@
        [System.Net.ServicePointManager]::CertificatePolicy = New-Object _TA
        $r = Invoke-WebRequest "https://localhost/$appAlias/login" -UseBasicParsing -TimeoutSec 30
    }
    if ($r.StatusCode -eq 200) {
        Write-Host "`n  SUCCESS: https://localhost/$appAlias/login returned HTTP 200" -ForegroundColor Green
    } else {
        Write-Host "`n  https://localhost/$appAlias/login returned HTTP $($r.StatusCode)" -ForegroundColor Yellow
    }
} catch {
    Write-Host "`n  Could not confirm via localhost (a cert warning here can be normal)." -ForegroundColor Yellow
    Write-Host "  Check the log: Get-Content $appPath\logs\stdout* -Tail 40" -ForegroundColor Yellow
    Write-Host "  ($($_.Exception.Message))" -ForegroundColor DarkYellow
}
Write-Host ""
