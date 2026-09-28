<#
  MXKPCS 대시보드(server.js)를 NSSM으로 Windows 서비스에 등록합니다.
  실행 전 준비:
    npm install
    npm run build
  사용법 (관리자 권한 PowerShell):
    .\deployment\install-service.ps1
    .\deployment\install-service.ps1 -Port 8080 -ServiceName MXKPCS-Dashboard -OpenFirewall
  nssm.exe가 PATH에 없다면 -NssmPath로 직접 지정하거나
  deployment\nssm.exe 에 파일을 넣어두면 자동으로 찾습니다.
#>
param(
    [string]$ServiceName = "MXKPCSServer",
    [int]$Port = 2100,
    [string]$HostAddress = "0.0.0.0",
    [string]$NssmPath = "",
    [switch]$OpenFirewall
)

$ErrorActionPreference = "Stop"

$ProjectRoot = Split-Path -Parent $PSScriptRoot
$DistDir = Join-Path $ProjectRoot "dist"
$LogDir = Join-Path $ProjectRoot "logs"
$ServerScript = Join-Path $ProjectRoot "server.js"

if (-not (Test-Path $ServerScript)) {
    throw "server.js를 찾을 수 없습니다: $ServerScript"
}
if (-not (Test-Path $DistDir)) {
    Write-Warning "dist 폴더가 없습니다. 먼저 'npm run build'를 실행하세요."
}

$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if ($existing) {
    throw "서비스 '$ServiceName'이(가) 이미 존재합니다. 먼저 uninstall-service.ps1을 실행하세요."
}

function Resolve-Nssm {
    param([string]$ExplicitPath)
    if ($ExplicitPath) {
        if (Test-Path $ExplicitPath) { return (Resolve-Path $ExplicitPath).Path }
        throw "지정한 -NssmPath에서 nssm.exe를 찾을 수 없습니다: $ExplicitPath"
    }
    $cmd = Get-Command nssm.exe -ErrorAction SilentlyContinue
    if ($cmd) { return $cmd.Source }
    $bundled = Join-Path $PSScriptRoot "nssm.exe"
    if (Test-Path $bundled) { return $bundled }
    throw "nssm.exe를 찾을 수 없습니다. https://nssm.cc/download 에서 받아 PATH에 추가하거나 -NssmPath로 경로를 지정하세요."
}

$nssm = Resolve-Nssm -ExplicitPath $NssmPath
$node = (Get-Command node.exe -ErrorAction Stop).Source

if (-not (Test-Path $LogDir)) {
    New-Item -ItemType Directory -Path $LogDir | Out-Null
}

Write-Host "node.exe        : $node"
Write-Host "nssm.exe        : $nssm"
Write-Host "프로젝트 경로   : $ProjectRoot"
Write-Host "서비스 이름     : $ServiceName"
Write-Host "리스닝 주소     : http://$($HostAddress):$Port"

& $nssm install $ServiceName $node "server.js"
& $nssm set $ServiceName AppDirectory $ProjectRoot
& $nssm set $ServiceName AppStdout (Join-Path $LogDir "service-out.log")
& $nssm set $ServiceName AppStderr (Join-Path $LogDir "service-err.log")
& $nssm set $ServiceName AppRotateFiles 1
& $nssm set $ServiceName AppRotateOnline 1
& $nssm set $ServiceName AppRotateBytes 10485760
& $nssm set $ServiceName AppEnvironmentExtra "PORT=$Port" "HOST=$HostAddress"
& $nssm set $ServiceName Start SERVICE_AUTO_START
& $nssm set $ServiceName AppExit Default Restart
& $nssm set $ServiceName AppRestartDelay 3000
& $nssm set $ServiceName DisplayName "MXKPCS XRF Dashboard"
& $nssm set $ServiceName Description "Vite build + /api/xrf-sharepoint self-hosted server (server.js)"

if ($OpenFirewall) {
    $ruleName = "$ServiceName-$Port"
    if (-not (Get-NetFirewallRule -DisplayName $ruleName -ErrorAction SilentlyContinue)) {
        New-NetFirewallRule -DisplayName $ruleName -Direction Inbound -Protocol TCP -LocalPort $Port -Action Allow | Out-Null
        Write-Host "방화벽 인바운드 규칙 추가: $ruleName (TCP $Port)"
    }
}

& $nssm start $ServiceName

Write-Host ""
Write-Host "서비스 '$ServiceName' 등록 및 시작 완료: http://localhost:$Port"
Write-Host "로그 위치: $LogDir"
