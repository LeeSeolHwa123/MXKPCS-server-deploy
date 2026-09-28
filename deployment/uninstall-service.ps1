<#
  install-service.ps1로 등록한 NSSM 서비스를 제거합니다.
  사용법 (관리자 권한 PowerShell):
    .\deployment\uninstall-service.ps1
    .\deployment\uninstall-service.ps1 -ServiceName MXKPCS-Dashboard
#>
param(
    [string]$ServiceName = "MXKPCSServer",
    [string]$NssmPath = ""
)

$ErrorActionPreference = "Stop"

$existing = Get-Service -Name $ServiceName -ErrorAction SilentlyContinue
if (-not $existing) {
    Write-Host "서비스 '$ServiceName'이(가) 존재하지 않습니다. 할 일이 없습니다."
    exit 0
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

& $nssm stop $ServiceName
& $nssm remove $ServiceName confirm

Write-Host "서비스 '$ServiceName' 제거 완료."
