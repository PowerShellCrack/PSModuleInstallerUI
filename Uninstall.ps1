# This script is intended to be used with the Intune/ConfigMgr
Param(
    $TagName,
    $TagVersion,
    $TagDetectionPath
)

$TagPath = "$TagDetectionPath\$TagName.$TagVersion.tag"

if(Test-Path $TagPath) {Remove-Item $TagPath -Force -ErrorAction SilentlyContinue | Out-Null}

if(Test-Path $TagPath) {
    exit 1
}else{
    #no file found
    Write-Host "Detection Found"
}