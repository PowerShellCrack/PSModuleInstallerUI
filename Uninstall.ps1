# This script is intended to be used with the Intune/ConfigMgr
$TagName = "PowerShellModuleInstaller"
$TagVersion = "2.3.0"
$TagPath = "$env:ALLUSERSPROFILE\$TagName.$TagVersion.tag"

Remove-Item $TagPath -Force -ErrorAction SilentlyContinue

if(Test-Path $TagPath) {
    exit 1
}else{
    #no file found
    Write-Host "Detection Found"
}