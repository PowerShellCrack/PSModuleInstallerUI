# Description: This script is used to detect if a tag is present on a resource
$TagName = "PowerShellModuleInstalle"
$TagVersion = "2.3.0"
$TagPath = "$env:ALLUSERSPROFILE\$TagName.$TagVersion.tag"

if(Test-Path $TagPath) {
    Write-Host "Detection Found"
}else{
    #no detection found
    exit 1
}