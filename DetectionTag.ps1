# Description: This script is used to detect if a tag is present on a resource
$TagName = "PowerShellModuleInstaller"
$TagVersion = "2.3.0"
$TagPath = "C:\ProgramData\Company\$TagName.$TagVersion.tag"

if(Test-Path $TagPath) {
    Write-Host "Detection Found"
}else{
    #no detection found
    exit 1
}