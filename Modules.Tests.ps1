#test update available module
Get-Module Microsoft.Graph.Authentication -ListAvailable | Remove-Module -Force
Get-Module Microsoft.Graph.Authentication -ListAvailable | Uninstall-Module -Force
Get-Module Microsoft.Graph.Authentication -ListAvailable
Install-Module  Microsoft.Graph.Authentication -RequiredVersion 2.25.0


#test multiple modules
Install-Module Az.Accounts -RequiredVersion 4.0.0
Install-Module Az.Accounts

#