# This is lightly revised version of the file UpdateAzureAutomationGraphModules.PS1
# Original github: https://github.com/12Knocksinna/Office365itpros/blob/master/Update%20AzureAutomationGraphModules.PS1
# A script to update the set of Graph modules for an Azure Automation account
# Revision made 17-May-2024 - Now includes updating all modules on both 5.1 and 7.2 powershell.
# Requires the Az.Automation PowerShell module

Write-Output "Connecting to Azure Automation"
# If your account uses MFA, as it should... you need to authenticate by passing the tenantid and subscriptionid 
# see https://learn.microsoft.com/en-us/powershell/module/az.accounts/connect-azaccount?view=azps-11.3.0 for more info

$SubscriptionId = "" ##Fill in the subscription ID connected to your Azure tenant
$TenantId = "" #Fill in the ID for your Microsoft tenant
$Status = Connect-AzAccount -TenantId $TenantId -SubscriptionId $SubscriptionId
If (!($Status)) { 
  Write-Output "Account not authenticated - exiting" ; break 
}

# Find Latest version from PowerShell Gallery
$DesiredVersion = (Find-Module -Name Microsoft.Graph | Select-Object -ExpandProperty Version)
If ($DesiredVersion -isnot [string]) { # Handle PowerShell 5 - PowerShell 7 returns a string
   $DesiredVersion = $DesiredVersion.Major.toString() + "." + $DesiredVersion.Minor.toString() + "." + $DesiredVersion.Build.toString()
}
Write-Output ("Checking for version {0} of the Microsoft.Graph PowerShell module" -f $DesiredVersion)
# Process Exchange Online also...
$DesiredExoVersion = (Find-Module -Name ExchangeOnlineManagement | Select-Object -ExpandProperty Version)
If ($DesiredExoVersion -isnot [string]) { # Handle PowerShell 5 - PowerShell 7 returns a string
    $DesiredExoVersion = $DesiredExoVersion.Major.toString() + "." + $DesiredExoVersion.Minor.toString() + "." + $DesiredExoVersion.Build.toString()
}
Write-Output "Checking for version $DesiredExoVersion of the Exchange Online Management module"

[Array]$AzAccounts = Get-AzAutomationAccount
If (!($AzAccounts)) { Write-Output "No Automation accounts found - existing" ; break }
Write-Output "$($AzAccounts.Count) Azure Automation accounts will be processed"

## \\ This is for runtime version 7.2
ForEach ($AzAccount in $AzAccounts) {
  $AzName = $AzAccount.AutomationAccountName
  $AzResourceGroup = $AzAccount.ResourceGroupName
  Write-Output "Checking Microsoft Graph Modules in Account $AzName for Powershell runtime version 7.2"

  [array]$GraphPSModules = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -RuntimeVersion 7.2 |  Where-Object {$_.Name -match "Microsoft.Graph"}
  If ($GraphPSModules.count -gt 0) {
    Write-Output ""
    Write-Output "Current Status"
    Write-Output "--------------"
    $GraphPSModules | Format-Table Name, Version, LastModifiedTime }
  
  $UpgradeNeeded = $True
  $ModulesToUpdate = $GraphPSModules | Where-Object {$_.Version -ne $DesiredVersion}
  $ModulesToUpdate = $ModulesToUpdate | Sort-Object Name
  If ($ModulesToUpdate.Count -eq 0) {
     Write-Output "No modules need to be updated for account $AzName"
     Write-Output ""
     $UpgradeNeeded = $False
  } Else {
    Write-Output ""
    Write-Output "Modules that need to be updated to $DesiredVersion"
    Write-Output ""
    $ModulesToUpdate | Format-Table Name, Version, LastModifiedTime
    Write-Output "Removing old modules..."
    ForEach ($Module in $ModulesToUpdate) {
       $ModuleName = $Module.Name
       Write-Output "Uninstalling module $ModuleName from Az Account $AzName"
       Remove-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -RuntimeVersion 7.2 -Confirm:$False -Force }
   }

# Check if Modules to be updated contain Microsoft.Graph.Authentication. It should be done first to avoid dependency issues
 If ($ModulesToUpdate.Name -contains "Microsoft.Graph.Authentication" -and $UpgradeNeeded -eq $True) { 
   Write-Output ""
   Write-Output "Updating Microsoft Graph Authentication module first"
   $ModuleName = "Microsoft.Graph.Authentication"
   $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredVersion"
   $Status = New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri -RuntimeVersion 7.2
   Start-Sleep -Seconds 180 
   # Remove authentication from the set of modules for update
   [array]$ModulesToUpdate = $ModulesToUpdate | Where-Object {$_.Name -ne "Microsoft.Graph.Authentication"}
 }

# Only process remaining modules if there are any to update
If ($ModulesToUpdate.Count -gt 0 -and $UpgradeNeeded -eq $True) {
  Write-Output "Adding new version of modules..."
  ForEach ($Module in $ModulesToUpdate) { 
    [string]$ModuleName = $Module.Name
    $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredVersion"
    Write-Output "Updating module $ModuleName from $Uri"
    $Status = (New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri -RuntimeVersion 7.2)
  } #End ForEach
  Write-Output "Waiting for module import processing to complete..."
  # Wait for to let everything finish
  [int]$x = 0
  Do  {
    Start-Sleep -Seconds 60
    # Check that all the modules we're interested in are fully provisioned with updated code
    [array]$GraphPSModules = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -RuntimeVersion 7.2 | `
       Where-Object {$_.Name -match "Microsoft.Graph" -and $_.ProvisioningState -eq "Succeeded"}
    [array]$ModulesToUpdate = $GraphPSModules | Where-Object {$_.Version -ne $DesiredVersion}
    If ($ModulesToUpdate.Count -eq 0) {
      $x = 1
    } Else {
      Write-Output "Still working..." 
    }
  } While ($x = 0)

  Write-Output ""
  Write-Output "Microsoft Graph modules are now upgraded to version $DesiredVersion for AZ account $AzName"
  Write-Output ""
  $GraphPSModules | Format-Table Name, Version, LastModifiedTime
 } # End If Modules

 # Check for updates to the Exchange Online Management module
 Write-Output "Checking Exchange Online Management module in Account $AzName"
 [array]$ExoPSModule = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -RuntimeVersion 7.2 | Where-Object {$_.Name -match "ExchangeOnlineManagement" }
  If ($ExoPSModule) {
    Write-Output ""
    Write-Output "Current Status"
    Write-Output "--------------"
    $ExoPSModule | Format-Table Name, Version, LastModifiedTime }
  
  $UpgradeNeeded = $True
  [array]$ModulesToUpdate = $ExoPSModule | Where-Object {$_.Version -ne $DesiredExoVersion}
  If (!($ModulesToUpdate)) {
     Write-Output "The Exchange Online Management module does not need to be updated for account $AzName"
     Write-Output ""
     $UpgradeNeeded = $False
  } Else {
    [string]$ModuleName = "ExchangeOnlineManagement"
    Write-Output ""
    Write-Output "Updating the Exchange Online management module to version $DesiredExoVersion"
    Write-Output "Removing old module..."
    Write-Output "Uninstalling module $ModuleName from Az Account $AzName"
    Remove-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -RuntimeVersion 7.2 -Confirm:$False -Force 
    $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredExoVersion"
    Write-Output "Updating module $ModuleName from $Uri"
    $Status = (New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri -RuntimeVersion 7.2)
   }

} #End ForEach Az Account
Write-Output "All done. The modules in your Azure Automation accounts are now up to date for Powershell Runtime version 7.2"

## \\ This is for runtime version 5.1
ForEach ($AzAccount in $AzAccounts) {
  $AzName = $AzAccount.AutomationAccountName
  $AzResourceGroup = $AzAccount.ResourceGroupName
  Write-Output "Checking Microsoft Graph Modules in Account $AzName for Powershell runtime version 5.1"

  [array]$GraphPSModules = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -RuntimeVersion 5.1 |  Where-Object {$_.Name -match "Microsoft.Graph"}
  If ($GraphPSModules.count -gt 0) {
    Write-Output ""
    Write-Output "Current Status"
    Write-Output "--------------"
    $GraphPSModules | Format-Table Name, Version, LastModifiedTime }
  
  $UpgradeNeeded = $True
  $ModulesToUpdate = $GraphPSModules | Where-Object {$_.Version -ne $DesiredVersion}
  $ModulesToUpdate = $ModulesToUpdate | Sort-Object Name
  If ($ModulesToUpdate.Count -eq 0) {
     Write-Output "No modules need to be updated for account $AzName"
     Write-Output ""
     $UpgradeNeeded = $False
  } Else {
    Write-Output ""
    Write-Output "Modules that need to be updated to $DesiredVersion"
    Write-Output ""
    $ModulesToUpdate | Format-Table Name, Version, LastModifiedTime
    Write-Output "Removing old modules..."
    ForEach ($Module in $ModulesToUpdate) {
       $ModuleName = $Module.Name
       Write-Output "Uninstalling module $ModuleName from Az Account $AzName"
       Remove-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -Confirm:$False -Force }
   }

# Check if Modules to be updated contain Microsoft.Graph.Authentication. It should be done first to avoid dependency issues
 If ($ModulesToUpdate.Name -contains "Microsoft.Graph.Authentication" -and $UpgradeNeeded -eq $True) { 
   Write-Output ""
   Write-Output "Updating Microsoft Graph Authentication module first"
   $ModuleName = "Microsoft.Graph.Authentication"
   $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredVersion"
   $Status = New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri 
   Start-Sleep -Seconds 180 
   # Remove authentication from the set of modules for update
   [array]$ModulesToUpdate = $ModulesToUpdate | Where-Object {$_.Name -ne "Microsoft.Graph.Authentication"}
 }

# Only process remaining modules if there are any to update
If ($ModulesToUpdate.Count -gt 0 -and $UpgradeNeeded -eq $True) {
  Write-Output "Adding new version of modules..."
  ForEach ($Module in $ModulesToUpdate) { 
    [string]$ModuleName = $Module.Name
    $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredVersion"
    Write-Output "Updating module $ModuleName from $Uri"
    $Status = (New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri)
  } #End ForEach
  Write-Output "Waiting for module import processing to complete..."
  # Wait for to let everything finish
  [int]$x = 0
  Do  {
    Start-Sleep -Seconds 60
    # Check that all the modules we're interested in are fully provisioned with updated code
    [array]$GraphPSModules = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup | `
       Where-Object {$_.Name -match "Microsoft.Graph" -and $_.ProvisioningState -eq "Succeeded"}
    [array]$ModulesToUpdate = $GraphPSModules | Where-Object {$_.Version -ne $DesiredVersion}
    If ($ModulesToUpdate.Count -eq 0) {
      $x = 1
    } Else {
      Write-Output "Still working..." 
    }
  } While ($x = 0)

  Write-Output ""
  Write-Output "Microsoft Graph modules are now upgraded to version $DesiredVersion for AZ account $AzName"
  Write-Output ""
  $GraphPSModules | Format-Table Name, Version, LastModifiedTime
 } # End If Modules

 # Check for updates to the Exchange Online Management module
 Write-Output "Checking Exchange Online Management module in Account $AzName"
 [array]$ExoPSModule = Get-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup | Where-Object {$_.Name -match "ExchangeOnlineManagement" }
  If ($ExoPSModule) {
    Write-Output ""
    Write-Output "Current Status"
    Write-Output "--------------"
    $ExoPSModule | Format-Table Name, Version, LastModifiedTime }
  
  $UpgradeNeeded = $True
  [array]$ModulesToUpdate = $ExoPSModule | Where-Object {$_.Version -ne $DesiredExoVersion}
  If (!($ModulesToUpdate)) {
     Write-Output "The Exchange Online Management module does not need to be updated for account $AzName"
     Write-Output ""
     $UpgradeNeeded = $False
  } Else {
    [string]$ModuleName = "ExchangeOnlineManagement"
    Write-Output ""
    Write-Output "Updating the Exchange Online management module to version $DesiredExoVersion"
    Write-Output "Removing old module..."
    Write-Output "Uninstalling module $ModuleName from Az Account $AzName"
    Remove-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -Confirm:$False -Force 
    $Uri = "https://www.powershellgallery.com/api/v2/package/$ModuleName/$DesiredExoVersion"
    Write-Output "Updating module $ModuleName from $Uri"
    $Status = (New-AzAutomationModule -AutomationAccountName $AzName -ResourceGroup $AzResourceGroup -Name $ModuleName -ContentLinkUri $Uri)
   }

} #End ForEach Az Account
Write-Output "All done. The modules in your Azure Automation accounts are now up to date for Powershell Runtime version 5.1"

# An example script used to illustrate a concept. More information about the topic can be found in the Office 365 for IT Pros eBook https://gum.co/O365IT/
# and/or a relevant article on https://office365itpros.com or https://www.practical365.com. See our post about the Office 365 for IT Pros repository 
# https://office365itpros.com/office-365-github-repository/ for information about the scripts we write.

# Do not use our scripts in production until you are satisfied that the code meets the needs of your organization. Never run any code downloaded from 
# the Internet without first validating the code in a non-production environment.
