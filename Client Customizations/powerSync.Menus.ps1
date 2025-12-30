#==============================================================================#
# (c) 2024 coolOrange s.r.l.                                                   #
#                                                                              #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER    #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES  #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.   #
#==============================================================================#

# How to fix DPI display issues in Vault:
# https://www.autodesk.com/support/technical/article/caas/tsarticles/ts/gyzDnXXycpDjsEGyzJ7TY.html

foreach ($module in Get-Childitem "C:\ProgramData\coolOrange\powerAPS" -Name -Filter "*.psm1") {
    Import-Module "C:\ProgramData\coolOrange\powerAPS\$module" -Force -Global
}

if ($processName -notin @('Connectivity.VaultPro')) {
	return
}

#region Tools Menu
Add-VaultMenuItem -Location ToolsMenu -Name "powerSync: APS Authentication Settings..." -Action {
    $missingRoles = GetMissingRoles @(77, 76)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $settings = GetVaultApsAuthenticationSettings $true
    if ($settings -is [powerSync.Error]) {
        $settings = @{}
        $settings["Pkce"] = $false
        $settings["ClientId"] = ""
        $settings["ClientSecret"] = ""
        $settings["Scope"] = "account:read account:write data:read data:write"
        $settings["CallbackUrl"] = "http://localhost:8080/"
    }

    $settings = Get-DialogApsAuthentication $settings

    if ($settings) {
        SetVaultApsAuthenticationSettings $settings
    }
}

# Add-VaultMenuItem -Location ToolsMenu -Name "powerSync: APS Authentication Info..." -Action {
#     if (-not (ApsTokenIsValid)) {
#         return
#     }
#     $message = "You are connected to APS (Autodesk Platform Services)!"
#     $message += [System.Environment]::NewLine
#     $ApsConnection | Get-Member -MemberType Properties | ForEach-Object {
#         if ($_.Name -ne "Token" -and $_.Name -ne "RefreshToken" -and $_.Name -ne "Username" -and $_.Name -ne "RequestHeaders") {
#             $message += [System.Environment]::NewLine + $_.Name + ": " + $ApsConnection.$($_.Name)
#         }
#     }
#     [System.Windows.MessageBox]::Show(
#         $message, 
#         "powerSync: APS Authentication Info", 
#         "OK", 
#         "Information")
# }

# Add-VaultMenuItem -Location ToolsMenu -Name "powerSync: Close APS Connection" -Action {
#     Close-APSConnection
# }

Add-VaultMenuItem -Location ToolsMenu -Name "powerSync: ACC Default Account..." -Action {
    $missingRoles = GetMissingRoles @(77, 76)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    if (-not (ApsTokenIsValid)) {
        return
    }

    $hubName = GetVaultAccDefaultAccount
    $hub = Get-DialogApsHub "ACC" $hubName
    if (-not $hub) {
        return
    }
    
    SetVaultAccDefaultAccount $hub.attributes.name
}

Add-VaultMenuItem -Location ToolsMenu -Name "powerSync: Vault Folder Behaviors..." -Action {
    $missingRoles = GetMissingRoles @(77, 76)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    if (-not (ApsTokenIsValid)) {
        return
    }

    $behaviors = GetVaultAccFolderBehaviors $true
    if ($behaviors -is [powerSync.Error]) {
        $behaviors = @{}
        $behaviors["Category"] = ""
        $behaviors["Hub"] = ""
        $behaviors["Project"] = ""
        $behaviors["Folder"] = ""
    }
    $behaviors = Get-DialogApsProjectSettings $behaviors
    if ($behaviors) {
        SetVaultAccFolderBehaviors $behaviors
    }
}
#endregion

#region Folder Context Menu
Add-VaultMenuItem -Location FolderContextMenu -Name "powerSync: Assign ACC Project to Vault Folder..." -Action {
    param($entities)
    $folder = $entities[0]
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(77, 216, 217, 218)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $behaviors = GetVaultAccFolderBehaviors
    if ($behaviors -is [powerSync.Error]) {
        ShowPowerSyncErrorMessage -err $behaviors
        return
    }
    
    $hubName = GetVaultAccDefaultAccount
    $result = Get-DialogApsHubAndProject "ACC" $hubName
    $hub = $result.Hub
    $project = $result.Project
    if (-not $hub -or -not $project) {
        return
    }

    if ($project.attributes.extension.data.projectType -ne "ACC") {
        [System.Windows.MessageBox]::Show(
            "Currently only ACC projects are supported. Please select another project!", 
            "powerSync: Project type mismatch", 
            "OK", 
            "Error")
        return
    }

    $accFolder = Get-DialogApsContent $hub $project $false
    if (-not $accFolder) {
        return
    }

    $cats = $vault.CategoryService.GetCategoriesByEntityClassId("FLDR", $true)
    $cat = $cats | Where-Object { $_.Name -eq $behaviors["Category"] }
  
    $properties = @{
        $behaviors["Hub"] = $hub.attributes.name
        $behaviors["Project"] = $project.attributes.name
        $behaviors["Folder"]  = $accFolder.Path
    }

    $propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")
    $propInstParamArray = New-Object Autodesk.Connectivity.WebServices.PropInstParamArray
    $propInstParams = @()
    foreach ($prop in $properties.GetEnumerator()) {
        $propDef = $propDefs | Where-Object { $_.DispName -eq $prop.Name }
        $propInstParam = New-Object Autodesk.Connectivity.WebServices.PropInstParam
        $propInstParam.PropDefId = $propDef.Id
        $propInstParam.Val = $prop.Value
        $propInstParams += $propInstParam
    }
    $propInstParamArray.Items = $propInstParams
    
    $vault.DocumentServiceExtensions.UpdateFolderProperties(@($folder.Id), @($propInstParamArray))
    $vault.DocumentServiceExtensions.UpdateFolderCategories(@($folder.Id), @($cat.Id))
    
    [System.Windows.Forms.SendKeys]::SendWait('{F5}')
}

Add-VaultMenuItem -Location FolderContextMenu -Name "powerSync: Edit Attribute Mappings..." -Action {
    param($entities)
    $folder = $entities[0]
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(77, 253, 254)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $mapping = GetVaultAccAttributeMapping $folder._FullPath
    if ($mapping -is [powerSync.Error]) { 
        ShowPowerSyncErrorMessage -err $mapping
        return
    }

    $mapping = Get-DialogApsAttributeMapping $folder._FullPath $mapping
    if ($mapping -is [powerSync.Error]) {
        ShowPowerSyncErrorMessage -err $mapping
		return
    }

    if ($null -ne $mapping) {
        SetVaultAccAttributeMapping $folder._FullPath $mapping
    }
}

Add-VaultMenuItem -Location FolderContextMenu -Name "powerSync: Go To ACC Docs Project..." -Action {
    param($entities)
    $folder = $entities[0]
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(77)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $projectProperties = GetVaultAccProjectProperties $folder._FullPath
    if ($projectProperties -is [powerSync.Error]) {
        ShowPowerSyncErrorMessage -err $projectProperties
        return
    }

    $hub = Get-ApsAccHub $projectProperties["Hub"]
    if (-not $hub) {
        return
    }

    $project = Get-ApsProject $hub $projectProperties["Project"]
    if (-not $project) {
        return
    }

    Start-Process "https://acc.autodesk.com/docs/files/projects/$(($project.id -replace '^b\.', ''))"
}

Add-VaultMenuItem -Location FolderContextMenu -Name "powerSync: Go To ACC Build Project..." -Action {
    param($entities)
    $folder = $entities[0]
    if (-not (ApsTokenIsValid)) {
        return
    }

    $projectProperties = GetVaultAccProjectProperties $folder._FullPath
    if ($projectProperties -is [powerSync.Error]) {
        ShowPowerSyncErrorMessage -err $projectProperties
        return
    }

    $hub = Get-ApsAccHub $projectProperties["Hub"]
    if (-not $hub) {
        return
    }

    $project = Get-ApsProject $hub $projectProperties["Project"]
    if (-not $project) {
        return
    }

    Start-Process "https://acc.autodesk.com/build/files/projects/$(($project.id -replace '^b\.', ''))"
}

Add-VaultMenuItem -Location FolderContextMenu -Name "powerSync: Download File from Cloud Drive..." -Action {
    param($entities)
    $folder = $entities[0]
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(77, 4)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $hubName = GetVaultAccDefaultAccount
    $result = Get-DialogApsHubAndProject $null $hubName
    $hub = $result.Hub
    $project = $result.Project
    if (-not $hub -or -not $project) {
        return
    }

    $dialogResult = Get-DialogApsContent $hub $project $true
    $content = $dialogResult.Object

    if ($content.attributes.extension.type -eq "folders:autodesk.bim360:Folder" -or $content.attributes.extension.type -eq "folders:autodesk.core:Folder") {
        [System.Windows.MessageBox]::Show(
            "Please select a file. Downloads for folders are currently not implemented!", 
            "powerSync: Download File", 
            "OK", 
            "Warning")
        return
    }

    if (@("items:autodesk.core:File", "items:autodesk.fusion360:Design", "items:autodesk.bim360:File", "items:autodesk.bim360:C4RModel") -contains $content.attributes.extension.type) {
        #TODO: Implement progress dialog
        $version = Get-ApsTipVersion $project $content
        Get-ApsBucketFile $version "C:\TEMP"

        $localFullFileName = "C:\TEMP\$($version.attributes.displayName)"

        Add-VaultFile -From $localFullFileName -To ($folder._FullPath + "/" + $version.attributes.displayName)
        Remove-Item -Path $localFullFileName -Force -Verbose
        [System.Windows.Forms.SendKeys]::SendWait('{F5}')
    }
}
#endregion

#region File Context Menu
Add-VaultMenuItem -Location FileContextMenu -Name "powerSync: Publish Drawings as PDF to ACC" -Action {
    param($entities)
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(119)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $user = Get-ApsUserInfo
    if (-not $user) {
        [System.Windows.MessageBox]::Show(
            "Autodesk User Account Information cannot be determined!", 
            "powerSync: Autodesk Account Error", 
            "OK", 
            "Error")
        return
    }

    $excluded = @()
    $files = @()
    foreach($file in $entities){
        if ( @("idw", "dwg") -notcontains $file._Extension ) {
            $excluded += $file._Name
            continue
        }
        $files += $file
    }

    foreach($file in $files){
        Add-VaultJob -Name "powerSync.ACC.Publish.PDF" -Description "Translate file '$($file._Name)' to PDF and publish to ACC" -Parameters @{
            "FileVersionId" = $file.Id
            "EntityId"= $file.Id
            "EntityClassId"= $file._EntityTypeID
            "AccountId" = $user.eidm_guid
        }
    }

    if ($excluded.Count -gt 0) {
        [System.Windows.MessageBox]::Show(
            "$($files.Count) jobs(s) have been created.$([Environment]::NewLine)The following file(s) are not supported: $($excluded -join [Environment]::NewLine)", 
            "powerSync: Job Warning", 
            "OK", 
            "Warning")
    }    
}

Add-VaultMenuItem -Location FileContextMenu -Name "powerSync: Publish Models as DWF to ACC" -Action {
    param($entities)
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(119)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $user = Get-ApsUserInfo
    if (-not $user) {
        [System.Windows.MessageBox]::Show(
            "Autodesk User Account Information cannot be determined!", 
            "powerSync: Autodesk Account Error", 
            "OK", 
            "Error")
        return
    }

    $excluded = @()
    $files = @()
    foreach($file in $entities){
        if ( @("iam", "ipt", "dwg", "sldasm", "sldprt") -notcontains $file._Extension ) {
            $excluded += $file._Name
            continue
        }
        $files += $file
    }

    foreach($file in $files){
        Add-VaultJob -Name "powerSync.ACC.Publish.DWF" -Description "Translate file '$($file._Name)' to DWF and publish to ACC" -Parameters @{
            "FileVersionId" = $file.Id
            "EntityId"= $file.Id
            "EntityClassId"= $file._EntityTypeID
            "AccountId" = $user.eidm_guid
        }
    }

    if ($excluded.Count -gt 0) {
        [System.Windows.MessageBox]::Show(
            "$($files.Count) jobs(s) have been created.$([Environment]::NewLine)The following file(s) are not supported: $($excluded -join [Environment]::NewLine)", 
            "powerSync: Job Warning", 
            "OK", 
            "Warning")
    }    
}

Add-VaultMenuItem -Location FileContextMenu -Name "powerSync: Publish Inventor Assemblies as RVT to ACC" -Action {
    param($entities)
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(119)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $user = Get-ApsUserInfo
    if (-not $user) {
        [System.Windows.MessageBox]::Show(
            "Autodesk User Account Information cannot be determined!", 
            "powerSync: Autodesk Account Error", 
            "OK", 
            "Error")
        return
    }

    $excluded = @()
    $files = @()
    foreach($file in $entities){
        if ( @("iam") -notcontains $file._Extension ) {
            $excluded += $file._Name
            continue
        }
        $files += $file
    }

    foreach($file in $files){
        Add-VaultJob -Name "powerSync.ACC.Publish.RVT" -Description "Translate file '$($file._Name)' to RVT and publish to ACC" -Parameters @{
            "FileVersionId" = $file.Id
            "EntityId"= $file.Id
            "EntityClassId"= $file._EntityTypeID
            "AccountId" = $user.eidm_guid
        }
    }

    if ($excluded.Count -gt 0) {
        [System.Windows.MessageBox]::Show(
            "$($files.Count) jobs(s) have been created.$([Environment]::NewLine)The following file(s) are not supported: $($excluded -join [Environment]::NewLine)", 
            "powerSync: Job Warning", 
            "OK", 
            "Warning")
    }    
}

Add-VaultMenuItem -Location FileContextMenu -Name "powerSync: Publish Native Files to ACC" -Action {
    param($entities)
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(119)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $user = Get-ApsUserInfo
    if (-not $user) {
        [System.Windows.MessageBox]::Show(
            "Autodesk User Account Information cannot be determined!", 
            "powerSync: Autodesk Account Error", 
            "OK", 
            "Error")
        return
    }
    
    foreach($file in $entities){
        $projectFolder = GetVaultAccProjectFolder $file._FolderPath
        if ($projectFolder -is [powerSync.Error]) {
            ShowPowerSyncErrorMessage -err $projectFolder
            return
        }
    }

    foreach($file in $entities){
        Add-VaultJob -Name "powerSync.ACC.Publish.Native" -Description "Publish file '$($file._Name)' and it's references to ACC" -Parameters @{
            "FileVersionId" = $file.Id
            "EntityId"= $file.Id
            "EntityClassId"= $file._EntityTypeID
            "AccountId" = $user.eidm_guid
        }
    }
}

Add-VaultMenuItem -Location FileContextMenu -Name "powerSync: Publish Models for Clash Detection (NWC) to ACC" -Action {
    param($entities)
    if (-not (ApsTokenIsValid)) {
        return
    }

    $missingRoles = GetMissingRoles @(119)
    if ($missingRoles) {
        [System.Windows.MessageBox]::Show(
            "The current user does not have the required permissions: $missingRoles!", 
            "powerSync: Permission error", 
            "OK", 
            "Error")
        return
    }

    $user = Get-ApsUserInfo
    if (-not $user) {
        [System.Windows.MessageBox]::Show(
            "Autodesk User Account Information cannot be determined!", 
            "powerSync: Autodesk Account Error", 
            "OK", 
            "Error")
        return
    }

    $excluded = @()
    $files = @()
    foreach($file in $entities){
        if ( @("iam", "ipt", "dwg", "sldasm", "sldprt") -notcontains $file._Extension ) {
            $excluded += $file._Name
            continue
        }
        $files += $file
    }

    foreach($file in $files){
        Add-VaultJob -Name "powerSync.ACC.Publish.NWC" -Description "Translate file '$($file._Name)' to NWC and publish to ACC" -Parameters @{
            "FileVersionId" = $file.Id
            "EntityId"= $file.Id
            "EntityClassId"= $file._EntityTypeID
            "AccountId" = $user.eidm_guid
        }
    }

    if ($excluded.Count -gt 0) {
        [System.Windows.MessageBox]::Show(
            "$($files.Count) jobs(s) have been created.$([Environment]::NewLine)The following file(s) are not supported: $($excluded -join [Environment]::NewLine)", 
            "powerSync: Job Warning", 
            "OK", 
            "Warning")
    }    
}
#endregion
