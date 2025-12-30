#==============================================================================#
# (c) 2024 coolOrange s.r.l.                                                   #
#                                                                              #
# THIS SCRIPT/CODE IS PROVIDED "AS IS" WITHOUT WARRANTY OF ANY KIND, EITHER    #
# EXPRESSED OR IMPLIED, INCLUDING, BUT NOT LIMITED TO, THE IMPLIED WARRANTIES  #
# OF MERCHANTABILITY, FITNESS FOR A PARTICULAR PURPOSE, OR NON-INFRINGEMENT.   #
#==============================================================================#

foreach ($module in Get-Childitem "C:\ProgramData\coolOrange\powerAPS" -Name -Filter "*.psm1") {
    Import-Module "C:\ProgramData\coolOrange\powerAPS\$module" -Force -Global
}

function Get-DialogApsAttributeMapping($vaultFolderPath, [Hashtable]$mapping) {
    class DataContext {
        [System.Collections.ObjectModel.ObservableCollection[powerSync.MappingItem]] $Mapping
        [System.Collections.ObjectModel.ObservableCollection[object]] $AccAttributes
        [System.Collections.ObjectModel.ObservableCollection[object]] $VaultProperties
    
        DataContext() {
            $this.Mapping = New-Object System.Collections.ObjectModel.ObservableCollection[powerSync.MappingItem]
            $this.AccAttributes = New-Object System.Collections.ObjectModel.ObservableCollection[object]
            $this.VaultProperties = New-Object System.Collections.ObjectModel.ObservableCollection[object]
        }
    }

    $projectProperties = GetVaultAccProjectProperties $vaultFolderPath
    if (-not $projectProperties) {
        throw "ACC Project folder properties cannot be found!"
    }
    
    $hub = Get-ApsAccHub $projectProperties["Hub"]
    $project = Get-ApsProject -hub $hub -projectName $projectProperties["Project"]

    $projectFilesFolder = Get-ApsProjectFilesFolder $hub $project
    $customAttributes = Get-ApsAccCustomAttributeDefinitions $project $projectFilesFolder

    $dataContext = [DataContext]::new()
    $customAttributes | Sort-Object -Property name | ForEach-Object {
        $dataContext.AccAttributes.Add($_.name)
    }
    $file = GetVaultSingleFile
    $file | Get-Member -MemberType Properties | Sort-Object -Property Name | ForEach-Object {
        $dataContext.VaultProperties.Add($_.Name)
    }

    if ($mapping) {
        $mapping.GetEnumerator() | ForEach-Object {
            $dataContext.Mapping.Add((New-Object powerSync.MappingItem -Property @{ Acc = $_.Key; Vault = $_.Value }))
        }
    }
    
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.AttributeMapping.xaml")
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.DataContext = $dataContext
    
    $window.FindName('Ok').add_Click({
        # $dataGrid = $window.FindName("MappingGrid")
        # $lcv = [System.Windows.Data.CollectionViewSource]::GetDefaultView($dataGrid.ItemsSource)
        # if ($lcv.IsAddingNew) {
        #     $lcv.CommitNew()
        # }
        # if ($lcv.IsEditingItem) {
        #     $lcv.CommitEdit()
        # }
        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())
    
    if ($window.ShowDialog()) {
        $result = @{}
        $dataContext.Mapping | ForEach-Object {
            if ($_.Acc -and $_.Vault) {
                $result[$_.Acc] = $_.Vault
            }
        }
        return $result
    }

    return $null
}
