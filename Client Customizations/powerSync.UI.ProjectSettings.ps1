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

function Get-DialogApsProjectSettings([Hashtable]$settings) {

    $categories = $vault.CategoryService.GetCategoriesByEntityClassId("FLDR", $true)
    $propDefs = $vault.PropertyService.GetPropertyDefinitionsByEntityClassId("FLDR")

    class DataContext {
        [string] $Category
        [string] $Hub
        [string] $Project
        [string] $Folder
        [System.Collections.ObjectModel.ObservableCollection[PsObject]] $Categories
        [System.Collections.ObjectModel.ObservableCollection[PsObject]] $Properties

        DataContext() {
            $this.Categories = New-Object System.Collections.ObjectModel.ObservableCollection[PsObject]
            $this.Properties = New-Object System.Collections.ObjectModel.ObservableCollection[PsObject]
        }
    }

    $dataContext = [DataContext]::new()
    $dataContext.Category = $settings["Category"]
    $dataContext.Hub = $settings["Hub"]
    $dataContext.Project = $settings["Project"]
    $dataContext.Folder = $settings["Folder"]

    $categories | Sort-Object -Property Name | ForEach-Object {
        $dataContext.Categories.Add($_)
    }

    $propDefs | Sort-Object -Property DispName | ForEach-Object {
        $dataContext.Properties.Add($_)
    }

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.ProjectSettings.xaml")
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.DataContext = $dataContext
    $window.FindName('Ok').add_Click({
        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())

    if ($window.ShowDialog()) {
        $settings["Category"] = $dataContext.Category
        $settings["Hub"] = $dataContext.Hub
        $settings["Project"] = $dataContext.Project
        $settings["Folder"] = $dataContext.Folder
        return $settings
    }
    return $null
}