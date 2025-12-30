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

function Get-DialogApsHubAndProject($typeFilter = $null, $selectedHub = $null, $selectedProject = $null) {

    class HubAndProjectDialogResult {
        [object] $Hub
        [object] $Project
    }

    $hubsItemsSource = @()

    if ($typeFilter -eq "Fusion") {
        $hubs = Get-ApsCoreHubs
    } elseif ($typeFilter -eq "Personal") {
        $hubs = Get-ApsPersonalHubs
    } elseif ($typeFilter -eq "ACC") {
        $hubs = Get-ApsAccHubs
    } else {
        $hubs = Get-ApsHubs
    }
    foreach ($hub in $hubs) {
        switch ($hub.attributes.extension.type) {
            "hubs:autodesk.core:Hub" { 
                $type = "Fusion"
            }
            "hubs:autodesk.a360:PersonalHub" {
                $type = "Fusion"
            }
            "hubs:autodesk.bim360:Account" {
                $type = "ACC"
            }
            default {
                $type = $null
            }
        }
        $hubsItemsSource += New-Object PsObject -Property @{ Hub = $hub; Type = $type }
    }

    $loadProject = {
        param($hub)

        $window.Cursor = [System.Windows.Input.Cursors]::Wait
        $projectsItemsSource = @()
        $projects = Get-ApsProjects $hub
        foreach ($project in $projects) {
            switch ($project.attributes.extension.type) {
                "projects:autodesk.core:Project" { 
                    $type = "Fusion"
                }
                "projects:autodesk.bim360:Project" {
                    if ($project.attributes.extension.data.projectType -eq "ACC") {
                        $type = "ACC"
                    } else {
                        $type = "BIM360"
                    }
                }
                default {
                    $type = $null
                }
            }
            $projectsItemsSource += New-Object PsObject -Property @{ Project = $project; Type = $type; Name = $project.attributes.name }
        }
        $projectsItemsSource = $projectsItemsSource | Sort-Object -Property Name
        #$projectsItemsSource = $projectsItemsSource | Where-Object { -not $_.Project.attributes.name.Contains("test") -and -not $_.Project.attributes.name.Contains("Test") }
        $window.Cursor = $null
        return $projectsItemsSource
    }
  
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.SelectHubAndProject.xaml")
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.FindName('Hub').ItemsSource = $hubsItemsSource

    if ($selectedHub) {
        $window.FindName('Hub').SelectedValue = $hubs | Where-Object { $_.attributes.name -eq $selectedHub }
        $projectsItemsSource = @(Invoke-Command $loadProject -ArgumentList $window.FindName('Hub').SelectedItem.Hub)
        $window.FindName('Project').ItemsSource = $projectsItemsSource
        $window.FindName('Project').SelectedValue = ($projectsItemsSource | Where-Object { $_.Project.attributes.name -eq $selectedProject }).Project
    }
    
    $window.FindName('Hub').add_SelectionChanged({
        $projectsItemsSource = @(Invoke-Command $loadProject -ArgumentList $window.FindName('Hub').SelectedItem.Hub)
        $window.FindName('Project').ItemsSource = $projectsItemsSource
        $window.FindName('Project').SelectedValue = $null
    }.GetNewClosure())
            
    $window.FindName('Ok').add_Click({
        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())

    if ($window.ShowDialog()) {
        $result = [HubAndProjectDialogResult]::new()
        $result.Hub = $window.FindName('Hub').SelectedValue
        $result.Project = $window.FindName('Project').SelectedValue
        return $result
    }

    return $null
}
