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

function Get-DialogApsProject($hub, $selectedProject = $null) {
    
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
        
        $projectsItemsSource += New-Object PsObject -Property @{ Project = $project; Type = $type }
    }  
  
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.SelectProject.xaml")
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.FindName("Label").Content = "Select Project from '$($hub.attributes.name.Replace("_", "__"))':"
    $window.FindName("Project").ItemsSource = $projectsItemsSource
    $window.FindName("Project").SelectedValue = $selectedProject
    
    $window.FindName('Ok').add_Click({
        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())   

    if ($window.ShowDialog()) {
        return $window.FindName("Project").SelectedValue
    }

    return $null
}