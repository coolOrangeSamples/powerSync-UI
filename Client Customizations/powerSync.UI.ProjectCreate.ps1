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

function Get-DialogApsProjectCreate($hub, $name = "", $number = "") {

    class DataContext {
        [string] $Name
        [string] $Number
        [string] $ProjectType
        [psobject] $Template
        [System.Collections.ObjectModel.ObservableCollection[PsObject]] $ProjectTypes
        [System.Collections.ObjectModel.ObservableCollection[PsObject]] $Templates

        DataContext() {
            $this.ProjectTypes = New-Object System.Collections.ObjectModel.ObservableCollection[string]
            $this.Templates = New-Object System.Collections.ObjectModel.ObservableCollection[PsObject]
        }
    }

    $dataContext = [DataContext]::new()
    $dataContext.Name = $name
    $dataContext.Number = $number
    $dataContext.ProjectType = "Airport"

    $projectTypes = @("Airport", "Assisted Living / Nursing Home", "Bridge", "Canal / Waterway", "Convention Center", "Court House", 
        "Data Center", "Dams / Flood Control / Reservoirs", "Demonstration Project", "Dormitory", "Education Facility", 
        "Government Building", "Harbor / River Development", "Hospital", "Hotel / Motel", "Library", "Manufacturing / Factory", 
        "Medical Laboratory", "Medical Office", "Military Facility", "Mining Facility", "Multi-Family Housing", "Museum", "Oil & Gas", 
        "Office", "OutPatient Surgery Center", "Parking Structure / Garage", "Performing Arts", "Plant", "Power Plant", 
        "Prison / Correctional Facility", "Rail", "Recreation Building", "Religious Building", "Research Facility / Laboratory", 
        "Restaurant", "Retail", "Seaport", "Single-Family Housing", "Solar Farm", "Stadium/Arena", "Streets / Roads / Highways", 
        "Template Project", "Theme Park", "Training Project", "Transportation Building", "Tunnel", "Utilities", 
        "Warehouse (non-manufacturing)", "Waste Water / Sewers", "Water Supply", "Wind Farm")
    $projectTypes | ForEach-Object {
        $dataContext.ProjectTypes.Add($_)
    }

    $projectTemplates = @(Get-ApsAccProjectTemplates $hub)
    $projectTemplates | ForEach-Object {
        $dataContext.Templates.Add($_)
    }

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.ProjectCreate.xaml")
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.DataContext = $dataContext
    $window.FindName('Ok').add_Click({
        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())

    if ($window.ShowDialog()) {
        $name = $dataContext.Name
        $number = $dataContext.Number
        $type = $dataContext.ProjectType
        $templateId = $dataContext.Template.id
        $accProject = Add-ApsAccProject $hub $name $number $type $templateId
        while ($accProject.status -ne "active") {
            $accProject = Get-ApsAccProject $accProject.id
        }

        Import-Module powerVault
        $vaultLogin = [Autodesk.Connectivity.WebServicesTools.AutodeskAccount]::Login([IntPtr]::Zero)
        $email = $vaultLogin.AccountEmail
        $user = Find-ApsAccProjectUser $dataContext.Template $email
        if (-not $user) {
            throw "User not found in Vault"
        }
        $u = Add-ApsAccProjectUser $accProject $user.id
        while($u.status -ne "active") {
            $u = Get-ApsAccProjectUser $accProject $user.id
        }

        $project = Get-ApsProject $hub $accProject.name
        return $project
    }
    
    return $null
}