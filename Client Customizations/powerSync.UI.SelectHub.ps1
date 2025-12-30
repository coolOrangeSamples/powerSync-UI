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

function Get-DialogApsHub($typeFilter = $null, $selectedHub = $null) {
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
  
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.SelectHub.xaml")
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.FindName("Hub").ItemsSource = $hubsItemsSource
    $window.FindName("Hub").SelectedValue = $hubs | Where-Object { $_.attributes.name -eq $selectedHub }

    $window.FindName('Ok').add_Click({
        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())

    if ($window.ShowDialog()) {
        return $window.FindName("Hub").SelectedValue
    }

    return $null
}