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

function Get-DialogApsReviewer($project, $step) {
    $itemsSource = @()
    $users = Get-ApsAccProjectUsers $project
    foreach ($user in ($users | Sort-Object -Property email)) {
        $itemsSource += New-Object PsObject -Property @{ 
            Name = $user.email; 
            Type = "User"; 
            IsSelected = $false; 
            Id=$user.autodeskId
        }
    }

    $step.candidates.companies | ForEach-Object {
        $itemsSource += New-Object PsObject -Property @{ 
            Name = $_.name; 
            Type = "Company"; 
            IsSelected = $false; 
            Id=$_.autodeskId
        }
    }

    $roles = @{
        "326052425" = "Architect"
        "326052429" = "BIM Manager"
        "326052434" = "Construction Manager"
        "326052435" = "Document Manager"
        "326052426" = "Engineer"
        "326052427" = "Estimator"
        "326052431" = "Executive"
        "326052437" = "Foreman"
        "326052428" = "IT"
        "326052430" = "Project Engineer"
        "326052432" = "Project Manager"
        "326052436" = "Superintendent"
        "326052433" = "VDC Manager"
    } | Sort-Object -Property Value
    $roles.GetEnumerator() | Sort-Object -Property Value | ForEach-Object {
        $itemsSource += New-Object PsObject -Property @{ 
            Name = $_.Value; 
            Type = "Role"; 
            IsSelected = $false; 
            Id=$_.Name
        }
    }

    $itemsSource = [System.Windows.Data.ListCollectionView]::new($itemsSource)
    $itemsSource.GroupDescriptions.Add((New-Object System.Windows.Data.PropertyGroupDescription "Type"))
  
    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.SelectReviewer.xaml")
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.FindName("Label").Content = "Select Candidates for '$($step.name.Replace("_", "__"))':"
    $window.FindName("Reviewer").ItemsSource = $itemsSource
    $window.FindName("Reviewer").SelectedValue = $selectedProject

    $window.FindName('Ok').add_Click({
        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())

    if ($window.ShowDialog()) {
        return $window.FindName("Reviewer").ItemsSource | Where-Object { $_.IsSelected }
    }

    return $null
}
