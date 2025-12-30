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

function Get-DialogApsAuthentication([Hashtable]$settings) {
    class DataContext {
        [string] $ClientId
        [string] $ClientSecret
        [string] $CallbackUrl
        [System.Collections.ObjectModel.ObservableCollection[PsObject]] $Scope

        DataContext() {
            $this.Scope = New-Object System.Collections.ObjectModel.ObservableCollection[PsObject]
        }
    }

    $dataContext = [DataContext]::new()
    $dataContext.ClientId = $settings["ClientId"]
    $dataContext.ClientSecret = $settings["ClientSecret"]
    $dataContext.CallbackUrl = $settings["CallbackUrl"]

    if (-not $settings["Scope"]) {
        $settings["Scope"] = "data:read data:write account:read"
    }

    Add-Type -AssemblyName PresentationFramework, PresentationCore, WindowsBase
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.UI.Authentication.xaml")
 
    $window = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    $window.WindowStartupLocation = "CenterScreen"
    $window.Owner = $Host.UI.RawUI.WindowHandle
    $window.DataContext = $dataContext
    $window.FindName('ClientSecret').Password = $window.DataContext.ClientSecret

    $window.FindName('Authenticate').add_Click({
        $window.DataContext.ClientSecret = $window.FindName('ClientSecret').Password

        $arguments = @{
            ClientId = $window.DataContext.ClientId
            CallbackUrl = $window.DataContext.CallbackUrl
            ClientSecret = $window.DataContext.ClientSecret
        }

        if ($arguments.ClientId -eq "" -or $arguments.CallbackUrl -eq "" -or $arguments.ClientSecret -eq "") {
            [System.Windows.MessageBox]::Show("Please fill in all fields", "APS Connection Test", "OK", "Error")
            return
        }
    
        $vaultLogin = [Autodesk.Connectivity.WebServicesTools.AutodeskAccount]::Login([IntPtr]::Zero)
        if ($vaultLogin -and $vaultLogin.AccountEmail) {
            $arguments.Username = $vaultLogin.AccountEmail
        }
        
        $connected = Connect-APS -ClientId $window.DataContext.ClientId -CallbackUrl $window.DataContext.CallbackUrl -ClientSecret $window.DataContext.ClientSecret
    

        if ($connected) {
            [System.Windows.MessageBox]::Show("Autodesk Platform Services (APS) Connection successful", "APS Connection Test", "OK", "Information")
        } else {
            [System.Windows.MessageBox]::Show("Autodesk Platform Services (APS) Connection failed", "APS Connection Test", "OK", "Error")
        }

        $window.DialogResult = $true
        $window.Close()
    }.GetNewClosure())

    $window.FindName('GoTo').add_Click({
        Start-Process "https://aps.autodesk.com/myapps/"
    }.GetNewClosure())

    if ($window.ShowDialog()) {
        $settings["ClientId"] = $dataContext.ClientId
        $settings["ClientSecret"] = $dataContext.ClientSecret
        $settings["CallbackUrl"] = $dataContext.CallbackUrl
        return $settings
    }

    return $null
}