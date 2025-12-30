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

if ($processName -notin @('Connectivity.VaultPro')) {
	return
}

Add-VaultTab -Name "ACC Issues" -EntityType File -Action {
	param($selectedFile)
    if (-not (ApsTokenIsValid)) {
        return
    }
    
    $xamlFile = [xml](Get-Content "C:\ProgramData\coolOrange\Client Customizations\powerSync.Tab.ACC.Issues.xaml")
	$tab_control = [Windows.Markup.XamlReader]::Load( (New-Object System.Xml.XmlNodeReader $xamlFile) )
    
    $projectProperties = GetVaultAccProjectProperties $selectedFile._FolderPath
    if ($projectProperties -is [powerSync.Error]) {
        return $tab_control
    }

    $hub = Get-ApsAccHub $projectProperties["Hub"]
    if (-not $hub) {
        return $tab_control
    }

    $project = Get-ApsProject $hub $projectProperties["Project"]
    if (-not $project) {
        return $tab_control
    }

    $tab_control.FindName('Title').Content = $project.attributes.name
    $issues = Get-ApsAccIssues $project
    $issueTypes = Get-ApsAccIssueTypes $project
    $users = Get-ApsAccProjectUsers $project

    class Issue {
        [string] $id
        [string] $title
        [string] $status
        [string] $type
        [string] $subType     
        [string] $dueDate
        [string] $startDate
        [string] $updatedDate
        [string] $createdBy
        [string] $closedBy
        [string] $assignedTo
        [bool] $snapshotHasMarkups
        [bool] $overdue
    }

    class DataContext {
        [object] $Project
        [System.Collections.ObjectModel.ObservableCollection[Issue]] $Children

        DataContext() {
            $this.Children = New-Object System.Collections.ObjectModel.ObservableCollection[Issue]
        }
    }

    $dataContext = [DataContext]::new()
    $dataContext.Project = $project
    foreach($issue in $issues) {
        $issueType = $issueTypes | Where-Object { $_.id -eq $issue.issueTypeId }
        $issueSubtype = $issueType.subtypes | Where-Object { $_.id -eq $issue.issueSubtypeId }
        $assignedToUser = $users | Where-Object { $_.autodeskId -eq $issue.assignedTo }
        $issueObj = [Issue]::new()
        $issueObj.id = "#" + $issue.displayID
        $issueObj.title = $issue.title
        $issueObj.status = $issue.status.ToUpper()
        $issueObj.type = $issueType.title
        $issueObj.subType = $issueSubtype.title
        $issueObj.startDate = $issue.startDate
        $issueObj.dueDate = $issue.dueDate
        $issueObj.createdBy = $issue.createdBy
        $issueObj.closedBy = $issue.closedBy
        $issueObj.assignedTo = $assignedToUser.name
        $issueObj.snapshotHasMarkups = $issue.snapshotHasMarkups
        $issueObj.overdue = ($issue.dueDate -ge (Get-Date))
        $dataContext.Children.Add($issueObj)
    }
    $tab_control.DataContext = $dataContext

    $sortDescription = New-Object System.ComponentModel.SortDescription 'id', 'Ascending'
    $tab_control.FindName('IssuesTable').Items.SortDescriptions.Add($sortDescription)

    $tab_control.FindName('ButtonGoToProject').IsEnabled = $true
    $tab_control.FindName('ButtonGoToProject').add_Click({
        Start-Process "https://acc.autodesk.com/docs/issues/projects/$(($project.id -replace '^b\.', ''))/issues"
    }.GetNewClosure())

	return $tab_control
}