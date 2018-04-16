Function Get-PowerBiReport {
    <#
    .SYNOPSIS
		Retrieves a list of all reports in the PowerBI group.
    .DESCRIPTION
        Retrieves a list of all reports in the PowerBI group.
	.NOTES
		Author: Shane Carvalho
	.LINK
		https://nullfactory.net
	.PARAMETER token
        Mandatory parameter that specifies the ADAL access token.
    .PARAMETER groupId
		Mandatory parameter is the group idenfier for which the reports retreived.
    .EXAMPLE
        Get-PowerBiReport -token $token -groupId fcf96fa6-ee3f-4a7e-bd52-3d4c5c6c5e48

        Retrieve a list of reports for the specified group identifier.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$token,
        [Parameter(Mandatory = $true)]
        [guid] $groupId
    )
   $result =  Invoke-RestMethod `
        -Method Get `
        -Uri  "https://api.powerbi.com/v1.0/myorg/groups/$groupId/reports" `
        -Headers @{ Authorization = "Bearer $token" } `
        -Verbose

    return $result;
}

Function Invoke-PowerBiReportDeployment {
    <#
    .SYNOPSIS
		Deploys the report to the specified Power BI app workspace group.
    .DESCRIPTION
        Deploys the report to the specified Power BI app workspace group.
	.NOTES
		Author: Shane Carvalho
	.LINK
		https://nullfactory.net
	.PARAMETER token
        Mandatory parameter that specifies the ADAL access token.
    .PARAMETER groupId
        Mandatory parameter is the app workspace group identifier into which the reports will be deployed.
	.PARAMETER reportFilePath
		Mandatory paramater that specifies the report file path of an individual report.
	.PARAMETER overwriteReportIfExists
		This optional switch tells the script if it should overwrite an existing report if one with the same name is found.
    .EXAMPLE
        Invoke-PowerBiReportDeployment -token $token -groupId  fcf96fa6-ee3f-4a7e-bd52-3d4c5c6c5e48 -reportFilePath c:\myreports\Test.pbix 

        Imports the report Test.pbix to app workspace group fcf96fa6-ee3f-4a7e-bd52-3d4c5c6c5e48
    .EXAMPLE
        Invoke-PowerBiReportDeployment -token $token -groupId  fcf96fa6-ee3f-4a7e-bd52-3d4c5c6c5e48 -reportFilePath c:\myreports\Test.pbix -overwriteReportIfExists

        Imports the report Test.pbix to app workspace group fcf96fa6-ee3f-4a7e-bd52-3d4c5c6c5e48 and if a report exists with the same name, then overwrite it with the new one.
    #>
    param(
        [string] $token,
        [guid] $groupId,
        [string] $reportFilePath,
        [switch] $overwriteReportIfExists
    )

    $path = Resolve-Path $reportFilePath
    $fileName = [IO.Path]::GetFileName($path)
    
    $filebody = [System.IO.File]::ReadAllBytes($path)
    $encoding = [System.Text.Encoding]::GetEncoding("iso-8859-1")
    $filebodytemplate = $encoding.GetString($filebody)
    
    $boundary = [guid]::NewGuid().ToString()
    [System.Text.StringBuilder]$contents = New-Object System.Text.StringBuilder
    $contents.AppendLine("--$boundary")
    $contents.AppendLine("Content-Disposition: form-data; name=""fileData""; filename=""$fileName""")
    $contents.AppendLine("Content-Type: application/octet-stream")
    $contents.AppendLine()
    $contents.AppendLine($filebodytemplate)
    $contents.AppendLine("--$boundary--")
    $body1 = $contents.ToString()
    
    $headers = @{ 
        "Authorization" = "Bearer $token" 
        "Content-Type"  = "application/json"
    }
    
    [string]$uri = "https://api.powerbi.com/v1.0/myorg/groups/$groupId/imports?datasetDisplayName=$fileName";

    if($overwriteReportIfExists)
    {        
        $existingReports = Get-PowerBiReport -token $token -groupId $groupId
        $fileNameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($reportFilePath);
        $currentReportExists = $existingReports.value | Where-Object {$_.name -eq $fileNameWithoutExtension} | Select-Object Id

        if($currentReportExists)
        {
            $uri = $uri + "&nameConflict=Overwrite"
        }
        else
        {
            $uri = $uri + "&nameConflict=Abort"
        }
    }
    else
    {
        $uri = $uri + "&nameConflict=Abort"
    }

    Invoke-RestMethod `
        -Method Post `
        -Uri  $uri `
        -Headers $headers `
        -Body $body1 `
        -ContentType "multipart/form-data; boundary=--$boundary" `
        -Verbose
}

Function Get-PowerBiAccessToken {
    <#
    .SYNOPSIS
		Retrieve Acccess Token used to access the Power BI Service.
    .DESCRIPTION
        Retrieve Acccess Token used to access the Power BI Service.
	.NOTES
		Author: Shane Carvalho
	.LINK
		https://nullfactory.net
	.PARAMETER authorityName
		Mandatory parameter that specifies the authentication authority name.
	.PARAMETER username
		Mandatory paramater is the username used to connect to PowerBi portal.
	.PARAMETER password
		Mandatory paramerter is the password of the user used to connect to the service.
	.PARAMETER clientId
		Required parameter. The client id used to connect to the service.
    .EXAMPLE
        Get-PowerBiAccessToken -authorityName "sndbx26.onmicrosoft.com" -username "admin@sndbx26.onmicrosoft.com" -password "P@ssw0rd!" -clientId "19da8650-b202-4bc9-95f3-e8daf38ec39e" 

        Retrieve the Adal Access Token for the client application.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$authorityName,
        [Parameter(Mandatory = $true)]
        [string]$username,
        [Parameter(Mandatory = $true)]
        [string]$password,
        [Parameter(Mandatory = $true)]
        [string]$clientId
    )

    $token = Get-ADALAccessToken -AuthorityName $authorityName `
        -ClientId $clientId `
        -ResourceId "https://analysis.windows.net/powerbi/api" `
        -UserName $username `
        -Password $password

    return $token;
}

Function Get-PowerBiAppWorkspaceGroupId {
    <#
    .SYNOPSIS
		Retrieve the PowerBi App Workspace Group Id
    .DESCRIPTION
        Retrieve the PowerBi App Workspace Group Id by Name
	.NOTES
		Author: Shane Carvalho
	.LINK
		https://nullfactory.net
	.PARAMETER token
		Mandatory parameter that specifies the ADAL access token.
	.PARAMETER groupName
        Mandatory paramater specifying the name of the app workspaces group.
	.PARAMETER createGroupIfMissing
		Switch indicating if a new group should be created if one does not exists
    .EXAMPLE
        Get-PowerBiAppWorkspaceGroupId -token $token -groupName "Sandbox Analytics" -

        Retrieves the groupId of the app workspace group based on the name.
    .EXAMPLE
        Get-PowerBiAppWorkspaceGroupId -token $token -groupName "Sandbox Analytics" -createGroupIfMissing

        Retrieves the app workspace group identifier of the report based on the report name. If it does not exists, a new one will be created.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$token,
        [Parameter(Mandatory = $true)]
        [string]$groupName,
        [switch]$createGroupIfMissing
    )

    Write-Information "Attempting to retrieve group by name $groupName...";
    $groupInfo = (Invoke-RestMethod -Method Get -Uri "https://api.powerbi.com/v1.0/myorg/groups" -Headers @{ Authorization = "Bearer $token" }).value | Where-Object { $_.name -eq $groupName }

    if (!$groupInfo) 
    {
        if ($createGroupIfMissing) 
        {
            Write-Information "Group not found, attempting to create new: $groupName";

            try 
            {
                $newGroup = Invoke-RestMethod -Method Post -Uri "https://api.powerbi.com/v1.0/myorg/groups" -Headers @{ Authorization = "Bearer $token" } -Body "{ ""name"": ""$groupName"" }" -ContentType "application/json"
            }
            catch 
            {
                throw "Error while attempting to create new group: $groupName";
            }

            $groupId = $newGroup.id;

            Write-Information "Group $groupId successfully created.";
        }
        else 
        {
            throw "Unable to find group name $groupName";
        }
    }
    else 
    {
        $groupId = $groupInfo.id;
    }

    return $groupId;
}


Function Import-PowerBiReport {
    <#
    .SYNOPSIS
		Deploy a Power BI report to a app workspace group.
    .DESCRIPTION
        Deploys a Power BI report to the specified app workspace group.

        Pre-requisites:
        - PowerBi Pro account.
        - Microsoft.ADAL.PowerShell 1.12
	.NOTES
		Author: Shane Carvalho
	.LINK
		https://nullfactory.net
	.PARAMETER groupName
        The parameter specifying the app workspace group name into which the report would be deployed to. This parameter is ignored if groupId is provided.
    .PARAMETER groupId
		The parameter specifying the app group identifier into which the report would be deployed to.
    .PARAMETER reportFileName
        This parameter specifies the single report that would be used to deploy. This parameter is ignored if a reportFolder parameter is provided.
    .PARAMETER reportFolder
        This parameter specifies the folder which contains the reports that should be uploaded. The script would look for the following types of files: *.pbix, *.xlsx, *.xlxm, *.csv
    .PARAMETER overwriteReportIfExists
        This optional switch tells the script if it should overwrite an existing report if one with the same name is found.
    .PARAMETER createGroupIfMissing
        This optional switch tells the script to create a new group with same name if one does not exists. This parameter is ignored if groupId is provided.
    .EXAMPLE
        Import-PowerBiReport -token $token -groupName "Sandbox Analytics" -reportFolder ".\ReportFolder\" -createGroupIfMissing
        
        All reports in the .\ReportFolder reports folder is imported to the PowerBI app workspace group "Sandbox Analytics"
        If a group with the sepcified name does not exists, the script would create one for you. However, if a report with the same name exists, the operation would fail.
    .EXAMPLE
        Import-PowerBiReport -token $token -groupName "Sandbox Analytics" -reportFileName ".\SuperReport.pbix" -overwriteReportIfExists

        The SuperReport.pbix report is imported to the remote PowerBI app workspace group "Sandbox Analytics". If a report with the same exists in the app workspace group, it would be overwritten.
        However, if a group with the specified name does not exists, the operation would fail.
    .EXAMPLE
        Import-PowerBiReport -groupId $token -groupId e0cbc83b-6629-43fd-9c69-25be3f6e3188 -reportFolder ".\ReportFolder\" -overwriteReportIfExists
        
        All reports in the .\ReportFolder reports folder is imported to the PowerBI app workspace group  e0cbc83b-6629-43fd-9c69-25be3f6e3188. 
        If report with the same name exists in the same app workspace group, the script would replace it with the new report.
    .EXAMPLE
        Import-PowerBiReport -groupId $token -groupId e0cbc83b-6629-43fd-9c69-25be3f6e3188 -reportFileName ".\SuperReport.pbix"

        The report SuperReport.pbix is imported to resoure group e0cbc83b-6629-43fd-9c69-25be3f6e3188.
    #>
    param(
        [Parameter(Mandatory = $true)]
        [string]$token,
        [Parameter(ParameterSetName = "GroupNameReportFile", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupNameReportFolder", Mandatory = $true)]
        [string]$groupName,
        [Parameter(ParameterSetName = "GroupIdReportFile", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupIdReportFolder", Mandatory = $true)]
        [guid]$groupId,
        [Parameter(ParameterSetName = "GroupIdReportFile", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupNameReportFile", Mandatory = $true)]
        [string]$reportFileName,
        [Parameter(ParameterSetName = "GroupIdReportFolder", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupNameReportFolder", Mandatory = $true)]
        [string]$reportFolder,
        [Parameter(Mandatory = $false)]
        [switch]$overwriteReportIfExists,
        [Parameter(Mandatory = $false)]
        [switch]$createGroupIfMissing
    )

    if (-Not (Get-Module -ListAvailable -Name  Microsoft.ADAL.PowerShell)) 
    {
        Write-Verbose "Initializing Microsoft.ADAL.PowerShell module ..."
        Install-Module -Name  Microsoft.ADAL.PowerShell -Scope CurrentUser -ErrorAction SilentlyContinue -Force
    }

    if (!$groupId) 
    {
        # Write-Information "Attempting to retrieve group by name $groupName...";
        # $groupInfo = (Invoke-RestMethod -Method Get -Uri "https://api.powerbi.com/v1.0/myorg/groups" -Headers @{ Authorization = "Bearer $token" }).value | Where-Object { $_.name -eq $groupName }

        # if (!$groupInfo) 
        # {
        #     if ($createGroupIfMissing) 
        #     {
        #         Write-Information "Group not found, attempting to create new: $groupName";

        #         try 
        #         {
        #             $newGroup = Invoke-RestMethod -Method Post -Uri "https://api.powerbi.com/v1.0/myorg/groups" -Headers @{ Authorization = "Bearer $token" } -Body "{ ""name"": ""$groupName"" }" -ContentType "application/json"
        #         }
        #         catch 
        #         {
        #             throw "Error while attempting to create new group: $groupName";
        #         }

        #         $groupId = $newGroup.id;

        #         Write-Information "Group $groupId successfully created.";
        #     }
        #     else 
        #     {
        #         throw "Unable to find group name $groupName";
        #     }
        # }
        # else 
        # {
        #     $groupId = $groupInfo.id;
        # }

        $groupId = Get-PowerBiAppWorkspaceGroupId -token $token -groupName $groupName -createGroupIfMissing:$createGroupIfMissing
    }

    Write-Verbose "Using GroupId: $groupId"

    if ($reportFolder) 
    {
        $files = Get-ChildItem -Path $reportFolder\* -Include *.pbix, *.xlsx, *.xlxm, *.csv
        foreach ($file in $files) 
        {
            Invoke-PowerBiReportDeployment -token $token -groupId $groupId -reportFilePath $file -overwriteReportIfExists:$overwriteReportIfExists
        }
    }
    else
    {
        Invoke-PowerBiReportDeployment -token $token -groupId $groupId -reportFilePath $reportFileName -overwriteReportIfExists:$overwriteReportIfExists
    }
}