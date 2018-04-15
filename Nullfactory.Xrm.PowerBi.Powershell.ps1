function Get-Report {
    param(
        [string]$token,
        [string] $groupId
    )
   $result =  Invoke-RestMethod `
        -Method Get `
        -Uri  "https://api.powerbi.com/v1.0/myorg/groups/$groupId/reports" `
        -Headers @{ Authorization = "Bearer $token" } `
        -Verbose

    return $result;
}

function Invoke-ReportDeploy {
    param(
        [string] $token,
        [string] $reportFilePath,
        [string] $groupId,
        [switch] $overwriteFileIfExists
    )
    
    if($overwriteFileIfExists)
    {        
        $existingReports = Get-Report -token $token -groupId $groupId
        $fileNameWithoutExtension = [IO.Path]::GetFileNameWithoutExtension($reportFilePath);
        $currentReportExists = $existingReports.value | Where-Object {$_.name -eq $fileNameWithoutExtension} | Select-Object Id

        if($currentReportExists)
        {
            $overwriteFile = $true;
        }
        else
        {
            $overwriteFile = $false;
        }
    }

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

    if ($overwriteFile) 
    {
        $uri = $uri + "&nameConflict=Overwrite"
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

function Get-ServiceAccessToken {
    param(
        [Parameter(Mandatory = $true)]
        [string]$authorityName,
        [Parameter(Mandatory = $true)]
        [string]$username,
        [Parameter(Mandatory = $true)]
        [string]$password,
        [Parameter(Mandatory = $true)]
        [guid]$clientId
    )

    $token = Get-ADALAccessToken -AuthorityName $authorityName `
        -ClientId $clientId `
        -ResourceId "https://analysis.windows.net/powerbi/api" `
        -UserName $username `
        -Password $password

    return $token
}

<#
    .SYNOPSIS
		Deploy a Power BI report to a resource group.
    .DESCRIPTION
        This scripts uploads a pbix file into the PowerBi Service.

    Pre-requisites:
      - PowerBi Pro account.
      - Microsoft.ADAL.PowerShell 1.12

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
	.PARAMETER groupName
        The parameter specifying the resource group name.
    .PARAMETER groupId
		The parameter specifying the resource group Id.
    .PARAMETER reportFileName
    
    .PARAMETER reportFolder

    .PARAMETER overwriteFile

    .PARAMETER createGroupIfMissing

    .EXAMPLE
        .\Deploy-PowerBiReport.ps1 -authorityName "sndbx26.onmicrosoft.com" -username "admin@sndbx26.onmicrosoft.com" -password "P@ssw0rd!" -clientId "19da8650-b202-4bc9-95f3-e8daf38ec39e" -resourceGroupName "Sandbox Analytics" -reportFileName ".\SuperReport.pbix"
        
        .\Deploy-PowerBiReport.ps1 -authorityName "sndbx26.onmicrosoft.com" -username "admin@sndbx26.onmicrosoft.com" -password "P@ssw0rd!" -clientId "19da8650-b202-4bc9-95f3-e8daf38ec39e" -resourceGroupName "Sandbox Analytics" -reportFileName ".\SuperReport.pbix" -overwriteIfExists -createGroupIfMissing
#>
function Import-Report {
    param(
        [Parameter(Mandatory = $true)]
        [string]$token,
        [Parameter(ParameterSetName = "GroupNameReportFile", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupNameReportFolder", Mandatory = $true)]
        [string]$groupName,
        [Parameter(ParameterSetName = "GroupIdReportFile", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupIdReportFolder", Mandatory = $true)]
        [string]$groupId,
        [Parameter(ParameterSetName = "GroupNameReportFile", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupIdReportFile", Mandatory = $true)]
        [string]$reportFileName,
        [Parameter(ParameterSetName = "GroupIdReportFolder", Mandatory = $true)]
        [Parameter(ParameterSetName = "GroupNameReportFolder", Mandatory = $true)]
        [string]$reportFolder,
        [switch]$overwriteFileIfExists,
        [Parameter(ParameterSetName = "GroupNameReportFile", Mandatory = $false)]
        [Parameter(ParameterSetName = "GroupNameReportFolder", Mandatory = $false)]
        [switch]$createGroupIfMissing
    )

    if (-Not (Get-Module -ListAvailable -Name  Microsoft.ADAL.PowerShell)) 
    {
        Write-Verbose "Initializing Microsoft.ADAL.PowerShell module ..."
        Install-Module -Name  Microsoft.ADAL.PowerShell -Scope CurrentUser -ErrorAction SilentlyContinue -Force
    }

    if (!$groupId) 
    {
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
    }

    Write-Verbose "Found GroupId: $groupId"

    if ($reportFolder) 
    {
        $files = Get-ChildItem -Path $reportFolder\* -Include *.pbix, *.xlsx, *.xlxm, *.csv
        foreach ($file in $files) 
        {
            Invoke-ReportDeploy -token $token -reportFilePath $file -groupId $groupId -overwriteFileIfExists:$overwriteFileIfExists
        }
    }
    else
    {
        Invoke-ReportDeploy -token $token -reportFilePath $reportFileName -groupId $groupId -overwriteFileIfExists:$overwriteFileIfExists
    }
}

