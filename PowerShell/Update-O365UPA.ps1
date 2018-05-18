<# 
    Requires SharePoint Client Components
    https://download.microsoft.com/download/B/3/D/B3DA6839-B852-41B3-A9DF-0AFA926242F2/sharepointclientcomponents_16-6906-1200_x64-en-us.msi
#>

#region Set Variables
    $adminUrl = "https://pcfromdc-admin.sharepoint.com"                                             # Read-Host -Prompt 'Enter the admin URL of your tenant'
    $userName = "pc@pcfromdc.com"                                                                   # Read-Host -Prompt 'Enter your user name'
    $siteURL = "https://pcfromdc.sharepoint.com/sites/spc18"                                        # Read-Host -Prompt 'Enter the URL of the site of the file'
    $importFileUrl = "https://pcfromdc.sharepoint.com/sites/spc18/upaSync"                          # Read-Host -Prompt 'Enter the URL to the file located in your tenant'
    $docLibName = "UPA Sync"                                                                        # Read-Host -Prompt 'Enter the name of the document library'
    $filePath = "C:\UPA Scripts\upaOutput.txt"                                                      # Read-Host -Prompt 'Enter the JSON File Output Path'
    $uploadPath = "https://pcfromdc.sharepoint.com/sites/spc18/upaSync/upaOutput-PowerShell.txt"    # Read-Host -Prompt 'Enter the JSON File Upload Path'
#endregion

#region Set Password
    function createPassword {
        Write-Host "Enter Password for $userName" -ForegroundColor Cyan
        Write-Host
        $PWInput = Read-Host -AsSecureString | ConvertFrom-SecureString
        $PWInput | Out-File C:\Scripts\pcPW.txt -Force
    }
#endregion

#region Read SQL and Create JSON'ish output
    # Read SQL Table and filter items
    $table = Read-SqlTableData -TableName "pcDemo_SystemUsers" -DatabaseName "pcDemo_personnel" -ServerInstance "SQL2017-01" -SchemaName 'dbo' -TopN 100 | 
             Where-Object {($_.mail -like "*@pcdemo.net" -or $_.mail -like "*@pcfromdc.com") -and ($_.City -ne $null)} | select mail, city

    # Start Creating JSON Output String
    $jasonOutput =  '{' + "`n"
    $jasonOutput += '"value":' + "`n"
    $jasonOutput += '['  + "`n"

    # Loop through SQL Table rows and add users to JSON Output String
    foreach ($user in $table) {
            $jasonOutput +=  '{' + "`n"
            $jasonOutput +=  '"IdName": "' + $user.mail + '",' + "`n"
            $jasonOutput +=  '"Property1": "' + $user.City + '"' + "`n"
          # $jasonOutput +=  '"Property2": "' + $user.State + '"' + "`n"
            $jasonOutput +=  '},' + "`n"
    }

    # Close JSON Output String
    $jasonOutput = ($jasonOutput.Trim()).TrimEnd(",")
    $jasonOutput += "`n"
    $jasonOutput += ']' + "`n"
    $jasonOutput += '}' + "`n"

    # Save JSON Output File
    $jasonOutput | Out-File $filePath -Force
#endregion

#region Load assemblies to PowerShell session
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client")
    [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.SharePoint.Client.Runtime")
    $a = [System.Reflection.Assembly]::LoadWithPartialName("Microsoft.Online.SharePoint.Client.Tenant")
    if( !$a ){
        # Let's try to load that from default location.
        $dll = Get-ChildItem "C:\Program Files" -Recurse -Filter "Microsoft.Online.SharePoint.Client.Tenant.dll"
        $defaultPath = $dll.FullName
        $a = [System.Reflection.Assembly]::LoadFile($defaultPath)
    }
#endregion

#region Get Password for $userName and Create Context for Upload Site
    $pwd = Get-Content -Path C:\Scripts\pcPW.txt | ConvertTo-SecureString
    #context to upload site
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($siteURL)
    $cred = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName,$pwd)
    $context.Credentials = $cred
#endregion

#region Upload JSON Output File
    $list = $context.Web.Lists.GetByTitle($docLibName)
    $context.Load($list)
    $context.ExecuteQuery()

    #Upload file
    $FileStream = New-Object IO.FileStream($filePath,[System.IO.FileMode]::Open)
    $FileCreationInfo = New-Object Microsoft.SharePoint.Client.FileCreationInformation
    $FileCreationInfo.Overwrite = $true
    $FileCreationInfo.ContentStream = $FileStream
    $FileCreationInfo.URL = $uploadPath
    $Upload = $List.RootFolder.Files.Add($FileCreationInfo)
    $Context.Load($Upload)
    $Context.ExecuteQuery()
#endregion

#region Bulk Upload API
    # context to Admin Portal
    $uri = New-Object System.Uri -ArgumentList $adminUrl
    $context = New-Object Microsoft.SharePoint.Client.ClientContext($uri)
    $context.Credentials = New-Object Microsoft.SharePoint.Client.SharePointOnlineCredentials($userName, $pwd)
    $o365 = New-Object Microsoft.Online.SharePoint.TenantManagement.Office365Tenant($context)
    $context.Load($o365)

    # Type of user identifier ["Email", "CloudId", "PrincipalName"] in the User Profile Service
    $userIdType=[Microsoft.Online.SharePoint.TenantManagement.ImportProfilePropertiesUserIdType]::Email

    # Name of user identifier property in the JSON
    $userLookupKey="IdName"

    # Create property mapping between on-premises name and O365 property name
    $propertyMap = New-Object -type 'System.Collections.Generic.Dictionary[String,String]'
    $propertyMap.Add("Property1", "City")
    # $propertyMap.Add("Property2", "State")

    # Call to queue UPA property import 
    $workItemId = $o365.QueueImportProfileProperties($userIdType, $userLookupKey, $propertyMap, $uploadPath);

    # Execute the CSOM command for queuing the import job
    $context.ExecuteQuery();

    # Output unique identifier of the job
    Write-Host "Import job created with following identifier:" $workItemId.Value
#endregion