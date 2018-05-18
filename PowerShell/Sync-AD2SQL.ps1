#region Variables
    $DCs = @("DC01","DC02","DC03") # Array of Domain Controller Server Names
    $dbServer = "SQL2017-01" # SQL Server Name
    $sqlInstance = "Default" # SQL Instance Name
    $databaseName = "pcDemo_Personnel_Live" # Database Name
    $activeTableName = "pcDemo_SystemUsers" # Table Name
    $saveLocation = "c:\psOutputs" # Create status file location
    $daysSaved = 6 # Days to Keep Synchronization File History
    $userFilter = "Filtered" # Get All Users ("All") or Display Name Populated Users ("Filtered")
    # AD Properties to Sync into SQL
    $properties = ("sAMAccountName","displayName","mail","telephoneNumber","physicalDeliveryOfficeName",
                   "department","userAccountControl","company","title","lastLogon","manager","givenName",
                   "Surname","StreetAddress","City","State","Country","PostalCode","SID")
#endregion

#region Check for installed modules
    if (-not (Get-WindowsFeature RSAT-AD-PowerShell).Installed) { Add-WindowsFeature RSAT-AD-PowerShell -IncludeManagementTools}    
    if (-not (Get-Module -Name SqlServer)) {
        Import-Module SqlServer
        if (-not (Get-Module -Name SqlServer)) {
            if (-not (Get-PackageProvider -Name NuGet)){
                Install-PackageProvider -Name NuGet -MinimumVersion 2.8.5.201 -Force -Confirm:$false
            }
            Set-PSRepository -Name "PSGallery" -InstallationPolicy Trusted
            Install-Module -Name SqlServer -AllowClobber -Confirm:$false -Force
            $item = Get-ChildItem -Path 'C:\Program Files\WindowsPowerShell\Modules\SqlServer' -Recurse -Filter "sqlserver.psm1"
            Import-Module $item.FullName
        }
    }
#endregion

#region Validate SQLSERVER Path else STOP
    $dbPath = "SQLSERVER:\SQL\$dbServer\$sqlInstance\Databases"
    try
    {
        Set-Location $dbPath -ErrorAction Stop
    }
    catch
    {
        Write-Output("Could not get to $dbPath... Check `$dbPath Path and Try Again...")
    }
    Set-Location $dbPath
#endregion

#region Check if DB Exists else Create It 
    $db = Get-SqlDatabase -Name $databaseName -ErrorAction SilentlyContinue 
       
    if ($db.count -lt 1) 
    {
        Write-Output("Creating Database $databaseName...")
        $query18 = "CREATE DATABASE $databaseName"
        Invoke-Sqlcmd -Query $query18 -ServerInstance $dbServer
    }
    $status = (Get-SqlDatabase -Name $databaseName -ErrorAction SilentlyContinue).Status
    do 
    {
        $status = (Get-SqlDatabase -Name $databaseName -ErrorAction SilentlyContinue).Status
        if ($status -ne "Normal") {Start-Sleep -Seconds 5}
    }
    while
    (
        $status -ne "Normal"
    )

    # Set DB Location
    Set-Location "$dbPath\$databaseName"    
#endregion

#region Create Saved File and Status File Location
    # Saved File Name
    $date = Get-Date -Format s
    $fileName = "synchedUsers_" + $date + ".txt"
    $fileName = $fileName.Replace(":","_")
    $savePath = $saveLocation.TrimEnd("\")

    # Saved File
    $file = "$savePath\$fileName"
    # Verify Folder Exists else Create It
    if (-not (Test-Path $saveLocation)) {New-Item -Path $saveLocation -ItemType Directory}
    # Create Out-File and add start time/date
    $PoSH_startTime = Get-Date
    "Synchronize AD to SQL PowerShell started: " + $PoSH_startTime | Out-File $file
#endregion

#region Function- Write Start Time
    function writeStartTime($string)
    {
        # add start time/date to outfile
        $startTime = Get-Date
        $string + " started: " + $startTime | Out-File $file -Append
    }
#endregion

#region Function- Write Stop Time
    function writeFinishTime($string)
    {
        # add finish time/date to outfile
        $endTime = Get-Date
        $string + " finished :" + $endTime | Out-File $file -Append

        # add duration time to outfile
        $queryDuration = ($endTime - $startTime).duration()
        $string + " duration: " + $queryDuration | Out-File $file -Append
    }
#endregion

#region Function- Validate Domain Controllers
    Function validateServer ($s)
    {
        $alive = $true
        if(!(Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet))
        {    
        "Problem connecting to $s" | Out-File $file -Append
        ipconfig /flushdns | Out-Null
        ipconfig /registerdns | Out-Null
        nslookup $s  | Out-Null
            if(!(Test-Connection -Cn $s -BufferSize 16 -Count 1 -ea 0 -quiet))
            {
                $alive = $false
            }
            ELSE 
            {
                "Resolved problem connecting to $s" | Out-File $file -Append
                $alive = $true
            } 
        } 
       return $alive # always a good sign!
    } 
#endregion

#region Function- Clean up data values and datatypes
    function Clean-UsersArray($users)
    {
        foreach ($user in $users)
        {
            foreach ($prop in $properties)
            {
                if ($user.$prop.Length -gt 0)
                {
                    if ($prop -eq "lastLogon")
                    {
                        if ($user.lastLogon -eq 0)
                        {
                            $user.lastLogon = $null
                        }
                        else
                        {
                            [datetime]$lastLogon = [datetime]$user.lastLogon = [datetime]::FromFileTime($user.lastLogon)
                            [datetime]$user.lastLogon = $lastLogon | Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
                        }
                    }
                    else  
                    {
                        [string]$user.$prop = $user.$prop.ToString()
                    }                                    
                }
                elseif (($user.$prop.Length -lt 1) -and ($prop -ne "lastLogon"))
                {
                    [string]$user.$prop = $null
                }
            }
            [datetime]$user.CreatedDate = Get-Date -Format 'yyyy-MM-dd HH:mm:ss.fff'
        }
        return $users
    }
#endregion

#region Function Check for $activeTableName else create it
    function Check-TableName ($activeTableName, $masterTable, $users)
    {
        $tables = (Get-SqlDatabase -Name $databaseName -ServerInstance $dbServer | select tables) 
        $table = ($tables.Tables | Where-Object {$_.Name -eq $activeTableName})
        if ($table.count -lt 1) {
            Write-Output("Creating $activeTableName Table...")
            $query20 =  "IF OBJECT_ID('$activeTableName', 'U') IS NULL "
            $query20 += "Select * INTO [$activeTableName] FROM [$masterTable];"
            Invoke-Sqlcmd -Query $query20 -Database $databaseName -ServerInstance $dbServer

            # Add Column RowID to $activeTableName
            $query21 =  "ALTER TABLE $activeTableName "
            $query21 += "ADD [RowID] [Int] IDENTITY(1,1)"
            Invoke-Sqlcmd -Query $query21 -Database $databaseName -ServerInstance $dbServer

            # Add Primary Key to $activeTableName
            $query22 =  "ALTER TABLE pcDemo_SystemUsers "
            $query22 += "ADD PRIMARY KEY (RowID)"
            Invoke-Sqlcmd -Query $query22 -Database $databaseName -ServerInstance $dbServer
        }      
    }
#endregion

#region Validate Domain Controllers
    $OUs = @() 
    foreach ($DC in $DCs)
    {
        $a = validateServer($DC)   
        if ($a)
        {
            "$DC is alive: " + $a | Out-File $file -Append
            $OUs += $DC
        }
    }
#endregion

#region Query DC(s) and Upload Data to SQL
    $counter = 0
    foreach ($OU in $OUs)
    {
        # Get current OU Server Name
        $ouServer = $OUs[$counter]

        # Create Table Name
        $tableName = "temp_" + $ouServer + "_Table"

        # Drop table if it exists
        $query1 = "IF OBJECT_ID('dbo.$tableName', 'U') IS NOT NULL DROP TABLE dbo.$tableName"
        Invoke-Sqlcmd -Query $query1 -Database $databaseName -ServerInstance $dbServer

        # add AD query start time/date to outfile
        $startTime = Get-Date
        "Query AD " + $ouServer + " started: " + $startTime | Out-File $file -Append

        # Set AD Properties to return 
        if ($counter -gt 0)
        {
            $properties = ("sAMAccountName","lastLogon")
        }

        # Get Users and their properties out of AD where the displayName is not blank
        $users = @()
        switch ($userFilter.ToLower()) 
        {
            all
            {
                $users = Get-ADUser -Filter * -Server $ouServer -Properties (foreach{$properties}) | Select (foreach{$properties})
            }
            filtered
            {
                $users = Get-ADUser -Filter {displayName -like "*"}  -Server $ouServer -Properties (foreach{$properties}) | Select (foreach{$properties})
            }
        }

        # add AD query finish time/date to outfile
        $endTime = Get-Date
        "Query AD " + $ouServer + " finished :" + $endTime | Out-File $file -Append

        # add duration time to outfile
        $queryDuration = ($endTime - $startTime).duration()
        "Query AD " + $ouServer + " duration: " + $queryDuration | Out-File $file -Append

        # add CreatedDate Column to Array
        $users | Add-Member -MemberType NoteProperty -Name CreatedDate -Value $null -Force

        # clean userdata and datatypes
        $users = Clean-UsersArray -users $users 
            
        # SQL Write start time/date to outfile
        $sqlStartTime = Get-Date
        "SQL Creation started: " + $sqlStartTime | Out-File $file -Append

        # Set-Location "SQLSERVER:\SQL\SQL2017-01\Default\Databases\pcDemo_personnel\Tables"
        $table = Write-SqlTableData -TableName $tableName -InputData $users -SchemaName "DBO" -Force

        # SQL Write finish time/date to outfile
        $sqlEndTime = Get-Date
        "SQL Creation finished :" + $sqlEndTime | Out-File $file -Append

        # add duration time to outfile
        $sqlQueryDuration = ($sqlEndTime - $sqlStartTime).duration()
        "SQL Creation duration: " + $sqlQueryDuration | Out-File $file -Append

        # check to see if $activeTableName has been created
        if ($counter -lt 1) {
            Check-TableName -activeTableName $activeTableName -users $users -masterTable $tableName
        }
        $counter ++   
    } 
#endregion

#region Move Last Logon Times to Temp Table If Multiple Domain Controllers 
    if ($OUs.Count -gt 1)
    {
        # Drop table if it exists
        $query3 = "IF OBJECT_ID('dbo.temp_lastLogonTimes', 'U') IS NOT NULL DROP TABLE dbo.temp_lastLogonTimes"
        Invoke-Sqlcmd -Query $query3 -Database $databaseName -ServerInstance $dbServer

        # Create temp_lastLogonTimes Table
        $query4 = "CREATE TABLE temp_lastLogonTimes (sAMAccountName varchar(1000))"
        Invoke-Sqlcmd -Query $query4 -Database $databaseName -ServerInstance $dbServer

        # Add a column for each OU
        foreach ($OU in $OUs)
        {
            # Create OU Columns
            $columnName = $OU + "_lastLogon"
            $query5 = "ALTER TABLE temp_lastLogonTimes ADD " + $columnName + " varchar(1000)"
            Invoke-Sqlcmd -Query $query5 -Database $databaseName -ServerInstance $dbServer
        }

        # Insert and Update Times Into Temp Table
        $counter = 0
        foreach ($OU in $OUs)
        {
            if ($counter -lt 1)
            {
                # Insert Names and Times
                $query6 = "INSERT INTO [dbo].[temp_lastLogonTimes] 
                                ([sAMAccountName]
                                ,[" + $OU + "_lastLogon])
                           Select
                                sAMAccountName 
                               ,lastLogon
                           FROM
                               temp_" + $OU + "_Table"
                Invoke-Sqlcmd -Query $query6 -Database $databaseName -ServerInstance $dbServer
            }

            # Update OU lastLogon Times *** Adjust Query Timeout Accordingly ***
            $query7 = "UPDATE [dbo].[temp_lastLogonTimes] 
                       SET " + $OU + "_lastLogon = lastLogon
                       FROM temp_" + $OU + "_Table
                       WHERE temp_lastLogonTimes.sAMAccountName = temp_" + $OU + "_Table.sAMAccountName"
            Invoke-Sqlcmd -Query $query7 -Database $databaseName -ServerInstance $dbServer
            $counter ++
        }

        # Get Max lastLogon Times 
        # Get Table and Update Last Logon Value
        $str_OUs = @()
        $str_Where = @()
        foreach ($OU in $OUs)
        {
            $str_OUs += "ISNULL(" + $OU + "_lastLogon, 0) as " + $OU + "_lastLogon"
            $str_Where += $OU + "_lastLogon <> '0'"
        }
        $str_OUs = $str_OUs -join ", "
        $str_Where = $str_Where -join " or "
    
        $query8 = "SELECT sAMAccountName, " + $str_OUs + " FROM temp_lastLogonTimes WHERE $str_Where"
        $arrayLLT = @()
        $arrayLLT = Invoke-Sqlcmd -Query $query8 -Database $databaseName -ServerInstance $dbServer
        $arrayLLT | Add-Member -MemberType NoteProperty -Name "lastLogon" -Value ""
        $arrayLength = $arrayLLT[0].Table.Columns.Count - 1

        $counter = 0
        foreach ($sAM in $arrayLLT.sAMAccountName)
        {
            $max = $arrayLLT[$counter][1..$arrayLength] | Measure -Maximum
            [datetime]$arrayLLT[$counter].lastLogon = [datetime]$max.Maximum
            $counter ++
        }

        # Drop table if it exists
        $tableNameLLT = "temp_lastLogons"
        $query9 = "IF OBJECT_ID('dbo.$tableNameLLT', 'U') IS NOT NULL DROP TABLE dbo.$tableNameLLT"
        Invoke-Sqlcmd -Query $query9 -Database $databaseName -ServerInstance $dbServer

        # Turn $users into DataTable
        $arrayLLT = $arrayLLT | Select sAMAccountName, lastLogon 

        # Set-Location "$dbPath\$databaseName\Tables"
        $table = Write-SqlTableData -TableName $tableNameLLT -InputData $arrayLLT -SchemaName "DBO" -Force
    }
#endregion

#region Update Current Users In $activeTableName
    $tempTableName = "temp_" + $OUs[0] + "_Table"
    $query11 = "UPDATE active
		        SET
	                active.sAMAccountName = LOWER(temp.sAMAccountName),
                    active.displayName = temp.displayName,
                    active.Surname = temp.Surname,
                    active.givenName = temp.givenName,
                    active.company = temp.company,
                    active.physicalDeliveryOfficeName = temp.physicalDeliveryOfficeName,
                    active.title = temp.title,
                    active.manager = temp.manager,
                    active.telephoneNumber = temp.telephoneNumber,
                    active.mail = temp.mail,
                    active.streetAddress = temp.streetAddress,
                    active.city = temp.city,
                    active.state = temp.state,
                    active.country = temp.country,
                    active.postalCode = temp.postalCode,
                    active.SID = temp.SID,
                    active.lastLogon = CONVERT(DATETIME, temp.lastLogon),
                    active.userAccountControl = temp.userAccountControl,
	                active.department = temp.department	  		 
               FROM " + $activeTableName + " active
		                inner join " + $tempTableName + " temp
		                    on active.sAMAccountName = temp.sAMAccountName
    	       WHERE LOWER(active.sAMAccountName) = LOWER(temp.sAMAccountName)"
    Invoke-Sqlcmd -Query $query11 -Database $databaseName -ServerInstance $dbServer
#endregion

#region Remove Users removed from AD
    # Check if Removed User Table exists, else create it...
    $removeUserTable = $activeTableName + "_Removed"
    $query14 =  "IF OBJECT_ID('$removeUserTable', 'U') IS NULL "
    $query14 += "Select * INTO [$removeUserTable] FROM [$activeTableName] WHERE 1 = 0;"
    Invoke-Sqlcmd -Query $query14 -Database $databaseName -ServerInstance $dbServer

    # Add Column to Track When Active User Was Moved
    $query15 =  "IF NOT EXISTS " 
    $query15 += "("
	$query15 += "SELECT * "
    $query15 += "FROM   INFORMATION_SCHEMA.COLUMNS "
    $query15 += "WHERE  TABLE_NAME = '$removeUserTable' AND COLUMN_NAME = 'MovedDate'"
	$query15 += ") "
	$query15 += "ALTER Table $removeUserTable ADD MovedDate DateTime2 NULL"
    Invoke-Sqlcmd -Query $query15 -Database $databaseName -ServerInstance $dbServer

    # Update Removed Users Database
    $query16 = "INSERT INTO [$databaseName].[dbo].[$removeUserTable]
    (
	    [sAMAccountName],
	    [displayName],
	    [givenName],
	    [Surname],
	    [company],
	    [physicalDeliveryOfficeName],
	    [department],
	    [title],
	    [manager],
	    [telephoneNumber],
	    [mail],
	    [lastLogon],
	    [userAccountControl],
        [StreetAddress],
        [City],
        [State],
        [Country],
        [PostalCode],
        [SID],
		[CreatedDate],
        [MovedDate]
    )
    SELECT 
	    [sAMAccountName],
	    [displayName],
	    [givenName],
	    [Surname],
	    [company],
	    [physicalDeliveryOfficeName],
	    [department],
	    [title],
	    [manager],
	    [telephoneNumber],
	    [mail],
	    [lastLogon],
	    [userAccountControl],
        [StreetAddress],
        [City],
        [State],
        [Country],
        [PostalCode],
        [SID],
		[CreatedDate],
        GetDate()
    FROM $activeTableName AS active
    WHERE sAMAccountName NOT IN    
    (
	    SELECT LOWER(sAMAccountName)
	    FROM $tempTableName AS temp
	    WHERE LOWER(active.sAMAccountName) = LOWER(temp.sAMAccountName)
    )"
    Invoke-Sqlcmd -Query $query16 -Database $databaseName -ServerInstance $dbServer

    # Remove Users from Active User Database
    $query17 =  "DELETE FROM $activeTableName "
    $query17 += "WHERE [sAMAccountName] = "
    $query17 += "("
    $query17 += "Select [sAMAccountName] "
    $query17 += "FROM $activeTableName AS active "
    $query17 += "WHERE sAMAccountName NOT IN "  
    $query17 += "("
	$query17 += "SELECT LOWER(sAMAccountName) "
	$query17 += "FROM $tempTableName AS temp "
	$query17 += "WHERE LOWER(active.sAMAccountName) = LOWER(temp.sAMAccountName)"
    $query17 += ")"
    $query17 += ")"
    Invoke-Sqlcmd -Query $query17 -Database $databaseName -ServerInstance $dbServer
#endregion

#region Insert New Accounts Into $activeTableName 
    $query12 = "INSERT INTO [" + $databaseName + "].[dbo].[" + $activeTableName + "]
    (
	    [sAMAccountName],
	    [displayName],
	    [givenName],
	    [Surname],
	    [company],
	    [physicalDeliveryOfficeName],
	    [department],
	    [title],
	    [manager],
	    [telephoneNumber],
	    [mail],
	    [lastLogon],
	    [userAccountControl],
        [StreetAddress],
        [City],
        [State],
        [Country],
        [PostalCode],
        [SID],
        [CreatedDate]
    )
    SELECT 
	    LOWER(sAMAccountName),
	    [displayName],
	    [givenName],
	    [Surname],
	    [company],
	    [physicalDeliveryOfficeName],
	    [department],
	    [title],
	    [manager],
	    [telephoneNumber],
	    [mail],
	    [lastLogon],
	    [userAccountControl],
        [StreetAddress],
        [City],
        [State],
        [Country],
        [PostalCode],
        [SID],
        GetDate()
    FROM " + $tempTableName + " AS temp
    WHERE sAMAccountName NOT IN
    (
	    SELECT LOWER(sAMAccountName)
	    FROM " + $activeTableName + " AS active
	    WHERE LOWER(active.sAMAccountName) = LOWER(temp.sAMAccountName)
    )"
    Invoke-Sqlcmd -Query $query12 -Database $databaseName -ServerInstance $dbServer -ConnectionTimeout 600
#endregion

#region Update lastLogon Time In $activeTableName IF more than 1 DC
    if ($OUs.Count -gt 1)
    {
            $query13 = "UPDATE [dbo].[" + $activeTableName + "] 
                       SET " + $activeTableName + ".lastLogon = temp_lastLogons.lastLogon
                       FROM temp_lastLogons
                       WHERE LOWER(temp_lastLogons.sAMAccountName) = LOWER(" + $activeTableName + ".sAMAccountName)"
            Invoke-Sqlcmd -Query $query13 -Database $databaseName -ServerInstance $dbServer 
    }

    # Write Number of People Found in AD
    "Total number of users imported from AD: " + $users.count | Out-File $file -Append

    # Clean Up Old Files
    Get-ChildItem $saveLocation -Recurse | Where {$_.LastWriteTime -lt (Get-Date).AddDays(-$daysSaved)} | Remove-Item -Force

    # add PoSH finish time/date to outfile
    $PoSH_endTime = Get-Date
    "Synchronize AD to SQL PowerShell finished: " + $PoSH_endTime | Out-File $file -Append

    # add PoSH duration time to outfile
    $queryDuration = ($PoSH_endTime - $PoSH_startTime).duration()
    "Synchronize AD to SQL PowerShell duration: " + $queryDuration | Out-File $file -Append
#endregion