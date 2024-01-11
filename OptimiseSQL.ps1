<# Prepare-D365DevelopmentMachine
 #
 # Preparation:
 # So that the installations do not step on each other: First run windows updates, also
 # wait for antimalware to run scan...otherwise this will take a long time and we do not
 # want an automatic reboot to occur while this script is executing.
 #
 # Execute this script:
 # Set-ExecutionPolicy Bypass -Scope Process -Force; iex ((New-Object System.Net.WebClient).DownloadString('http://192.166.1.15:8000/Prepare-D365DevelopmentMachine.ps1'))
 #
 # Tested on Windows 10 and Windows Server 2016
 # Tested on F&O 7.3 OneBox and F&O 8.1 OneBox and a 10.0.11 Azure Cloud Hosted Environment (CHE) deployed from LCS
 #
 # Ideas:
 #  Download useful SQL and PowerShell scripts, using Git?
 #>

#region Installing d365fo.tools
# This is requried by Find-Module, by doing it beforehand we remove some warning messages
Set-PSRepository -Name PSGallery -InstallationPolicy Trusted

# Installing d365fo.tools
$Module2Service = $('dbatools',
    'd365fo.tools')

$Module2Service | ForEach-Object {
    if (Get-Module -ListAvailable -Name $_) {
        Write-Host "Updating " + $_
        Update-Module -Name $_ -Force
    } 
    else {
        Write-Host "Installing " + $_
        Install-Module -Name $_ -SkipPublisherCheck -Scope AllUsers
        Import-Module $_
    }
}
#endregion

Install-D365SupportingSoftware -Name "7zip" , "adobereader" , "azure-cli" , "azure-data-studio" , "azurepowershell" , "dotnetcore" , "fiddler" , "git.install", "notepadplusplus.install", "postman" , "sysinternals" , "visualstudio-codealignment" , "vscode-azurerm-tools" , "vscode-powershell" , "vscode"

Write-Host "Setting web browser homepage to the local environment"
Get-D365Url | Set-D365StartPage

Write-Host "Setting Management Reporter to manual startup to reduce churn and Event Log messages"
Get-D365Environment -FinancialReporter | Set-Service -StartupType Manual

Write-Host "Setting Windows Defender rules to speed up compilation time"
Add-D365WindowsDefenderRules -Silent

#region Local User Policy

# Set the password to never expire
Get-WmiObject Win32_UserAccount -filter "LocalAccount=True" | Where-Object { $_.SID -Like "S-1-5-21-*-500" } | Set-LocalUser -PasswordNeverExpires 1

# Disable changing the password
$registryPath = "HKCU:\Software\Microsoft\Windows\CurrentVersion\Policies\System"
$name = "DisableChangePassword"
$value = "1"

If (!(Test-Path $registryPath)) {
    New-Item -Path $registryPath -Force | Out-Null
    New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
}
Else {
    $passwordChangeRegKey = Get-ItemProperty -Path $registryPath -Name $Name -ErrorAction SilentlyContinue

    If (-Not $passwordChangeRegKey) {
        New-ItemProperty -Path $registryPath -Name $name -Value $value -PropertyType DWORD -Force | Out-Null
    }
    Else {
        Set-ItemProperty -Path $registryPath -Name $name -Value $value
    }
}

#endregion

#region Privacy

# Disable Windows Telemetry (requires a reboot to take effect)
Set-ItemProperty -Path HKLM:\SOFTWARE\Policies\Microsoft\Windows\DataCollection -Name AllowTelemetry -Type DWord -Value 0
Get-Service DiagTrack, Dmwappushservice | Stop-Service | Set-Service -StartupType Disabled

# Start Menu: Disable Bing Search Results
Set-ItemProperty -Path HKCU:\SOFTWARE\Microsoft\Windows\CurrentVersion\Search -Name BingSearchEnabled -Type DWord -Value 0


# Start Menu: Disable Cortana
If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\Personalization\Settings")) {
    New-Item -Path "HKCU:\SOFTWARE\Microsoft\Personalization\Settings" -Force | Out-Null
}
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\Personalization\Settings" -Name "AcceptedPrivacyPolicy" -Type DWord -Value 0
If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\InputPersonalization")) {
    New-Item -Path "HKCU:\SOFTWARE\Microsoft\InputPersonalization" -Force | Out-Null
}
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\InputPersonalization" -Name "RestrictImplicitTextCollection" -Type DWord -Value 1
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\InputPersonalization" -Name "RestrictImplicitInkCollection" -Type DWord -Value 1
If (!(Test-Path "HKCU:\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore")) {
    New-Item -Path "HKCU:\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore" -Force | Out-Null
}
Set-ItemProperty -Path "HKCU:\SOFTWARE\Microsoft\InputPersonalization\TrainedDataStore" -Name "HarvestContacts" -Type DWord -Value 0
If (!(Test-Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search")) {
    New-Item -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Force | Out-Null
}
Set-ItemProperty -Path "HKLM:\SOFTWARE\Policies\Microsoft\Windows\Windows Search" -Name "AllowCortana" -Type DWord -Value 0

#endregion


#region Install and run Ola Hallengren's IndexOptimize

Function Execute-Sql {
    Param(
        [Parameter(Mandatory = $true)][string]$server,
        [Parameter(Mandatory = $true)][string]$database,
        [Parameter(Mandatory = $true)][string]$command
    )
    Process {
        $scon = New-Object System.Data.SqlClient.SqlConnection
        $scon.ConnectionString = "Data Source=$server;Initial Catalog=$database;Integrated Security=true"

        $cmd = New-Object System.Data.SqlClient.SqlCommand
        $cmd.Connection = $scon
        $cmd.CommandTimeout = 0
        $cmd.CommandText = $command

        try {
            $scon.Open()
            $cmd.ExecuteNonQuery()
        }
        catch [Exception] {
            Write-Warning $_.Exception.Message
        }
        finally {
            $scon.Dispose()
            $cmd.Dispose()
        }
    }
}

If (Test-Path "HKLM:\Software\Microsoft\Microsoft SQL Server\Instance Names\SQL") {

    #Alocating 70% of the total server memory for sql server
    $totalServerMemory = Get-WMIObject -Computername . -class win32_ComputerSystem | Select-Object -Expand TotalPhysicalMemory
    $memoryForSqlServer = ($totalServerMemory * 0.7) / 1024 / 1024

    Set-DbaMaxMemory -SqlInstance . -Max $memoryForSqlServer

    Write-Host "Installing Ola Hallengren's SQL Maintenance scripts"
    Import-Module -Name dbatools
    Install-DbaMaintenanceSolution -SqlInstance . -Database master

    Write-Host "Installing FirstAidResponder PowerShell module"
    Install-DbaFirstResponderKit -SqlInstance . -Database master

    Invoke-D365InstallSqlPackage
    Invoke-D365InstallAzCopy

    Write-Host "Install latest CU"
    $DownloadPath = "C:\temp\SqlKB"
    $PathExists = Test-Path($DownloadPath)
    if ($PathExists -eq $false) {
        mkdir $DownloadPath
    }

    $BuildTargets = Test-DbaBuild -SqlInstance . -MaxBehind 0CU -Update | Where-Object { !$PSItem.Compliant } | Select-Object -ExpandProperty BuildTarget -Unique
    Get-DbaBuildReference -Build $BuildTargets | ForEach-Object { Save-DbaKBUpdate -Path $DownloadPath -Name $PSItem.KBLevel };
    Update-DbaInstance -ComputerName . -Path $DownloadPath -Confirm:$false
    Remove-Item $DownloadPath

    Write-Host "Adding trace flags"
    Enable-DbaTraceFlag -SqlInstance . -TraceFlag 174, 834, 1204, 1222, 1224, 2505, 7412

    Write-Host "Restarting service"
    Restart-DbaService -Type Engine -Force

    Write-Host "Setting recovery model"
    Set-DbaDbRecoveryModel -SqlInstance . -RecoveryModel Simple -Database AxDB -Confirm:$false

    Write-Host "Setting database options"
    $sql = "ALTER DATABASE [AxDB] SET AUTO_CLOSE OFF"
    Execute-Sql -server "." -database "AxDB" -command $sql

    $sql = "ALTER DATABASE [AxDB] SET AUTO_UPDATE_STATISTICS_ASYNC OFF"
    Execute-Sql -server "." -database "AxDB" -command $sql

    Write-Host "Setting batchservergroup options"
    $sql = "delete batchservergroup where SERVERID <> 'Batch:'+@@servername

    insert into batchservergroup(GROUPID, SERVERID, RECID, RECVERSION, CREATEDDATETIME, CREATEDBY)
    select GROUP_, 'Batch:'+@@servername, 5900000000 + cast(CRYPT_GEN_RANDOM(4) as bigint), 1, GETUTCDATE(), '-admin-' from batchgroup
        where not EXISTS (select recid from batchservergroup where batchservergroup.GROUPID = batchgroup.GROUP_)"
    Execute-Sql -server "." -database "AxDB" -command $sql

    Write-Host "purging disposable data"

$DiposableTables = @(
    "batchjobhistory"
    ,"BatchConstraintsHistory"
    ,"batchhistory"
    ,"DMFDEFINITIONGROUPEXECUTION"
    ,"DMFDEFINITIONGROUPEXECUTIONHISTORY"
    ,"DMFEXECUTION"
    ,"DMFSTAGINGEXECUTIONERRORS"
    ,"DMFSTAGINGLOG"
    ,"DMFSTAGINGLOGDETAILS"
    ,"DMFSTAGINGVALIDATIONLOG"
    ,"eventcud"
    ,"EVENTCUDLINES"
    ,"formRunConfiguration"
    ,"INVENTSUMLOGTTS"
    ,"MP.PeggingIdMapping"
    ,"REQPO"
    ,"REQTRANS"
    ,"REQTRANSCOV"
    ,"RETAILLOG"
    ,"SALESPARMLINE"
    ,"SALESPARMSUBLINE"
    ,"SALESPARMSUBTABLE"
    ,"SALESPARMTABLE"
    ,"SALESPARMUPDATE"
    ,"SUNTAFRELEASEFAILURES"
    ,"SUNTAFRELEASELOGLINEDETAILS"
    ,"SUNTAFRELEASELOGTABLE"
    ,"SUNTAFRELEASELOGTRANS"
    ,"sysdatabaselog"
    ,"syslastvalue"
)

$DiposableTables | ForEach-Object {
    Write-Host "purging $_"
    $sql = "truncate table $_"
    Execute-Sql -server "." -database "AxDB" -command $sql
}
    
    Write-Host "purging disposable batch job data"
    $sql = "delete batchjob where status in (3, 4, 8)
    delete batch where not exists (select recid from batchjob where batch.BATCHJOBID = BATCHJOB.recid)"
    Execute-Sql -server "." -database "AxDB" -command $sql

    Write-Host "purging staging tables data"
    $sql = "EXEC sp_msforeachtable
    @command1 ='truncate table ?'
    ,@whereand = ' And Object_id In (Select Object_id From sys.objects
    Where name like ''%staging'')'"

    Execute-Sql -server "." -database "AxDB" -command $sql

    Write-Host "purging disposable report data"
    $sql = "EXEC sp_msforeachtable
    @command1 ='truncate table ?'
    ,@whereand = ' And Object_id In (Select Object_id From sys.objects
    Where name like ''%tmp'')'"
    Execute-Sql -server "." -database "AxDB" -command $sql

    Write-Host "dropping temp tables"
    $sql = "EXEC sp_msforeachtable 
    @command1 ='drop table ?'
    ,@whereand = ' And Object_id In (Select Object_id FROM SYS.OBJECTS AS O WITH (NOLOCK), SYS.SCHEMAS AS S WITH (NOLOCK) WHERE S.NAME = ''DBO'' AND S.SCHEMA_ID = O.SCHEMA_ID AND O.TYPE = ''U'' AND O.NAME LIKE ''T[0-9]%'')' "
    Execute-Sql -server "." -database "AxDB" -command $sql

    Write-Host "dropping oledb error tmp tables"
    $sql = "EXEC sp_msforeachtable 
    @command1 ='drop table ?'
    ,@whereand = ' And Object_id In (Select Object_id FROM SYS.OBJECTS AS O WITH (NOLOCK), SYS.SCHEMAS AS S WITH (NOLOCK) WHERE S.NAME = ''DBO'' AND S.SCHEMA_ID = O.SCHEMA_ID AND O.TYPE = ''U'' AND O.NAME LIKE ''DMF_OLEDB_Error_%'')' "
    Execute-Sql -server "." -database "AxDB" -command $sql

    $sql = "EXEC sp_msforeachtable 
    @command1 ='drop table ?'
    ,@whereand = ' And Object_id In (Select Object_id FROM SYS.OBJECTS AS O WITH (NOLOCK), SYS.SCHEMAS AS S WITH (NOLOCK) WHERE S.NAME = ''DBO'' AND S.SCHEMA_ID = O.SCHEMA_ID AND O.TYPE = ''U'' AND O.NAME LIKE ''DMF_FLAT_Error_%'')' "
    Execute-Sql -server "." -database "AxDB" -command $sql

    $sql = "EXEC sp_msforeachtable 
    @command1 ='drop table ?'
    ,@whereand = ' And Object_id In (Select Object_id FROM SYS.OBJECTS AS O WITH (NOLOCK), SYS.SCHEMAS AS S WITH (NOLOCK) WHERE S.NAME = ''DBO'' AND S.SCHEMA_ID = O.SCHEMA_ID AND O.TYPE = ''U'' AND O.NAME LIKE ''DMF_[0-9]%'')' "
    Execute-Sql -server "." -database "AxDB" -command $sql

    Write-Host "purging disposable large tables data"
    $LargeTables | ForEach-Object {
        $sql = "delete $_ where $_.CREATEDDATETIME < dateadd(""MM"", -2, getdate())"
        Execute-Sql -server "." -database "AxDB" -command $sql
    }

    $sql = "DELETE [REFERENCES] FROM [REFERENCES]
    JOIN Names ON (Names.Id = [REFERENCES].SourceId OR Names.Id = [REFERENCES].TargetId)
    JOIN Modules ON Names.ModuleId = Modules.Id
    WHERE Module LIKE '%Test%' AND Module <> 'TestEssentials'"

    Execute-Sql -server "." -database "DYNAMICSXREFDB" -command $sql

    Write-Host "Reclaiming freed database space"
    Invoke-DbaDbShrink -SqlInstance . -Database "AxDb" -FileType Data
    Invoke-DbaDbShrink -SqlInstance . -Database "AxDb", "DYNAMICSXREFDB" -FileType Data

    Write-Host "Running Ola Hallengren's IndexOptimize tool"
    # http://calafell.me/defragment-indexes-on-d365-finance-operations-virtual-machine/
    $sql = "EXECUTE master.dbo.IndexOptimize
        @Databases = 'ALL_DATABASES',
        @FragmentationLow = NULL,
        @FragmentationMedium = 'INDEX_REORGANIZE,INDEX_REBUILD_ONLINE,INDEX_REBUILD_OFFLINE',
        @FragmentationHigh = 'INDEX_REBUILD_ONLINE,INDEX_REBUILD_OFFLINE',
        @FragmentationLevel1 = 5,
        @FragmentationLevel2 = 25,
        @LogToTable = 'N',
        @UpdateStatistics = 'ALL',
        @OnlyModifiedStatistics = 'Y'"

    Execute-Sql -server "." -database "master" -command $sql

    Write-Host "Reclaiming database log space"
    Invoke-DbaDbShrink -SqlInstance . -Database "AxDb", "DYNAMICSXREFDB" -FileType Log -ShrinkMethod TruncateOnly

    Set-DbaMaxMemory -SqlInstance . -Max 4096
}
Else {
    Write-Verbose "SQL not installed.  Skipped Ola Hallengren's index optimization"
}

#endregion


#region Update PowerShell Help, power settings

Write-Host "Updating PowerShell help"
$what = ""
Update-Help  -Force -Ea 0 -Ev what
If ($what) {
    Write-Warning "Minor error when updating PowerShell help"
    Write-Host $what.Exception
}

# Set power settings to High Performance
Write-Host "Setting power settings to High Performance"
powercfg.exe /SetActive 8c5e7fda-e8bf-4a96-9a85-a6e23a8c635c
#endregion
