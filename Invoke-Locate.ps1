<#  
.SYNOPSIS  
This script was made in the spirit of (Linux/Unix) GNU findutils' locate. "locate" and "updatedb" aliases are automatically created.

.DESCRIPTION
While the name of this script is Invoke-Locate, it actually creates two persistent aliases: locate and updatedb. A fresh index is automatically created every 6 hours, updatedb can be used force a refresh. Indexing takes anywhere from 30 seconds to 15 minutes, depending on the speed of your drives. Performing the actual locate takes about 300 milliseconds. Invoke-Locate supports both case-sensitive, and case-insensitive searches, and is case-insensitive by default.

locate queries a user-specific SQLite database prepared by updatedb (Task Scheduler) and writes file names matching the pattern to standard output, one per line. Since the back-end is SQL, SQL "LIKE" syntax can be used for the search pattern (ie % and _). Asterisks are automatically translated to % for people who are used to searching with * wildcards. So locate SQ*ite, and locate SQ%ite will return the same results.

By default, locate does not check whether files found in database still exist; locate cannot report files created after the most recent update of the relevant database.

Installing the script using -advanced option additional data collection (fullname, name, directory, size, created, lastmodified), and enables SQL querying.

.PARAMETER filename
You actually don't have to specify the -filename parameter. Just locate whatever.

.PARAMETER install
Installs script to $env:LOCALAPPDATA\locate, which allows each user to have their own secured locate database.
	- Sets persistent locate and updatedb user aliases.
	- Checks for the existence of System.Data.SQLite.dll. If it does not exist, it will be automatically downloaded to $env:LOCALAPPDATA\locate.
	  To skip this step, download System.Data.SQLite and register it to the GAC.
	- Creates the database in $env:LOCALAPPDATA\locate.
	- Creates the initial table.
	- Creates a schedule task named "updatedb cron job for user [username] (PowerShell Invoke-Locate)" that runs every 6 to 24 hours. This step may prompt for a username and password if the account does not have adequate access to create the scheduled task automatically.
	  *Note: even though an elevated SYSTEM account is used for administrator installs, home directories of other users are excluded from index.
	- Prompts users to specify if mapped drives should be indexed. Take note that mapped drives can be huge.
	- Prompts user to run updatedb for the first time.

Install can be run repeatedly with no issues. To uninstall, delete $env:LOCALAPPDATA\locate, the Scheduled Task, and the two aliases within $profile.

.PARAMETER advanced
The advanced switch tells locate to index: fullname, name, directory, size, created, lastmodified. Because updatedb will take about 20% longer, Task Scheduler will execute updatedb every 12 hours instead of 6. Installing with the advanced options will allow you to perform advanced queries on the database.

.PARAMETER s
Similar to findutils locate's "-i" switch for case-insensitive, this switch makes the search sensitive. By default, Windows searches are insensitive, so the default search behavior of this script is case-insensitive.

.PARAMETER sql
This parameter will allow you to perform SQL queries directly to your SQLite database. Default installs only allow "SELECT fullname..." Advanced installs also allow "SELECT fullname, name, directory, size, created, lastmodified..."

.PARAMETER where
This parameter will allow you to perform a WHERE on the default select. Ignored if you specify -sql.

.PARAMETER orderby
This parameter will allow you to perform an ORDER BY by on the default select. Ignored if you specify -sql.

.PARAMETER Descending
Adds descending sort order to -orderby.

.PARAMETER columns
This parameter will allow you to return specific columns from a dataset by on the default select. Columns names are populated from database. Ignored if you specify -sql.

.PARAMETER limit
Limits returned rows by specified numbers

.PARAMETER du
When Invoke-Locate.ps1 is installed with the -advanced switch, a Disk Usage (du) alias is installed. This command displays disk usage output for the current directory, including files and directories.

.PARAMETER topdirectories
When used in conjunction with -du, -topdirecotries displays a human-readable disk usage report of largest directories. Beware, you'll be tempted to run this command on C:\ or C:\Program Files. I wouldn't.

.PARAMETER updatedb
Internal parameter. You don't need to pass -updatedb. Just use the updatedb alias.

.PARAMETER includemappeddrives
Internal parameter. Tells updatedb to include mapped drives.

.PARAMETER locatepath
Internal parameter. Specifies locate's program directory.

.PARAMETER userprofile
Internal parameter. This helps support mulit-user scheduled task updates. 

.PARAMETER homepath
Internal parameter. This helps support mulit-user scheduled task updates.

.NOTES  
Author  : Chrissy LeMaire 
Requires:     PowerShell Version 3.0
DateUpdated: 2015-May-12
Version: 0.8
 
.LINK
https://gallery.technet.microsoft.com/scriptcenter/Invoke-Locate-PowerShell-0aa2673a

.EXAMPLE
.\Invoke-Locate.ps1 -install
Copies necessary files, adds aliases, sets up updatedb Scheduled Task to run every 6 hours, prompts user to populate database.
.EXAMPLE
.\Invoke-Locate.ps1 -install -advanced
Copies necessary files, adds aliases, sets up updatedb Scheduled Task to run only once a day, and adds additional fields to the database.
.EXAMPLE
locate powershell.exe
Case-insensitive search which return the path to any file or directory named powershell.exe	
.EXAMPLE
updatedb
Forces a database refresh. This generally takes just a few minutes, unless you've specified -advanced during the install.
.EXAMPLE
locate power*.exe
Case-insensitive search which return the path to any file or directory that starts with power and has exe after csv in the path.
.EXAMPLE
locate -s System.Data.SQLite
Case-sensitive search which return the path to any file or directory named System.Data.SQLite.
.EXAMPLE
locate powers_ell.exe
Similar to SQL's "LIKE" syntax, underscores are used to specify "any single character."
.EXAMPLE
locate .iso -columns name, mb, gb  -where { gb -gt 1 } -orderby size
Searches for iso files larger than 1 gigabyte and orders the results by size. Returns only specified columns. Note that the -columns and -orderby columns are auto-populated so you do not have to guess.
.EXAMPLE du
When Invoke-Locate.ps1 is installed with the -advanced switch, a Disk Usage (du) alias is installed. This command displays disk usage output for the current directory, including files and directories.
.EXAMPLE du -topdirectories C:\inetpub
Aggregates disk usage information by directory for C:\inetpub

#> 
#Requires -Version 3.0
[CmdletBinding(DefaultParameterSetName="Default")] 

Param(
	[parameter(Position=0)]
	[string]$filename,
	[switch]$install,
	[switch]$updatedb,
	[string]$locatepath,
	[string]$userprofile,
	[string]$homepath,
	[switch]$s,
	[switch]$includemappedrives,
	[string]$sql,
	[string]$where,
	[int]$limit,
	[switch]$advanced,
	[switch]$recurse,
	[switch]$du,
	[switch]$topdirectories,
    [switch]$descending
	)
	
DynamicParam  {
	$database = "$env:LOCALAPPDATA\locate\locate.sqlite"

	if (!(Test-Path $database)) { return }

	try {
		if ([Reflection.Assembly]::LoadWithPartialName("System.Data.SQLite") -eq $null) { 
		[void][Reflection.Assembly]::LoadFile("$env:LOCALAPPDATA\locate\System.Data.SQLite.dll") }
	} catch { return } 
	
	# Setup connect
	try {
		$connString = "Data Source=$database"
		$connection = New-Object System.Data.SQLite.SQLiteConnection($connString) 
		$connection.Open()
		$command = $connection.CreateCommand()
		$command.CommandText = "PRAGMA table_info(files);"
		$datatable = New-Object System.Data.DataTable
		$datatable.load($command.ExecuteReader())
		$columnarray = $datatable.name
		$command.Dispose()
		$connection.Close()
		$connection.Dispose()
	}
	catch { return }

	# Parameter setup
	$newparams = New-Object System.Management.Automation.RuntimeDefinedParameterDictionary
	$attributes = New-Object System.Management.Automation.ParameterAttribute
	$attributes.ParameterSetName = "__AllParameterSets"
	$attributes.Mandatory = $false
	
	# Do it
	if ($columnarray) { $colsvalidationset = New-Object System.Management.Automation.ValidateSetAttribute -ArgumentList $columnarray }
	$colsattributes = New-Object -Type System.Collections.ObjectModel.Collection[System.Attribute]
	$colsattributes.Add($attributes)
	if ($columnarray) { $colsattributes.Add($colsvalidationset) }
	$columns = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("columns", [string[]], $colsattributes)
	$orderby = New-Object -Type System.Management.Automation.RuntimeDefinedParameter("orderby", [string[]], $colsattributes)
	
	$newparams.Add("columns", $columns)
	$newparams.Add("orderby", $orderby)
	return $newparams
	
}
	
BEGIN {
	Function Install-Locate   {
		<#
		.SYNOPSIS
		  Installs Invoke-Locate.ps1 to the current user's $env:localappdata.
		#>
		
		param(
            [bool]$noprompt,
			[string]$locatepath,
			[bool]$advanced
		)
		
		if ($locatepath.length -eq 0) { $locatepath = "$env:LOCALAPPDATA\locate" }
		
		# Create locate's program directory within user $env:localappdata.
		if (!(Test-path $locatepath)) { $null = New-Item $locatepath -Type Directory }
		
		# Copy the files to the new directory
		$script = "$locatepath\Invoke-Locate.ps1"
		try { Get-Content $PSCommandPath | Set-Content $script } catch { } # whatever
		
		# Set persistent aliases by writing to $profile
		Write-Host "Setting persistent locate and updatedb aliases" -ForegroundColor Green
		$locatealias = "New-Alias -name locate -value ""$script"" -scope Global -force"
		$exists = Test-Path ALIAS:locate
		if ($exists -eq $false) { 
			if (!(Test-Path $profile)) {
				$profiledir = Split-Path $profile
				If (!(Test-Path $profiledir)) { $null = New-Item $profiledir -Type Directory }			
			} 
		Add-Content $profile $locatealias
		Invoke-Expression $locatealias
		} else { Write-Warning "Alias locate exists. Skipping." }
		
		<# 
			Prompt user to see if they want to index mapped drives
		#>
		
		$message = "This script can index mapped drives, too."
		$question = "Would you like to index your mapped drives?"
		$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
		$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
		$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
		
		$decision = $Host.UI.PromptForChoice($message, $question, $choices, 1)
		if ($decision -eq 1) { 
			Write-Host "Mapped drives will not be indexed."
			 $includemappedrives = $false
		} else { 
			Write-Host "Mapped drives will be indexed."
			$includemappedrives = $true }

		<# 
			Setup updatedb alias
		#>
		
		# Aliases don't allow params, so a new file must be created in localappdata to support using the alias updatedb
		$updatefilename = "$locatepath\Update-LocateDB.ps1"
		if ($homepath.length -eq 0) { $homepath = "$env:HOMEDRIVE$env:HOMEPATH" }
		if ($userprofile.length -eq 0) { $userprofile = $env:USERPROFILE } 
		if (!$includemappedrives) { $mapped = ' -includemappedrives:`$false ' } else  { $mapped = ' -includemappedrives:`$true ' }
		if (!$advanced) { $advancedfeatures = ' -advanced:`$false ' } else  { $advancedfeatures = ' -advanced:`$true ' }
		
		$updatescript +=  "Invoke-Expression "
		$updatescript += """& """"$script"""""
		$updatescript += " -updatedb -locatepath """"$locatepath"""" -homepath """"$homepath"""" -userprofile """"$userprofile"""""
		$updatescript += $mapped 
		$updatescript += $advancedfeatures 
		$updatescript += '"'

		Set-Content  $updatefilename  $updatescript
		Add-Content  $updatefilename '$computername = "$($env:COMPUTERNAME)`$".ToUpper()'
		Add-Content  $updatefilename '$username = "$env:USERNAME".ToUpper()'
		# Returns $true for Scheduled Tasks, so that they do not show up as failed.
		Add-Content  $updatefilename 'if ($computername -eq $username) { return $true }'
		Add-Content  $updatefilename 'exit'	
		
		# Add persistent updatedb alias
		$updatealias = "New-Alias -name updatedb -value ""$updatefilename"" -scope Global -force"
		$exists = Test-Path ALIAS:updatedb
		if ($exists -eq $false) { Add-Content $profile $updatealias; Invoke-Expression $updatealias }
		else { Write-Warning "Alias updatedb exists. Skipping." }
		
		<# 
			Set du alias
		#>
		if ($advanced) {
			# Aliases don't allow params, so a new file must be created in localappdata to support using the alias du
			$diskusagefilename = "$locatepath\Get-DiskUsage.ps1"
			$getdu = (Get-Command Get-DiskUsage).Definition #<- Awesome capability!
			$getdu = $getdu.Replace('[bool]','[switch]')
			Set-Content  $diskusagefilename  $getdu
			
			Write-Host "Setting persistent du alias" -ForegroundColor Green
			# Add persistent updatedb alias
			$diskusagealias = "New-Alias -name du -value ""$diskusagefilename"" -scope Global -force"
			$exists = Test-Path ALIAS:du
			if ($exists -eq $false) { Add-Content $profile $diskusagealias; Invoke-Expression $diskusagealias }
			else { Write-Warning "Alias du exists. Skipping." }
		}
		
		# Download the DLL if System.Data.SQLite cannot be found. Copy to $env:localappdata.
		$sqlite = "System.Data.SQLite"
		$globalsqlite = [Reflection.Assembly]::LoadWithPartialName($sqlite)
		if ($globalsqlite -eq $null -and !(Test-Path("$locatepath\$sqlite.dll")) ) {
			# Check architecture
			if (!(Test-Path "locatepath\$sqlite.dll")) {
				Write-Host "Downloading $sqlite.dll" -ForegroundColor Green
				if ($env:Processor_Architecture -eq "x86")   { $url = "http://bit.ly/sqlitedllx86Net35" } else {  $url = "http://bit.ly/sqlitedllx64net35"  }
					try { 
						Invoke-WebRequest $url -OutFile "$locatepath\$sqlite.dll"
						Unblock-File "$locatepath\$sqlite.dll"
					} catch {
						Remove-Item "$locatepath\$sqlite.dll"
						throw "The SQLite DLL cannot be automatically downloaded and loaded. Please try again or install SQLite (http://bit.ly/1CPVDsP), and register it to the GAC. Quitting."; 

					}
				}
		}
		try { 
		if ($globalsqlite -eq $null) { [void][Reflection.Assembly]::LoadFile("$locatepath\System.Data.SQLite.dll") } 
		} catch { throw "The SQLite DLL cannot be loaded. Do you have .NET 3.5+ installed? Please try again or install SQLite (http://bit.ly/1CPVDsP), and register it to the GAC. Quitting."; break }
		
		# Setup connstring
		$database = "$locatepath\locate.sqlite"
		$connString = "Data Source=$database"
		
		#Create the database if it doesn't exist.
		if (!(Test-Path $database)) {
			Write-Host "Creating database" -ForegroundColor Green
			# Create database
			[void][System.Data.SQLite.SQLiteConnection]::CreateFile($database); 
			$connection = New-Object System.Data.SQLite.SQLiteConnection($connString)
			$connection.Open()
			$connection.Close()
			$connection.Dispose()
		} else { Write-Warning "database exists. Skipping." }
		
		# Create scheduled task. This scheduled task will run updatedb every 6 hours,
		# as an elevated SYSTEM account. By default, the script wil not index any home directories, other than the user who installed it.	
		if ($advanced -eq $true) { Write-Host "Setting up Scheduled Task to run once a day at 12:00 am." -ForegroundColor Green }
		else { Write-Host "Setting up Scheduled Task to run every 6 hours" -ForegroundColor Green }
		
		$null = New-LocateScheduledTask -locatepath $locatepath -advanced $advanced
		
		if ($noprompt -ne $true) {
			Write-Warning "The database must be populated before it will return any results."
			$message = $null
			$question = "Would you like to run updatedb to populate the database now?"

			$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
			$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
			$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
			
			$decision = $Host.UI.PromptForChoice($message, $question, $choices, 0)
			if ($decision -eq 1) { 
				Write-Host "updatedb skipped. locate may return no results, or may be out of date." -ForegroundColor Red -BackgroundColor Black
			} else { Update-LocateDB -locatepath $locatepath -homepath $homepath, -userprofile $userprofile -includemappedrives $includemappedrives -advanced $advanced }
		} else { Update-LocateDB -locatepath $locatepath -homepath $homepath -userprofile $userprofile -includemappedrives $includemappedrives -advanced $advanced }
		
		# Finish up
		Write-Host "Installation to $locatepath complete." -ForegroundColor Green 
}

	Function New-LocateScheduledTask  {
		<#
		.SYNOPSIS
		 Creates a new scheduled task in Windows 7 and below. This scheduled task is run as an elevated SYSTEM account, but will only search the home directory
		 of the user that installed locate. Supports multiple users. Administrator access required because it uses a system account.
		#>
		
		param(
			[Parameter(Mandatory = $false)]
            [string]$locatepath,
			[bool]$advanced
		)

		# Get locate program directory
		if ($locatepath -eq $null) { $locatepath = "$env:LOCALAPPDATA\locate" }
		
		# Name the task, and check to see if it exists, if so, skip.
		$taskuser = ($env:USERNAME).Replace("'","")
		$taskname = "updatedb cron job for user $taskuser (PowerShell Invoke-Locate)" 

		# Check if admin. 
		If (-NOT ([Security.Principal.WindowsPrincipal] [Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole] "Administrator")) {
			Write-Warning "Not an admin, can't automatically create Scheduled Task without username and password."
			$message = "Would you like the script to automatically create a scheduled task now?"
			$question = "You wil be prompted for your credentials."

			$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
			$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
			$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
			
			$decision = $Host.UI.PromptForChoice($message, $question, $choices, 0)
			if ($decision -eq 0) { 
				$credential = Get-Credential -Username "$env:USERDOMAIN\$env:USERNAME" -Message "Please enter your username and password."
				$domain =  $credential.GetNetworkCredential().domain
				$user = $credential.GetNetworkCredential().username
				$username = "$domain\$user"
			} else {
				Write-Warning "You must setup the Scheduled Task manually. Please see http://bit.ly/1zHldCU for more information."
				return
			}
			$isadmin = $false
			
		} else {  $isadmin = $true; $username = "NT AUTHORITY\SYSTEM" }
		
		# Script to execute
		$updatefilename = """$locatepath\Update-LocateDB.ps1"""
		$taskscheduler = New-Object -ComObject Schedule.Service
		$taskscheduler.Connect()
		
		# Place Task in root
		$rootfolder = $taskscheduler.GetFolder("\")
		$definition = $taskscheduler.NewTask(0)
		
		# Get base info
		$registrationInformation = $definition.RegistrationInfo
		
		# Run as built in
		$principal = $definition.Principal
		if ($isadmin) { $principal.LogonType = 5 } else { $principal.LogonType = 1 } 
		$principal.UserID = $username
		$principal.RunLevel = 0
		
		# Set options
		$settings = $definition.Settings
		$settings.StartWhenAvailable = $true
		$settings.RunOnlyIfNetworkAvailable = $false
		$settings.ExecutionTimeLimit =  "PT1H"
		$settings.RunOnlyIfIdle = $false
		$settings.AllowDemandStart = $true
		$settings.AllowHardTerminate = $true
		$settings.DisallowStartIfOnBatteries = $false
		$settings.Priority = 7
		$settings.StopIfGoingOnBatteries = $false
		$settings.idlesettings.StopOnIdleEnd = $false
		
		# Set script to run every 6 hours, indefinitely. 
		if ($advanced -eq $true) { $repeat = "P0DT24H0M0S" } else { $repeat = "P0DT6H0M0S" }
		$triggers = $definition.Triggers
		$trigger = $triggers.Create(2)
		$trigger.Repetition.Interval = $repeat
		$trigger.Repetition.StopAtDurationEnd = $false
		$trigger.StartBoundary = (Get-Date "00:00:00" -Format s)
		
		# Set the action to have powershell.exe call a script.
		$action = $definition.Actions.Create(0)
		$action.Path = "powershell.exe"
		$action.Arguments =  "-File $updatefilename"
		$action.WorkingDirectory = $locatepath

		# 6 = update or delete, 0 is no password needed
		if ($isadmin) {
			try { 
				$null = $rootfolder.RegisterTaskDefinition($taskname, $definition, 6, $username, $null, 0) 
				Write-Host "Scheduled Task successfully created." -ForegroundColor Green
				} catch { Write-Warning "Task registration failed. You must setup the Scheduled Task manually. Please see http://bit.ly/1zHldCU for more information."}
		} else { 
			try { 
				$null = $rootfolder.RegisterTaskDefinition($taskname, $definition, 6, $username, $credential.GetNetworkCredential().password, 1)
				Write-Host "Scheduled Task successfully created." -ForegroundColor Green
			} catch { Write-Warning "Task registration failed. You must setup the Scheduled Task manually. Please see http://bit.ly/1zHldCU for more information."}
		}
	}
	
	Function Update-LocateDB {
		<#
		.SYNOPSIS
		  Updates the SQLite database at $env:LOCALAPPDATA\locate\locate.sqlite using a single transaction.
		#>
		param(
            [string]$locatepath,
			[string]$homepath,
			[string]$userprofile,
			[bool]$includemappedrives,
			[bool]$advanced
		) 

		Write-Host "Updating locate database" -ForegroundColor Green
		if ($advanced -eq $false) { Write-Host "This should only take a few minutes." -ForegroundColor Green }
		
		# Set variables and load up assembly
		if ($locatepath.length -eq 0) {$locatepath = "$env:LOCALAPPDATA\locate" }
		if ($homepath.length -eq 0) { $homepath = "$env:HOMEDRIVE$env:HOMEPATH" }
		if ($userprofile.length -eq 0) { $userprofile = $env:USERPROFILE } 
		if ([Reflection.Assembly]::LoadWithPartialName("System.Data.SQLite") -eq $null) { [void][Reflection.Assembly]::LoadFile("$locatepath\System.Data.SQLite.dll") }
		$elapsed = [System.Diagnostics.Stopwatch]::StartNew() 
		
		Remove-Item "$locatepath\*populat*" -Force -ErrorAction SilentlyContinue
		$populatedb = "$locatepath\locate-populating.sqlite"
		$livedb = "$locatepath\locate.sqlite"
		$connString = "Data Source=$populatedb"
		$connection = New-Object System.Data.SQLite.SQLiteConnection($connString)
		$connection.Open()
		$command = $connection.CreateCommand()
		
		# SQLite doesn't support truncate, let's just drop the table and add it back.

		if ($advanced -eq $true) {
			$command.CommandText = "DROP VIEW IF EXISTS files; DROP TABLE IF EXISTS f"
			[void]$command.ExecuteNonQuery()		
			$createsql = "CREATE TABLE f (fullname NVARCHAR(260) PRIMARY KEY, name NVARCHAR(260), directory NVARCHAR(260), size REAL, created DATETIME, lastmodified DATETIME)"
			$createsql += ";CREATE VIEW files AS SELECT fullname, name, directory, size, created, lastmodified, 
							round(size/1024,2) as kb, round(size/1048576,2) as mb, round(size/1073741824,2) AS gb, round(size/1099511627776,2) as tb FROM f"	
		} else {
			$command.CommandText = "DROP TABLE IF EXISTS files"
			[void]$command.ExecuteNonQuery()
			$createsql = "CREATE TABLE files (fullname NVARCHAR(260) PRIMARY KEY)" 
			$command.CommandText = $createsql
		}
		$command.CommandText = $createsql
		[void]$command.ExecuteNonQuery()
		
		# Use a single transaction to speed up insert.
		$transaction = $connection.BeginTransaction()
		
		# Get local drives. Like GNU locate, this includes your local DVD-CDROM, etc drives.
		$disks = Get-WmiObject Win32_Volume -Filter "Label!='System Reserved'"
		
		foreach ($disk in $disks.name) {
			Get-Filenames -path $disk -locatepath $locatepath -advanced $advanced
			$diskcount++
		}
		
		# Since C:\Users is ignored by default in the above routine, $homepath and $userprofile must be explicitly indexed.
		Get-Filenames -path $homepath -locatepath $locatepath -advanced $advanced
		if ($homepath -ne $userprofile) { Get-Filenames -path $userprofile -locatepath $locatepath -advanced $advanced }
		
		# When locate was installed, the user was prompted to answer whether they wanted to index their mapped drives.
		If ($includemappedrives -eq $true) {
			$disks = Get-WmiObject Win32_MappedLogicalDisk
			foreach ($disk in $disks.name) {
				$diskcount++
				Get-Filenames -path $disk -locatepath $locatepath -advanced $advanced
			}
		}
		
		# Commit the transaction
		$transaction.Commit()
		
		# Count the number of files indexed and report
		$totaltime = [math]::Round($elapsed.Elapsed.TotalMinutes,2)
		$totaltime = (($elapsed.Elapsed.ToString()).Split("."))[0]
		$command.CommandText = "SELECT COUNT(*) FROM files"
		$rowcount = $command.ExecuteScalar()
		
		Write-Host "$rowcount files on $diskcount drives have been indexed in $totaltime." -ForegroundColor Green
		$command.Dispose()
		$connection.Close()
		$connection.Dispose()
		
		Move-Item $populatedb $livedb -force -ErrorAction SilentlyContinue
		if (Test-Path  $populatedb) {
			Remove-Item $populatedb -force
			Write-Warning "Temporary db could not overwrite the live database. Results will be out of date until the next update."
		}
	}
	
	Function Get-Filenames {
		<#
		.SYNOPSIS
		 This function is called recursively to get filenames and insert them into the database. Skips 
		 $env:APPDATA, $env:LOCALAPPDATA, $env:TMP, $env:TEMP.
		 
		 The system drive's Users directory is also excluded, but then the locate user's homepath and userprofile
		 are explicitly included.
		 
		#>
		
		param(
			[string]$path,
            [string]$locatepath,
			[string]$homepath,
			[string]$userprofile,
			[bool]$advanced
		) 
		
		# Set variables and load SQLite assembly
		if ($locatepath -eq $null) { $locatepath = "$env:LOCALAPPDATA\locate" }
		if ([Reflection.Assembly]::LoadWithPartialName("System.Data.SQLite") -eq $null) { [void][Reflection.Assembly]::LoadFile("$locatepath\System.Data.SQLite.dll") }
		
		# IO.Directory throws a lot of access denied exceptions, ignore them.
		Set-Variable -ErrorAction SilentlyContinue -Name files
		Set-Variable -ErrorAction SilentlyContinue -Name folders
		
		# Get the directories, and make a list of the files within them
		try {
			$directoryInfo = New-Object IO.DirectoryInfo($path)
			$folders = $directoryInfo.GetDirectories() | Where-Object {$_.Name -ne "`$Recycle.Bin" -and $folder -ne "System Volume Information" }
			
			# Get & sanitize directory info for advanced queries
			$directoryname = $directoryInfo.FullName
			$directoryname = $directoryname.replace('\\','\')
			$directoryname = $directoryname.replace("'","''")
		} catch { $folders = @()}
	
		if ($advanced -eq $false) {
			try {
				$files = [IO.Directory]::GetFiles($path)
				# For each file, clean up the SQL syntax and insert into database.
				foreach($filename in $files) 
					{
						$filename = $filename.replace('\\','\')
						$filename = $filename.replace("'","''")
						$command.CommandText = "insert into files values ('$filename')"
						[void]$command.ExecuteNonQuery()
					}
				} catch {} # Access Denied
		} else {
			try { 
				$files = $directoryInfo.GetFileSystemInfos().GetEnumerator()
				# For each file, clean up the SQL syntax and insert into database.
				foreach($file in $files) {	
					$filename = $file.fullname.ToString()			
					if ($filename.length -lt 260) {
						$filename = $filename.replace('\\','\')
						$filename = $filename.replace("'","''")	
						$name = $file.name.replace('\\','\')
						$name = $name.replace("'","''")	
						$lastmodified = $file.LastWriteTime.ToString("yyyy-MM-dd HH:mm:ss")
						$created = $file.CreationTime.ToString("yyyy-MM-dd HH:mm:ss")
						$filesize = $file.length
						$insertsql = "insert into f (fullname, name, directory, size, created, lastmodified) values ('$filename','$name','$directoryname',$filesize, '$created', '$lastmodified')"
						$command.CommandText = $insertsql
						[void]$command.ExecuteNonQuery()
					}
				}
			} catch {}
		}

		# Process folders and subfolders
		foreach($folder in $folders)
		{ 
			if ("$env:systemdrive\Users" -ne "$path$folder") { 
				Get-Filenames -path "$path\$folder" -locatepath $locatepath  -advanced $advanced
				Write-Verbose "Indexing $path\$folder"
			}
		}
		
		# Remove the erroraction variable
		Remove-Variable -ErrorAction SilentlyContinue -Name files
		Remove-Variable -ErrorAction SilentlyContinue -Name folders 
	}
	
	Function Show-SQLcolumns {
		<#
		.SYNOPSIS
		 Shows the columns of a table
		 
		 .OUTPUT
		 Display text
		
		#>
		
		param(
			[object]$connection,
			[string]$tablename
		) 
			$columncommand = $connection.CreateCommand()
			$columncommand.CommandText = "PRAGMA table_info($tablename);"
			$sqlcolumns = New-Object System.Data.DataTable
			$sqlcolumns.load($columncommand.ExecuteReader())
			$sqlcolumns = $sqlcolumns.name -join " "
			Write-Host "Valid column name(s): $sqlcolumns" -BackgroundColor Black -ForegroundColor Red
	}

	Function Get-DiskUsage  {
		<#
		.SYNOPSIS
		 Linux du-ish
		
		#>
		
		param(
			[parameter(Position=0,Mandatory=$false)]
			[string]$path,
            [string]$locatepath,
			[int]$limit,
			[string]$orderby,
			[bool]$descending,
			[bool]$topdirectories
		) 

		# Get variables, load assembly
		$currentdir = (Get-Location).Path
		if ($locatepath.length -eq 0) { $locatepath = "$env:LOCALAPPDATA\locate" }
		if ($path.length -eq 0) { $path = $currentdir }
		if ($path -notlike '?:\*') { $path = $path.TrimStart(".\"); $path = "$currentdir\$path"  }

		if ([Reflection.Assembly]::LoadWithPartialName("System.Data.SQLite") -eq $null) { [void][Reflection.Assembly]::LoadFile("$locatepath\System.Data.SQLite.dll") }
		
		# Setup connect
		$database = "$locatepath\locate.sqlite"
		$connString = "Data Source=$database"

		try { $connection = New-Object System.Data.SQLite.SQLiteConnection($connString) }
		catch { throw "Can't load System.Data.SQLite.SQLite. Architecture mismatch or access denied. Quitting." }
		$connection.Open()
		$command = $connection.CreateCommand()
		
		# Clean up
		$path = $path.Replace("'","''")
		$path = $path.TrimEnd("\")
		
		$directory = (Get-Item $path) -is [System.IO.DirectoryInfo]
		if ($topdirectories) { $column = "directory" } else { $column = "fullname" }		
		if ($directory) { 
			$where = "where $column LIKE '$path\%' or $column COLLATE NOCASE = '$path'" 
		} else { $where = "where $column COLLATE NOCASE = '$path'" }
		
		$pragma = "PRAGMA case_sensitive_like = 0; "
		if ($topdirectories)  { $select = "select $column as name, sum(kb) as kb, sum(mb) as mb, sum(gb) as gb" } 
		else { $select = "select $column as name, round(kb,1) as kb, round(mb,1) as mb, round(gb,1) as gb" }
		
		$sql = "$pragma $select from files $where"
	
		if ($topdirectories)  { $sql += " GROUP BY $column " }
		if ($orderby.length -gt 0) { 
			$sql += " order by $orderby $sortorder" 
			if ($descending) { $sql += " DESC" }
		}
		if ($limit -gt 0) { $sql += " LIMIT $limit"}
		
		Write-Verbose "SQL string executed: $sql"
		$command.CommandText = $sql.Trim()
		
		# Create datatable and fill it with results
		$datatable = New-Object System.Data.DataTable
		try { $datatable.load($command.ExecuteReader()) }
		catch {
			$msg = $_.Exception.InnerException.Message.ToString() -replace "`r`n",". "
			Write-Host $msg -BackgroundColor Black -ForegroundColor Red
			$columncommand = $connection.CreateCommand()
			$columncommand.CommandText = "PRAGMA table_info('files');"
			$sqlcolumns = New-Object System.Data.DataTable
			$sqlcolumns.load($columncommand.ExecuteReader())
			$sqlcolumns = $sqlcolumns.name -join " "
			Write-Host "Valid column name(s): $sqlcolumns" -BackgroundColor Black -ForegroundColor Red
			return
		}
		$command.Dispose()
		$connection.Close()
		$connection.Dispose()
				
		$null = $datatable.Columns.Add("totalkb",[int64])
		$null = $datatable.Columns.Add("totalmb",[int64])
		$null = $datatable.Columns.Add("totalgb",[int64])
		
		foreach ($row in $datatable.rows) {
			try {
				$filename = $row.name 
				$filename = $row.name.replace("'","''")
				$where = "name like '$filename\*' or name = '$filename'"
				$row["totalkb"] = ($datatable.Compute("sum(kb)",$where)) 
				$row["totalmb"] = ($datatable.Compute("sum(mb)",$where))
				$row["totalgb"] = ($datatable.Compute("sum(gb)",$where))
			} catch { Write-Warning "Could not parse $filename info." }
		}
		
		$totalsize = $datatable.Columns.Add("totalsize")
		$totalsize.Expression = "IIF(totalkb<1025, totalkb + 'K', IIF(totalmb<1025, totalmb + 'M', totalgb + 'G'))"
		
		$datatable | Select totalsize, name | Sort-Object name | Format-Table -Auto -HideTableHeaders
	}
	
	Function Search-Filenames  {
		<#
		.SYNOPSIS
		 Performs a LIKE query on the SQLite database. 
		 
		 .OUTPUT
		 System.Data.Datatable
		
		#>
		
		param(
			[string]$filename,
            [string]$locatepath,
			[bool]$s,
			[string]$sql,
			[string[]]$columns,
			[string]$where,
			[string]$orderby,
			[bool]$descending
		) 
		
		# Get variables, load assembly
		if ($locatepath -eq $null) { $locatepath = "$env:LOCALAPPDATA\locate" }
		if ([Reflection.Assembly]::LoadWithPartialName("System.Data.SQLite") -eq $null) { [void][Reflection.Assembly]::LoadFile("$locatepath\System.Data.SQLite.dll") }
		
		# Setup connect
		$database = "$locatepath\locate.sqlite"
		$connString = "Data Source=$database"
		try { $connection = New-Object System.Data.SQLite.SQLiteConnection($connString) }
		catch { throw "Can't load System.Data.SQLite.SQLite. Architecture mismatch or access denied. Quitting." }
		$connection.Open()
		$command = $connection.CreateCommand()
		
		# Allow users to use * as wildcards and ? as single characters.
		$filename = $filename.Replace('*','%')
		$filename = $filename.Replace('?','_')
		# Escape SQL string
		$filename = $filename.Replace("'","''")

		if ($columns.length -eq 0) { $columns = "*" } else { $columns = $columns -join ", " }

		if ($sql.length -eq 0) {
			if ($s -eq $false) {
				$sql = "PRAGMA case_sensitive_like = 0;select $columns from files where fullname like '%$filename%'"
			} else { $sql = "PRAGMA case_sensitive_like = 1;select $columns from files where fullname like '%$filename%'" }
		}
		
		if ($where.length -gt 0) {
			$where = $where.Replace(" -eq "," = ")
			$where = $where.Replace(" -ne "," != ")
			$where = $where.Replace(" -gt "," > ")
			$where = $where.Replace(" -lt "," < ")
			$where = $where.Replace(" -ge "," >= ")
			$where = $where.Replace(" -le "," <= ")
			$where = $where.Replace(" -and "," and ")
			$where = $where.Replace(" -or "," or ")
			$sql += " and $where"
		}
		
		if ($orderby.length -gt 0) {
			$sql += " order by $orderby"
			if ($descending) { $sql += " DESC" }
		}

		if ($limit -gt 0) {
			$sql += " LIMIT $limit"
		}
		
		Write-Verbose "SQL string executed: $sql"
		$command.CommandText = $sql.Trim()
		
		# Create datatable and fill it with results
		$datatable = New-Object System.Data.DataTable
		try { $datatable.load($command.ExecuteReader()) }
		catch {
			$msg = $_.Exception.InnerException.Message.ToString() -replace "`r`n",". "
			Write-Host $msg -BackgroundColor Black -ForegroundColor Red
			Show-SQLcolumns -connection $connection -tablename files
		}
		$command.Dispose()
		$connection.Close()
		$connection.Dispose()
		
		# return the datatable
		return $datatable
		
	}
}

PROCESS {

	if ($columns.Value -ne $null) {$columns = @($columns.Value)}  else {$columns = $null}
	if ($orderby.Value -ne $null) {$orderby = @($orderby.Value)}  else {$orderby = $null}
	
	# Set locate's program directory
	if ($locatepath.length -eq 0) { $locatepath = "$env:LOCALAPPDATA\locate" }
	if ($install -eq $true){ Install-Locate -noprompt $false -locatepath $locatepath -advanced $advanced ; return }
	if ($du -eq $true) { 
		Get-DiskUsage -path $filename -locatepath $locatepath -limit $limit -orderby $orderby -descending $descending -topdirectories $topdirectories
		return 
	}
			
	# Check to see if the SQLite database exists, if it doesn't, prompt the user to install locate and populate the database.
	$locatedb = "$locatepath\locate.sqlite"

	if (!(Test-Path $locatedb)) {
		Write-Warning "locate database not found"
		$question = "Would you like to run the installer and populate the database now?"
		$choices = New-Object Collections.ObjectModel.Collection[Management.Automation.Host.ChoiceDescription]
		$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&Yes'))
		$choices.Add((New-Object Management.Automation.Host.ChoiceDescription -ArgumentList '&No'))
		
		$decision = $Host.UI.PromptForChoice($message, $question, $choices, 0)
		if ($decision -eq 1) { 
			Write-Host "Install skipped and no database to query. Quitting." -ForegroundColor Red -BackgroundColor Black
			break
		} else { Install-Locate -noprompt $true -locatepath $locatepath -advanced $advanced }
	}
	
	# If updatedb is called
	if ($updatedb -eq $true) { Update-LocateDB -locatepath $locatepath -homepath $homepath, -userprofile $userprofile -includemappedrives $includemappedrives -advanced $advanced; return }
	
	if ($sql.length -gt 0) {
		$dt = (Search-Filenames -locatepath $locatepath -sql $sql)
		return $dt
	} else {
		# If no arguments are passed, show message similar to gnu locate :)
		if ($filename.length -eq 0) { 
			Get-Help "$locatepath\Invoke-Locate.ps1"
			Write-Host "Report bugs to <clemaire@gmail.com>."
			return
		}
		
		# Perform a search, get datatable
		$dt = (Search-Filenames  $filename -locatepath $locatepath -s $s -columns $columns -where $where -orderby $orderby -descending $descending -limit $limit)
		
		# Show the simplified output for default searches
		if ($limit -eq 0) { $maxcolumn = 1 } else { $maxcolumn = 2 }
		if ($PSBoundParameters.count -eq $maxcolumn -or $dt.table.columns.count -eq 1) { return $dt.fullname } else { return $dt }
	}
}

END {
	# Clean up connections, if needed
	if ($command.connection -ne $null) { $command.Dispose() }
	if ($connection.state -ne $null) { $connection.Close(); $connection.Dispose() }
}
