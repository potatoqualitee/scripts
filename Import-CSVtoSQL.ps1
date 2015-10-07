<# 
	 .SYNOPSIS
		
	 .DESCRIPTION
		
	 .PARAMETER CSV
	  The location of the CSV files to be imported. Multiple files are allowed, so long as they all support the same SQL query, and delimiter.
		
	 .PARAMETER FirstRowColumnNames
	  This parameter specifies whether the first row contains column names. If the first row does not contain column names, the query engine automatically names the columns or "fields", F1, F2, F3 and so on.
	  
	 .PARAMETER Delimiter
	  Optional. If you do not pass a Delimiter, then a comma will be used. Valid Delimiters include: tab "`t", pipe "|", semicolon ";", space " " and maybe a couple other things.
	  
	  When this parameter is used, a schema.ini must be created. In the event that one already exists, it will be moved to TEMP, then moved back once the module is finished executing.
	  
	 .PARAMETER SqlServer
	  
	 .PARAMETER Database
	  
	 .PARAMETER Table
	  
	 .PARAMETER 
		
	 .PARAMETER SQL
	  Optional. If you want to import just a specific query from your CSV file, use this parameter.
	  The SQL statement to be executed. To make command line queries easier, this module will convert the word "table" to the actual CSV formatted table name. 
	  If the FirstRowColumnNames switch is not used, the query engine automatically names the columns or "fields", F1, F2, F3 and so on.
	  
	  Example: select F1, F2, F3, F4 from table where F1 > 5. See EXAMPLES for more example syntax.
	 
	 .PARAMETER shellswitch
	  Internal parameter.
		
	 .NOTES
		Author  : Chrissy LeMaire
		Requires: 	PowerShell 3.0, SQL Client Stuff
		Version: 0.5
		DateUpdated: 2015-May-17

	 .LINK 
		https://gallery.technet.microsoft.com/scriptcenter/Query-CSV-with-SQL-c6c3c7e5
		
	 .EXAMPLE   
		.\Import-CSVtoSQL.ps1 -csv C:\temp\allCountries.txt -Delimiter `t -sqlserver sqlserver2014a -Database locations -Table allcountries -sql "select top 100000 * from table"
	 
	#> 
	#Requires -Version 3.0
	[CmdletBinding()] 
	Param(
		[Parameter(Mandatory=$true)] 
		[string[]]$csv,
		[switch]$FirstRowColumnNames,
		[string]$Delimiter = ",",
		[Parameter(Mandatory=$true)] 
		[string]$sqlserver,
		[Parameter(Mandatory=$true)] 
		[string]$Database,
		[Parameter(Mandatory=$true)] 
		[string]$Table,
		[string]$sql,
		[switch]$shellswitch
		)
		
	BEGIN {
		Write-Host "Script started..."
		$elapsed = [System.Diagnostics.Stopwatch]::StartNew() 
		try { 
			[void][Reflection.Assembly]::LoadWithPartialName("System.Data")
			[void][Reflection.Assembly]::LoadWithPartialName("System.Data.SqlClient")
		} catch { throw "SQL Client tools must be installed." }


		# In order to ensure consistent results, a schema.ini file must be created.
		# If a schema.ini currently exists, it will be moved to TEMP temporarily.
		
		if (!$shellswitch) {
			$resolvedcsv = @()
			foreach ($file in $csv) { $resolvedcsv += (Resolve-Path $file).Path }
			$csv = $resolvedcsv
			
			$movedschemaini = @{}
			foreach ($file in $csv) {
				$directory = Split-Path $file
				$schemaexists = Test-Path "$directory\schema.ini"
				if ($schemaexists -eq $true) {
					$newschemaname = "$env:TEMP\$(Split-Path $file -leaf)-schema.ini"
					$movedschemaini.Add($newschemaname,"$directory\schema.ini")
					Move-Item "$directory\schema.ini" $newschemaname -Force
				}
			}
		}
		
		# Check for drivers. 
		$provider = (New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | Where-Object { $_.SOURCES_NAME -like "Microsoft.ACE.OLEDB.*" }
		
		if ($provider -eq $null) {
			$provider = (New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | Where-Object { $_.SOURCES_NAME -like "Microsoft.Jet.OLEDB.*" }	
		}
		
		if ($provider -eq $null) { 
			Write-Warning "Switching to x86 shell, then switching back." 
			Write-Warning "This also requires a temporary file to be written, so patience may be necessary." 
		} else { 
			if ($provider -is [system.array]) { $provider = $provider[$provider.GetUpperBound(0)].SOURCES_NAME } else {  $provider = $provider.SOURCES_NAME }
		}
	}

	PROCESS {
		
		# Try hard to find a suitable provider; switch to x86 if necessary.
		# Encode the SQL string, since some characters
		if ($provider -eq $null) {
			$bytes  = [System.Text.Encoding]::UTF8.GetBytes($sql)
			$sql = [System.Convert]::ToBase64String($bytes)
			
			if ($firstRowColumnNames) { $frcn = "-FirstRowColumnNames" }
				$csv = $csv -join ","
				&"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" "Set-ExecutionPolicy Bypass -confirm:0; Invoke-CsvSqlcmd -csv $csv $frcn -Delimiter '$Delimiter' -SQL $sql -shellswitch" 
				return
		}
		# If the shell has switched, decode the $sql string.
		if ($shellswitch) {
			$bytes  = [System.Convert]::FromBase64String($sql)
			$sql = [System.Text.Encoding]::UTF8.GetString($bytes)
			$csv = $csv -Split ","
		}
		
		if ($sql.length -eq 0) { $sql = "select * from table" }
		
		# Check for proper SQL syntax, which for the purposes of this module must include the word "table"
		if ($sql.ToLower() -notmatch "\btable\b") {
			throw "SQL statement must contain the word 'table'. Please see this module's documentation for more details."
		}
		
		switch ($FirstRowColumnNames) {
				$true { $frcn = "Yes" }
				$false { $frcn = "No" }
		}
		
		# Does first line contain the specified delimiter?
		foreach ($file in $csv) {
			$firstline = Get-Content $file -First 1
			if (($firstline -match $Delimiter) -eq $false) {  throw "Delimiter $Delimiter not found in first row of $file." }
		}
		
		# If more than one csv specified, check to ensure number of columns match
		if ($csv -is [system.array]){ 
			$numberofcolumns = ((Get-Content $csv[0] -First 1) -Split $delimiter).Count 
			foreach ($file in $csv) {
				$firstline = Get-Content $file -First 1
				$newnumcolumns = ($firstline -Split $Delimiter).Count
				if ($newnumcolumns -ne $numberofcolumns) { throw "Multiple csv file mismatch. Do both use the same delimiter and have the same number of columns?" }
			}
		}
		
		# Create the resulting datatable
		$dt = New-Object System.Data.DataTable
		
		# Go through each file
		foreach ($file in $csv) {
		
			# Unfortunately, passing delimiter within the connection string is unreliable, so we'll use schema.ini instead
			# The default delimiter in Windows changes depending on country, so we'll do this for every delimiter, even commas.
			$filename = Split-Path $file -leaf; $directory = Split-Path $file
			Add-Content -Path "$directory\schema.ini" -Value "[$filename]"
			Add-Content -Path "$directory\schema.ini" -Value "Format=Delimited($Delimiter)"
			Add-Content -Path "$directory\schema.ini" -Value "ColNameHeader=$FirstRowColumnNames"
			
			# Setup the connection string. Data Source is the directory that contains the csv.
			# The file name is also the table name, but with a "#" instead of a "."
			$datasource = Split-Path $file
			$tablename = (Split-Path $file -leaf).Replace(".","#")
			
			$connstring = "Provider=$provider;Data Source=$datasource;Extended Properties='text;HDR=$frcn;';"
		
			# Do everything if no subset was selected via SQL
			if ($sql -eq $null) { $sql = "select * from table" }
			
			# To make command line queries easier, let the user just specify "table" instead of the
			# OleDbconnection formatted name (file.csv -> file#csv)
			$sql = $sql -replace "\btable\b"," [$tablename]"
			
			# Setup the OleDbconnection
			$conn = New-Object System.Data.OleDb.OleDbconnection
			$conn.ConnectionString = $connstring
			try { $conn.Open() } catch { throw "Could not open OLEDB connection." }
			
			# Setup the OleDBCommand
			try {
				$cmd = New-Object System.Data.OleDB.OleDBCommand
				$cmd.Connection = $conn
				$cmd.CommandText = $sql
			} catch { throw "Could not create OLEDB command." }
			
			Function OnSqlRowsCopied()
			{
				Write-Host "$($_.RowsCopied.ToString(`"N0`")) rows copied."
			}

			# Setup bulk copy
			Write-Host "Prepping bulk copy..."
			$connectionstring = "Data Source=$sqlserver;Integrated Security=true;Initial Catalog=$database" 
			$bulkcopy = New-Object Data.SqlClient.SqlBulkCopy($connectionstring, [System.Data.SqlClient.SqlBulkCopyOptions]::TableLock)
			$bulkcopy.DestinationTableName = $table
			$bulkcopy.bulkcopyTimeout = 0 
			$bulkCopy.BatchSize = 50000
			$bulkCopy.NotifyAfter = 50000
			$bulkCopy.Add_SqlRowscopied({Write-Host "$($args[1].RowsCopied) rows copied" })
	
			# Execute the query, then load it into a datatable
			try {
				$null = $bulkCopy.WriteToServer($cmd.ExecuteReader([System.Data.CommandBehavior]::CloseConnection))
			} catch { 
				$errormessage = $_.Exception.Message.ToString()
				if ($errormessage -like "*for one or more required parameters*") {
					throw "Looks like your SQL syntax may be invalid. `nCheck the documentation for more information or start with a simple -sql 'select top 10 * from table'"
				} else { Write-Error "Aww, malformed data. $errormessage" }
			}
			
			 $objectEvent = Register-ObjectEvent $bulkcopy SqlRowsCopied -Action { Write-Host "Copied $($eventArgs.RowsCopied) rows"}
			 
			# This should automatically close, but just in case...
			try { $null = $conn.close; $null = $cmd.Dispose(); $null = $conn.Dispose(); $null = $bulkCopy.close(); $null = $bulkcopy.dispose() } catch {}
		}
	}

	END {
		try { $null = $conn.close; $null = $cmd.Dispose(); $null = $conn.Dispose(); $null = $bulkCopy.close(); $null = $bulkcopy.dispose() } catch {} # don't care if it fails
		
		# Delete new schema files
		foreach ($file in $csv) {
			$directory = Split-Path $file
			$null = Remove-Item "$directory\schema.ini" -Force -ErrorAction SilentlyContinue
		}
		
		# Move original schema.ini's back if they existed
		if ($movedschemaini.count -gt 0) {
			foreach ($item in $movedschemaini) {
				Move-Item $item.keys $item.values -Force -ErrorAction SilentlyContinue	
			}
		}
		
		if ($shellswitch -eq $false) { 
		$totaltime = [math]::Round($elapsed.Elapsed.TotalSeconds,2)
		Write-Host "Total Elapsed Time: $totaltime seconds." -ForegroundColor Green
		}	
	}