<# 
 .SYNOPSIS
	Natively query CSV files using SQL syntax
	
 .DESCRIPTION
	This script will enable you to natively query a CSV file using SQL syntax using Microsoft's Text Driver.

	If you are running this script on a 64-bit system, and the 64-bit Text Driver is not installed, the script will automatically switch to a 32-bit shell 
	and execute the query. It will then communicate the data results to the 64-bit shell using Export-Clixml/Import-Clixml. 

	While the shell switch process is rather quick, you can avoid this step by running the script within a 32-bit 
	PowerShell shell ("$env:windir\syswow64\windowspowershell\v1.0\powershell.exe")

	The script returns datarows. See the examples for more details.
	
 .PARAMETER CSV
  The location of the CSV file to be queried.
	
 .PARAMETER FirstRowColumnNames
  This parameter specifies whether the first row contains column names. If the first row does not contain column names, the query engine automatically names the columns or "fields", F1, F2, F3 and so on.
  
 .PARAMETER Delimiter
  Optional. If you do not pass a Delimiter, then a comma will be used. Valid Delimiters include: tab "`t", pipe "|", semicolon ";", space " " and maybe a couple other things.
  
  When this parameter is used, a schema.ini must be created. In the event that one already exists, it will be moved to TEMP, then moved back once the script is finished executing.
  
 .PARAMETER SQL
  The SQL statement to be executed. To make command line queries easier, this script will convert the word "table" to the actual CSV formatted table name. 
  If the FirstRowColumnNames switch is not used, the query engine automatically names the columns or "fields", F1, F2, F3 and so on.
  
  Example: select F1, F2, F3, F4 from table where F1 > 5. See EXAMPLES for more example syntax.
 
 .PARAMETER shellswitch
  Internal parameter.
	
 .NOTES
    Author  : Chrissy LeMaire
    Requires: 	PowerShell 3.0
	Version: 0.7
	DateUpdated: 2015-May-19

 .LINK 
	https://gallery.technet.microsoft.com/scriptcenter/Query-CSV-with-SQL-c6c3c7e5
  	
 .EXAMPLE   
	.\Invoke-CsvSqlcmd.ps1 -csv C:\temp\housingmarket.csv -sql "select address from table where price < 250000" -FirstRowColumnNames
	
	This example return all rows with a price less than 250000 to the screen. The first row of the CSV file, C:\temp\housingmarket.csv, contains column names.
	
 .EXAMPLE 
	.\Invoke-CsvSqlcmd.ps1 -csv C:\temp\unstructured.csv -sql "select F1, F2, F3 from table" 
	
	This example will return the first three columns of all rows within the CSV file C:\temp\unstructured.csv to the screen. 
	Since the -FirstRowColumnNames switch was not used, the query engine automatically names the columns or "fields", F1, F2, F3 and so on.
 
 .EXAMPLE 
	$datatable = .\Invoke-CsvSqlcmd.ps1 -csv C:\temp\unstructured.csv -sql "select F1, F2, F3 from table" 
	$datatable.rows.count
 
	The script returns rows of a datatable, and in this case, we create a datatable by assigning the output of the script to a variable, instead of to the screen.
 
#> 
#Requires -Version 3.0

Param(
	[Parameter(Mandatory=$true)] 
	[ValidateScript({Test-Path $_ })]
	[string]$csv,
	[switch]$FirstRowColumnNames,
	[string]$Delimiter = ",",
	[Parameter(Mandatory=$true)] 
	[string]$sql,
	[switch]$shellswitch
	)
	
BEGIN {
	# If a non-default Delimiter is specified, a schema.ini file must be used.
	# If a schema.ini currently exists, it will be moved to TEMP temporarily.
	# Once the script has finished executing, it will be moved back.
	
	if ($Delimiter -ne ",") {
		$schemaexists = Test-Path schema.ini
		if ($schemaexists -eq $true) {
			Move-Item schema.ini $env:TEMP -Force
		}
	}
		
	# Check for Jet driver. 
	$jet = (New-Object System.Data.OleDb.OleDbEnumerator).GetElements() | Where-Object { $_.SOURCES_NAME -eq "Microsoft.Jet.OLEDB.4.0" }
	if ($jet -eq $null) { 
		Write-Warning "Switching to x86 shell, then switching back." 
		Write-Warning "This also requires a file to be written, so patience may be necesary." 
	}
}

PROCESS {

	# If the jet driver does not exist, the system is x64 and the Access Database Engine has not been installed.
	# Switch to x86 shell, which natively supports jet, then encode the SQL string, since some characters
	# can cause issues when being re-passed.
	
	if ($jet -eq $null) {
		$bytes  = [System.Text.Encoding]::UTF8.GetBytes($sql)
		$sql = [System.Convert]::ToBase64String($bytes)
		if ($firstRowColumnNames) { $frcn = "-FirstRowColumnNames" }
			&"$env:windir\syswow64\windowspowershell\v1.0\powershell.exe" "$PSCommandPath -csv '$csv' $frcn -Delimiter '$Delimiter' -SQL $sql -shellswitch" 
			return
	}
	
	# If the shell has switched, decode the $sql string.
	if ($shellswitch) {
		$bytes  = [System.Convert]::FromBase64String($sql)
		$sql = [System.Text.Encoding]::UTF8.GetString($bytes)
	}
	
	# Check for proper SQL syntax, which for the purposes of this script must include the word "table"
	if ($sql.ToLower() -notmatch "\btable\b") {
		throw "SQL statement must contain the word 'table'. Please see this script's documentation for more details."
	}

	# Unfortunately, passing delimiter within the connection string is unreliable, so we'll use schema.ini instead
	# if using any delimiter other than a comma.
	
	if ($Delimiter -ne ",") {
		$filename = Split-Path $csv -leaf
		Set-Content -Path schema.ini -Value "[$filename]"
		Add-Content -Path schema.ini -Value "Format=Delimited($Delimiter)"
	}
	
	# Setup the connection string. Data Source is the directory that contains the csv.
	# The file name is also the table name, but with a "#" intstead of a "."
	$csv = (Resolve-Path $csv).Path
	$datasource = Split-Path $csv
	$tablename = (Split-Path $csv -leaf).Replace(".","#")
	
	switch ($FirstRowColumnNames) {
		$true { $frcn = "Yes" }
		$false { $frcn = "No" }
	}
		
	$connstring = "Provider=Microsoft.Jet.OLEDB.4.0;Data Source=$datasource;Extended Properties='text;HDR=$frcn;';"

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
	} catch { throw "Could not open OLEDB connection." }
	
	# Execute the query, then load it into a datatable
	$dt = New-Object System.Data.DataTable
	try {
		$null = $dt.Load($cmd.ExecuteReader([System.Data.CommandBehavior]::CloseConnection))
	} catch { 
		$errormessage = $_.Exception.Message.ToString()
		if ($errormessage -like "*for one or more required parameters*") {
			Write-Error "Looks like your SQL syntax may be invalid. `nCheck the documentation for more information or start with a simple 'select top 10 * from table'"
		} else { Write-Error "Execute failed: $errormessage" }
	}
		
	# Use a file to facilitate the passing of a datatable from x86 to x64 if necessary
	if ($shellswitch) { 
		try { $dt | Export-Clixml "$env:TEMP\dt.xml" } catch { throw "Could not export datatable to file." }
	}
	
	# This should automatically close, but just in case...
	try {
		if ($conn.State -eq "Open") { $null = $conn.close }
		$null = $cmd.Dispose; $null = $conn.Dispose
	} catch { Write-Warning "Could not close connection. This is just an informational message." }
	
}

END {
	# Move original schema.ini back if it existed
	if ($schemaexists) { Move-Item "$env:TEMP\schema.ini" . -Force }
	if ($Delimiter -ne "," -and $schemaexists -eq $false) { Remove-Item schema.ini -Force }

	# If going between shell architectures, import a properly structured datatable.
	if ($dt -eq $null) { $dt = Import-Clixml "$env:TEMP\dt.xml"; Remove-Item  "$env:TEMP\dt.xml" }
	
	# Finally, return the resulting datatable
	if ($shellswitch -eq $false) { return $dt }
}