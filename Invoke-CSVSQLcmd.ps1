<# 
 .SYNOPSIS
	Natively query CSV files using SQL syntax
	
 .DESCRIPTION
	This script will enable you to natively query a CSV file using SQL syntax using Microsoft's Text Driver.

	The script returns datarows. See the examples for more details.
	
 .PARAMETER CSV
  The location of the CSV files to be queried. Multiple files are allowed, so long as they all support the same SQL query, and delimiter.
	
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
	Version: 0.9.3
	DateUpdated: 2015-May-16

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
[CmdletBinding()] 
Param(
	[Parameter(Mandatory=$true)] 
	[ValidateScript({Test-Path $_ })]
	[string[]]$csv,
	[switch]$FirstRowColumnNames,
	[string]$Delimiter = ",",
	[Parameter(Mandatory=$true)] 
	[string]$sql,
	[switch]$shellswitch
	)
	
BEGIN {
	# In order to ensure consistent results, a schema.ini file must be created.
	# If a schema.ini currently exists, it will be moved to TEMP temporarily.
	$movedschemaini = @{}
	foreach ($file in $csv) {
		$file = (Resolve-Path $file).Path; $directory = Split-Path $file
		$schemaexists = Test-Path "$directory\schema.ini"
		if ($schemaexists -eq $true) {
			$newschemaname = "$env:TEMP\$(Split-Path $file -leaf)-schema.ini"
			$movedschemaini.Add($newschemaname,"$directory\schema.ini")
			Move-Item "$directory\schema.ini" $newschemaname -Force
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

	# Create the resulting datatable
	$dt = New-Object System.Data.DataTable
	
	# Try hard to find a suitable provider; switch to x86 if necessary.
	# Encode the SQL string, since some characters
	
	if ($provider -eq $null) {
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
	
	
	
	switch ($FirstRowColumnNames) {
		$true { $frcn = "Yes" }
		$false { $frcn = "No" }
	}
		

		
	}
}

END {
	# If going between shell architectures, import a properly structured datatable.
	
	# Finally, return the resulting datatable
	if ($shellswitch -eq $false) { return $dt }
}