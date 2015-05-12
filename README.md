# Scripts and files
Various scripts like Invoke-CsvSqlcmd.ps1

Invoke-CsvSqlcmd.ps1
--------------
Invoke-CsvSqlcmd.ps1 will enable you to natively query a CSV file using SQL syntax using Microsoft's Text Driver. The syntax is as simple as:

    .\Invoke-CsvSqlcmd.ps1 -csv file.csv -sql "select * from table"
	
To make command line queries easier, this script will convert the word "table" within the -sql parameter to the actual CSV formatted table name.   If the FirstRowColumnNames switch is not used, the query engine automatically names the columns or "fields", F1, F2, F3, etc.

If you are running Invoke-CsvSqlcmd.ps1 on a 64-bit system, and the 64-bit Text Driver is not installed, the script will automatically switch to a 32-bit shell and execute the query. It will then communicate the data results to the 64-bit shell using Export-Clixml/Import-Clixml. 

While the shell switch process is rather quick, you can avoid this step by running the script within a 32-bit  PowerShell shell ("$env:windir\syswow64\windowspowershell\v1.0\powershell.exe")

Other examples
-----
    .\Invoke-CsvSqlcmd.ps1 -csv C:\temp\housingmarket.csv -sql "select address from table where price < 250000" -FirstRowColumnNames

This example return all rows with a price less than 250000 to the screen. The first row of the CSV file, C:\temp\housingmarket.csv, contains column names.

    .\Invoke-CsvSqlcmd.ps1 -csv C:\temp\unstructured.csv -sql "select F1, F2, F3 from table where F3 > 7" 

This example will return the first three columns of all rows within the CSV file C:\temp\unstructured.csv to the screen. 
Since the -FirstRowColumnNames switch was not used, the query engine automatically names the columns or "fields", F1, F2, F3 and so on.

    $datatable = .\Invoke-CsvSqlcmd.ps1 -csv C:\temp\unstructured.csv -sql "select F1, F2, F3 from table"  
    $datatable.rows.count

Invoke-CsvSqlcmd.ps1 returns rows of a datatable, and in this case, we create a datatable by assigning the output of the script to a variable, instead of to the screen.

Import-vCentertoRDG.ps1
--------------
Imports vCenter Server folders and servers to an Remote Desktop Connection Manager 2.7 XML file named vSphere.rdg in the current directory. The display name of each server is the vCenter server name, and servername is the IP address, unless -DNSPreferred is specified. Note that it uses Get-Folder, which connects to all currently connected vCenter servers unless -Server is specified.

Only Windows servers are added, and any empty folder (to include folders that only contain non-Windows VMs) will be skipped. Because RDCM does not really support nested groups, subgroups are named $folder-$subfolder. 
	
Remote Desktop Connection Manager 2.7 can be downloaded here: http://www.microsoft.com/en-us/download/details.aspx?id=44989

    .\Import-vCentertoRDG.ps1
Exports all folders and servers within the currently connected vCenter server. Server names appear as they do in vCenter, and use IP addresses to connect.

    .\Import-vCentertoRDG.ps1 -Server vcenter.ad.local -Template H:\AD.rdg -DNSPreferred -Folder Infrastructure
	
Exports all folders and servers within the Infrastructure folder in the vcenter.ad.local vCenter server, and builds on top of AD.rdg, but saves to vSphere.rdg. Uses DNS names instead of IPs.
If a DNS name is not available, it uses the IP address instead. If neither are available, it uses the vCenter server name.
 