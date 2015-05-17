# Scripts and files
Various scripts like Invoke-CsvSqlcmd.ps1

Invoke-Locate.ps1 
--------------
Invoke-Locate.ps1 Port for GNU Locate within Windows. Uses SQLite.

This script was made in the spirit of (Linux/Unix) GNU findutils' locate. 

While the name of this script is Invoke-Locate, it actually creates two persistent aliases: locate and updatedb. A fresh index is automatically created every 6 hours, and updatedb can be used force a refresh. Indexing takes anywhere from 30 seconds to 15 minutes, depending on the speed of your drives. Performing the actual locate takes about 300 milliseconds. This is made possible by using SQLite as the backend. Invoke-Locate supports both case-sensitive, and case-insensitive searches, and is case-insensitive by default. 

Locate searches are per-user, and the database is stored securely in your home directory. You can search system files and your own home directory, but will not be able to search for filenames in other users' directories. 

Note: This is a work in progress (version 0.x), and I'm currently testing it in various environments. Please let me know if you have any issues. I fixed a few bugs over the weekend. Please download the newest version.

Basic functionality
---
	# install locate with basic functionality
    .\Invoke-Locate.ps1 -install
	
	# Perform a case-insensitive search for *csv*.ps1
	locate csv*.ps1

	# Perform case-sensitive search which return the path to any file or directory with System.Data in the name.
	locate -s System.Data
	
	# Force database refresh. This generally takes just a few minutes.
	updatedb
	
	# Displays filenames in alphabetical order.
	locate .iso  -orderby fullname
	
	# Similar to SQL's "LIKE" syntax, underscores are used to specify "any single character." locate powers?ell.exe also works.
	locate powers_ell.exe
	
Advanced functionality
---

	# install locate with advanced functionality -- takes longer to populate database, but provides additional features
    .\Invoke-Locate.ps1 -install -advanced
	
	# Execute SQL statement and return a datatable. When the -sql switch is used, all other switches are ignored.
	locate -sql "select directory from files where fullname like'%chrissy%resume%.docx' and lastmodified > '1/1/2015' order by lastmodified"
	
	# Return total number and size of the files within your home directory
	locate -sql "select count(*), sum(gb) from files where directory like '%$env:HOMEPATH%'"

	# Find the 10 largest indexed files, ordered by size
	locate -sql "SELECT fullname, gb FROM files ORDER BY size DESC LIMIT 10"

	# This command displays disk usage output for the current directory, including files and directories. 
	du
	
	# Aggregates disk usage information by directory for C:\inetpub. Don't run this in C: or C:\Program Files unless you have a lot of patience and RAM
	du -topdirectories C:\inetpub 

	
Import-vCentertoRDG.ps1
--------------
Imports vCenter Server folders and servers to an Remote Desktop Connection Manager 2.7 XML file named vSphere.rdg in the current directory. The display name of each server is the vCenter server name, and servername is the IP address, unless -DNSPreferred is specified. Note that it uses Get-Folder, which connects to all currently connected vCenter servers unless -Server is specified.

Only Windows servers are added, and any empty folder (to include folders that only contain non-Windows VMs) will be skipped. Because RDCM does not really support nested groups, subgroups are named $folder-$subfolder. 
	
Remote Desktop Connection Manager 2.7 can be downloaded here: http://www.microsoft.com/en-us/download/details.aspx?id=44989

    .\Import-vCentertoRDG.ps1
Exports all folders and servers within the currently connected vCenter server. Server names appear as they do in vCenter, and use IP addresses to connect.

    .\Import-vCentertoRDG.ps1 -Server vcenter.ad.local -Template H:\AD.rdg -DNSPreferred -Folder Infrastructure
	
Exports all folders and servers within the Infrastructure folder in the vcenter.ad.local vCenter server, and builds on top of AD.rdg, but saves to vSphere.rdg. Uses DNS names instead of IPs.
If a DNS name is not available, it uses the IP address instead. If neither are available, it uses the vCenter server name


Import-CsvToSQL.ps1
--------------
Documentation to come later.
 