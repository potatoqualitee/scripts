<# 
 .SYNOPSIS
	Imports vCenter Server folders and servers to a Remote Desktop Connection Manager 2.7 XML file
	
 .DESCRIPTION
	Imports vCenter Server folders and servers to an Remote Desktop Connection Manager 2.7 XML file named vSphere.rdg in the current directory. The display name of each server is the vCenter server name, and servername is the IP address, unless -DNSPreferred is specified. Note that it uses Get-Folder, which connects to all currently connected vCenter servers unless -Server is specified.

	Only Windows servers are added, and any empty folder (to include folders that only contain non-Windows VMs) will be skipped. Because RDCM does not really support nested groups, subgroups are named $folder-$subfolder. 
	
	Remote Desktop Connection Manager 2.7 can be downloaded here: http://www.microsoft.com/en-us/download/details.aspx?id=44989
	
 .PARAMETER Server
	Connects to a specified vCenter server. If no server is specified, the currently connected vCenter server(s) will be used.
	
 .PARAMETER Template
	Not required. This builds on an existing RDG 2.7 template. If a vCenter folder name currently exists in the RDG file as a group, all servers within it will be updated. Note that this will not overwrite your template file; the output will still be written to vSphere.rdg.
 
 .PARAMETER Folder
	Only exports a specific folder and its subfolders.
 
 .PARAMETER DNSPreferred
	Uses DNS hostnames instead of IP addresses to connect to servers. In the event that a DNS entry cannot be found, the IP address is used as a fallback. In the event that an IP cannot be found, the servername is used.

 .NOTES
    Author  : Chrissy LeMaire
    Requires: 	VMware.VimAutomation.Core Snapin (PowerCLI), read access to vCenter
	Version: 0.5
	DateUpdated: 2015-Apr-13

 .LINK 
	https://gallery.technet.microsoft.com/scriptcenter/Import-vCenter-to-RDG-4fb173f0
  	
 .EXAMPLE   
	.\Import-vCentertoRDG.ps1
	Exports all folders and servers within the currently connected vCenter server. Server names appear as they do in vCenter, and use IP addresses to connect.
	
 .EXAMPLE 
	.\Import-vCentertoRDG.ps1 -Server vcenter.ad.local -Template H:\AD.rdg -DNSPreferred -Folder Infrastructure
	Exports all folders and servers within the Infrastructure folder in the vcenter.ad.local vCenter server, and builds on top of AD.rdg, but saves to vSphere.rdg. Uses DNS names instead of IPs.
	If a DNS name is not available, it uses the IP address instead. If neither are available, it uses the vCenter server name.
 
#> 
#Requires -Version 3.0
[CmdletBinding(DefaultParameterSetName="Default")]

Param(
	[parameter(Position=0)]
	[string]$Server,
	[string]$Template,
	[string]$Folder,
	[switch]$DNSPreferred
	)
	
BEGIN {

	Function Get-GroupsXML {
		param(
			[string]$server,
			[string]$folder,
			[bool]$DNSPreferred
		)

		if ($server) { $folders = Get-Folder -Type VM -Server $server } else { $folders = Get-Folder -Type VM }
		if ($folder) { $folders = $folders | Where-Object { $_.Name -eq $folder -or $_.Parent.Name -eq $folder } }		
		$folders = $folders | Sort-Object | Get-Unique

		$groupsxml = @()

		foreach ($subfolder in $folders) {
			$groupname = Get-RDCGroupName $subfolder
			$folderid = $subfolder.id
			$foldername = $subfolder.name

			Write-Host "Discovering $groupname" -ForegroundColor Yellow
			$escapedgroupname = [System.Security.SecurityElement]::Escape($groupname)
			[xml]$groupxml = "<group><properties><name>$escapedgroupname</name><expanded>True</expanded></properties></group>"
			
			$serversxml = Get-ServersXML $foldername $DNSPreferred
			
			if ($serversxml -eq $null) { Write-Host "No Windows servers found. Moving on." ; continue }
		
			foreach ($server in $serversxml) {
				$serverxml = [xml]$server
				$serverchild = $groupxml.ImportNode($serverxml.server, $true)
				$null = $groupxml.DocumentElement.AppendChild($serverchild)
			}
			$groupsxml += $groupxml
		}
		return $groupsxml
	}
	
	Function Get-RDCGroupName ($vmfolder) {
		if ($vmfolder.name -eq "vm" ) { return "vSphere Root" }
		$groupname = $vmfolder.name
		
		if ($vmfolder.parent.name -ne "vm") {
			do {
				$parentname = $vmfolder.parent.name
				$groupname = "$parentname-$groupname"
				$vmfolder = $vmfolder.parent
			} while ($vmfolder.parent.name -ne "vm")
		}
		return $groupname
	}
	 
	Function Get-ServersXML {
			param(
				[string]$foldername,
				[bool]$DNSPreferred
			)
		$serversxml = @()
		$vms = Get-VM | Where-Object { $_.Folder.name -eq $foldername }	
		foreach ($vm in $vms) {
			$vmguest = $vm | Get-VMguest
			$vmview = $vm |Get-View
			$servername = $vmguest.IPAddress | Select -First 1
			$dns = $vm.ExtensionData.Guest.HostName
			$description = ([System.Security.SecurityElement]::Escape($vm.Notes)).Trim()
			$vmname = $vm.Name
			$escapedvmname = [System.Security.SecurityElement]::Escape($vm.Name)
			$os = $vmguest.OSFullName
			$guest = $vmview.Summary.Config.GuestFullName

			if ($dnspreferred -and $dns -ne $null) { $servername = $dns }
			if ($servername -eq $null) { $servername = $escapedvmname }
			if ($os -like '*Windows*' -or $guest -like "*windows*") {
				$serversxml += "<server><properties><name>$servername</name><displayName>$escapedvmname</displayName><comment>$description</comment></properties></server>"
				Write-Host "Found $vmname"
			}
		}
		return $serversxml
	}
		
	Function Get-RDGXML {
			param(
				[xml]$xmlroot,
				[array]$groupsxml
			)
		
		# Add back displayname element for any server that is missing it.		
		foreach ($rootserver in $xmlroot.DocumentElement.file.group.server) {
			if ($rootserver.properties.displayName -eq $null) {
				[xml]$displayxml = "<displayName>$($rootserver.properties.name)</displayName>"
				$newitem = $rootserver.OwnerDocument.ImportNode($displayxml.SelectSingleNode("displayName"), $true)
				$updatedserver = $rootserver.properties.InsertAfter($newitem,$_.firstchild)
			}
		}
		
		$rootgroups = $xmlroot.DocumentElement.file
		$addgroups = $groupsxml | Where-Object { $rootgroups.group.properties.name -notcontains $_.group.properties.name }
		$updategroups = $groupsxml | Where-Object { $rootgroups.group.properties.name -contains $_.group.properties.name }
		
		foreach ($group in $addgroups) {
			if ($group.group.server.count -gt 0) {
				Write-Verbose "Adding $($group.group.properties.name)"
				$newitem = $xmlroot.ImportNode($group.group, $true)
				$null = $xmlroot.DocumentElement.file.AppendChild($newitem) 
			}
		}
 
		foreach ($updategroup in $updategroups) {
			$groupname = $updategroup.group.properties.name
			Write-Verbose "Updating $groupname"
			$rootgroup = $rootgroups.group | Where-Object { $_.properties.name -eq $updategroup.group.properties.name }
			$removeservers = $rootgroup.server | Where-Object { $updategroup.group.server.properties.displayname -contains $_.properties.displayname }
				
			foreach ($removal in $removeservers) { Write-Verbose "Updating $($removal.properties.name) in $groupname"; $null = $rootgroup.RemoveChild($removal) }
			foreach ($groupserver in $updategroup.group.server) {
				$newitem = $xmlroot.ImportNode($groupserver, $true)
				$updatedgroup = $rootgroup.AppendChild($newitem)
			}
		}
		return  $xmlroot
	}
}

PROCESS {
	
	# If a template is specified, get it, and check version 
	# If no template is specified, build a basic one.
	if ($template) {
		if ((Test-Path $template) -eq $false) { throw "Cannot find template file" }
		[xml]$xmlroot = Get-Content -Path $template
		if ($xmlroot.RDCMan.programVersion -ne "2.7") { throw "RDCMan programVersion must be 2.7" }
	} else {
		[xml]$xmlroot = '<?xml version="1.0" encoding="utf-8"?><RDCMan programVersion="2.7" schemaVersion="3">
		<file><credentialsProfiles /><properties><name>vSphere</name><expanded>True</expanded></properties>
		</file><connected /><favorites /><recentlyUsed /></RDCMan>'
	}
	
	# This allows you to run it in a regular PowerShell shell. Or you can use PowerCLI.
	if(-not (Get-PSSnapin VMware.VimAutomation.Core -ErrorAction SilentlyContinue)) { Add-PSSnapin VMware.VimAutomation.Core }
	if ($Server.length -ne 0) { Write-Host "Connecting to vCenter Server..."; $null = Connect-VIServer $Server -WarningAction SilentlyContinue }
	if ($defaultVIServer -eq $null -and $Server.length -eq 0) { throw "You are not currently connected to a VIServer and did not specify -Server. Please connect or specify a server." }
	
	# Get groups & servers
	$groupsxml = Get-GroupsXML -server $server -folder $folder -dnspreferred $dnspreferred
	
	# Build RDG file
	$xmlroot = Get-RDGXML -xmlroot $xmlroot -groupsxml $groupsxml
	
	#Save
	$xmlroot.Save("$pwd\vSphere.rdg")
}

END {
	Write-Host "$pwd\vSphere.rdg has been created. Script complete." -ForegroundColor Green
}