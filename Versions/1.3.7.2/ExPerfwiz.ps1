#################################################################################
# 
# The sample scripts are not supported under any Microsoft standard support 
# program or service. The sample scripts are provided AS IS without warranty 
# of any kind. Microsoft further disclaims all implied warranties including, without 
# limitation, any implied warranties of merchantability or of fitness for a particular 
# purpose. The entire risk arising out of the use or performance of the sample scripts 
# and documentation remains with you. In no event shall Microsoft, its authors, or 
# anyone else involved in the creation, production, or delivery of the scripts be liable 
# for any damages whatsoever (including, without limitation, damages for loss of business 
# profits, business interruption, loss of business information, or other pecuniary loss) 
# arising out of the use of or inability to use the sample scripts or documentation, 
# even if Microsoft has been advised of the possibility of such damages
#
#################################################################################
#
# Script to help automate the collection of performance data on Exchange 2007/2010 servers 
# Created by mikelag@microsoft.com 
# Last Update 3.12.2012
# Version 1.3.7
#
# 1.3.1 - Resolved encoding problem in Exchange 2010 perfmon counter sets
# 1.3.2 - Fixed Turkish character encoding out-file issue
# 1.3.3 - Fixed -full switch not working on Exchange 2010 servers
# 1.3.4 - Added Client: RPCs Failed counter for IS and RPC Client Access
#		- Added .NET CLR Memory Gen Collections, .NET Promoted Memory Counters, and .NET Pinned objects
#		- Removed \MSExchangeWS\Average Response Time from 2007 CAS counters since it does not exist.
# 1.3.5 - Added script variables for Operating System
#		- Rewrote CreateCounter Function to remove duplicate data
#		- Fixed StartCollection function
#		- Removed -cnf option for Windows 2003 based servers since the logs would run continuously. Log roll is disabled. Maxsize or duration is used, whichever one comes first
# 1.3.6 (Beta) 
#		- Added -server switch to allow remote servers to be specified. If server switch is not specified, then the local server is used
#		- Added function that tests remote registry access and whether or not the launching user has required permissions to access the remote servers registry
#		- Added additional ActiveSync Counters to help track queuing and latencies
#		- Added -begin and -end times for scheduling purposes
#		- Added additional error handling
#		- Updated function on how we obtain CMS name information
#		- Added check for Windows 2008 R2 servers to ensure that EMS is being launched as Administrator
#		- Added Exmon support (-exmon and -exmonduration)
#		- Added Database Counters to HUB Transport role
# 1.3.6 (Release)
#		- Updated 2010 Transport counters to include DeliveryAgents.
#		- Updated UM, MSExchangeAB, and RPC/HTTP counters
#
# 1.3.7 (a-larryh,mikelag,amyma)
#		- Added quiet mode for full automation, assumes if an existing Exchange_PerfWiz data collector exists
#		- then we will internally force the StopAndDeleteCounter, then resume current execution
#		- Fixed error where -delete was throwing an error since the OS was not detected properly
#		- Added HTTP Service Request Queues counters for 2007/2010 servers
#		- Updated 2010 Mailbox counters to include Minimsg Msg table seeks/sec
# 1.3.7.1
#		- Added -CustomCounterPath switch
#		- Fixed ESE extended counter bug where it would create it as the wrong Value Type.
#		- Added \MSExchange Active Manager(*)\Database Mounted
#		- Added Database Threads Blocked for Exchange 2007
#		- Added Search indexer counters for amount of paused/disabled databases
#		- Added "\MSExchangeIS Mailbox(*)\*" to Exchange 2010 Full counter set since it was missing
#		- Added \MSExchangeIS Mailbox(*)\RPC Average Latency (Client) for Exchange 2010 SP1 servers
# 1.3.7.2
#		- Added "\Processor Information(*)\*" counters


Param (
[int]$interval,
[string]$duration,
[int]$maxsize = 512,
[switch]$stop,
[switch]$threads,
[switch]$query,
[switch]$full,
[switch]$start,
[switch]$delete,
[switch]$circular,
[switch]$StoreExtendedOn, 
[switch]$StoreExtendedOff, 
[switch]$EseExtendedOn, 
[switch]$EseExtendedOff, 
[switch]$WebHelp, 
[string]$filepath, 
[string]$begin, 
[string]$end, 
[string]$Server,
[switch]$debug,
[switch]$Exmon,
[string]$ExmonDuration,
[switch]$quiet,
[string]$CustomCounterPath
)
$script:Windows2003 = $false
$script:Windows2008 = $false
$script:Windows2008R2 = $false
$oldDebugPreference = $DebugPreference  

function GetExServerInfo
{
	if(!$server)
	{
		$Server = ${env:computername}
	}
	else 
	{
		$Server = $Server
	}
	$Error.Clear()
	$TestServerName = Get-ExchangeServer -Identity $Server -ErrorAction SilentlyContinue
	if ($Error)
	{
		# Get CMS Name
		$Server = (Get-MailboxServer | Where-Object {$_.RedundantMachines -eq $server}).name
		
		if ($Server -eq $null)
		{
			Write-Host "================================================================"
			Write-Host ""
			Write-Host "Server name not found or server specified does not have Exchange installed. Exiting script." -ForegroundColor Yellow
			Write-Host ""
			Exit
		}
		else
		{
			$Script:ServerName = $Server
		}
	}
	else
	{
		$Script:ServerName = $Server
	}
	Write-Debug "Servername: $Servername"
	$ExVersion = (get-exchangeserver -Identity $ServerName).AdminDisplayVersion.Major 
	if ($ExVersion -eq 8){[bool]$script:Exchange2007 = $true}
	elseif ($ExVersion -eq 14){[bool]$script:Exchange2010 = $true}
	Write-Debug "Exchange Version: $ExVersion"
}

function GetOSVersion
{
	#Added Remoting
	$script:OSVerMajor = ((Get-WmiObject Win32_OperatingSystem -ComputerName $ServerName).Version).Split(".")[0]
	$script:OSVerMinor = ((Get-WmiObject Win32_OperatingSystem -ComputerName $ServerName).Version).Split(".")[1]
	If (($OSVerMajor -eq 5) -and ($OSVerMinor -eq 2)){$script:Windows2003 = $true}
	If (($OSVerMajor -eq 6) -and ($OSVerMinor -eq 0)){$script:Windows2008 = $true}
	If (($OSVerMajor -eq 6) -and ($OSVerMinor -eq 1)){$script:Windows2008R2 = $true}
	Write-Debug "OS Version: $OSVerMajor.$OSVerMinor"
}

function IsAdmin 
{  
	$identity = [System.Security.Principal.WindowsIdentity]::GetCurrent()  
	$principal = new-object System.Security.Principal.WindowsPrincipal($identity)  
	$admin = [System.Security.Principal.WindowsBuiltInRole]::Administrator  
	$IsAdmin = $principal.IsInRole($admin)  
	Write-Debug "IsAdmin: $Admin"
	if ($Windows2008R2  -and !$IsAdmin)
	{
		Write-Host ""
		Write-warning "Script requires elevated access to run. Open the Exchange Management Shell using Run as Administrator"
		Write-Host ""
		exit
	}
} 

function CreateCounterList{
	
	$script:roles = @()
	Write-Host ""
	Write-Host "Exchange Server:" $ServerName
	Write-Host ""
	$GetServer = Get-ExchangeServer -Identity $ServerName
	
	if ($full -eq $true -and $Exchange2007){
	# Full Counter set for Mbx-Cas-Hub
	$Counters = @(
"\.NET CLR Exceptions(*)\*"
"\.NET CLR Memory(*)\*"
"\ASP.NET Apps v2.0.50727(*)\*"
"\ASP.NET v2.0.50727\*"
"\Cache\*"
"\LogicalDisk(*)\*"
"\Memory\*"
"\MSExchange ActiveSync\*"
"\MSExchange AD RMS Prelicensing Agent\*"
"\MSExchange ADAccess Caches(*)\*"
"\MSExchange ADAccess Domain Controllers(*)\*"
"\MSExchange ADAccess Global Counters\*"
"\MSExchange ADAccess Local Site Domain Controllers(*)\*"
"\MSExchange ADAccess Processes(*)\*"
"\MSExchange Availability Service\*"
"\MSExchange Calendar Attendant\*"
"\MSExchange Connection Filtering Agent\*"
"\MSExchange Content Filter Agent\*"
"\MSExchange Database ==> Instances(*)\*"
"\MSExchange Database ==> TableClasses(*)\*"
"\MSExchange Database(*)\*"
"\MSExchange Extensibility Agents(*)\*"
"\MSExchange Journaling Agent\*"
"\MSExchange Managed Folder Assistant\*"
"\MSExchange Oledb Events(*)\*"
"\MSExchange Oledb Resource(*)\*"
"\MSExchange OWA\*"
"\MSExchange Protocol Analysis Agent\*"
"\MSExchange Protocol Analysis Background Agent\*"
"\MSExchange Recipient Cache(*)\*"
"\MSExchange Recipient Filter Agent\*"
"\MSExchange Search Indexer\*"
"\MSExchange Search Indices(*)\*"
"\MSExchange Secure Mail Transport(*)\*"
"\MSExchange Sender Filter Agent\*"
"\MSExchange Sender Id Agent\*"
"\MSExchange Store Driver(*)\*"
"\MSExchange Store Interface(*)\*"
"\MSExchange Topology(*)\*"
"\MSExchange Transport Rules(*)\*"
"\MSExchange Update Agent\*"
"\MSExchange Web Mail(*)\*"
"\MSExchangeAL(*)\*"
"\MSExchangeAutodiscover\*"
"\MSExchangeEdgeSync Topology\*"
"\MSExchangeImap4\*"
"\MSExchangeIS Client(*)\*"
"\MSExchangeIS Mailbox(*)\*"
"\MSExchangeIS Public(*)\*"
"\MSExchangeIS\*"
"\MSExchangeMailSubmission(*)\*"
"\MSExchangeSA - NSPI Proxy\*"
"\MSExchangeTransport Batch Point(*)\*"
"\MSExchangeTransport Database(*)\*"
"\MSExchangeTransport DSN(*)\*"
"\MSExchangeTransport Dumpster\*"
"\MSExchangeTransport Pickup(*)\*"
"\MSExchangeTransport Queues(*)\*"
"\MSExchangeTransport Resolver(*)\*"
"\MSExchangeTransport Routing(*)\*"
"\MSExchangeTransport SmtpReceive(*)\*"
"\MSExchangeTransport SmtpSend(*)\*"
"\MSExchangeUMClientAccess(*)\*"
"\MSExchangeWS\*"
"\MSFTESQL-Exchange:Catalogs(*)\*"
"\MSFTESQL-Exchange:FD(*)\*"
"\MSFTESQL-Exchange:Indexer PlugIn(*)\*"
"\MSFTESQL-Exchange:Service\*"
"\Netlogon(*)\*"
"\Network Interface(*)\*"
"\Objects\*"
"\Paging File(*)\*"
"\PhysicalDisk(*)\*"
"\Process(*)\*"
"\Processor(*)\*"
"\Processor Information(*)\*"
"\Redirector\*"
"\RPC/HTTP Proxy Per Server\*"
"\RPC/HTTP Proxy\*"
"\Server Work Queues(*)\*"
"\Server\*"
"\System\Context Switches/sec"
"\System\Processor Queue Length"
"\TCPv4\*"
"\TCPv6\*"
"\Web Service(*)\*"
)
$script:roles += [string]"Full"
Write-Debug "Added Exchange 2007 Full Counters"
}
elseif ($full -eq $true -and $Exchange2010){
	$Counters = @(
"\.NET CLR Exceptions(*)\*"
"\.NET CLR Memory(*)\*"
"\ASP.NET Apps v2.0.50727(*)\*"
"\ASP.NET v2.0.50727\*"
"\Cache\*"
"\LogicalDisk(*)\*"
"\Memory\*"
"\MSFTESQL-Exchange:Catalogs(*)\*"
"\MSFTESQL-Exchange:FD(*)\*"
"\MSFTESQL-Exchange:Indexer PlugIn(*)\*"
"\MSFTESQL-Exchange:Service\*"
"\Network Interface(*)\*"
"\Objects\*"
"\Paging File(*)\*"
"\PhysicalDisk(*)\*"
"\Process(*)\*"
"\Processor(*)\*"
"\Redirector\*"
"\Server\*"
"\Server Work Queues(*)\*"
"\System\Context Switches/sec"
"\System\Processor Queue Length"
"\Web Service(*)\*"
"\RPC/HTTP Proxy\*"
"\RPC/HTTP Proxy Per Server\*"
"\TCPv4\*"
"\TCPv6\*"
"\Netlogon(*)\*"
"\MSExchange Active Manager(*)\*"
"\MSExchange Active Manager Client(*)\*"
"\MSExchange Active Manager Server\*"
"\MSExchange ActiveSync\*"
"\MSExchange ADAccess Caches(*)\*"
"\MSExchange ADAccess Domain Controllers(*)\*"
"\MSExchange ADAccess Global Counters\*"
"\MSExchange ADAccess Local Site Domain Controllers(*)\*"
"\MSExchange ADAccess Processes(*)\*"
"\MSExchange Approval Assistant\*"
"\MSExchange Approval Framework(_total)\*"
"\MSExchange Assistants - Per Assistant(*)\*"
"\MSExchange Assistants - Per Database(*)\*"
"\MSExchange Availability Service\*"
"\MSExchange Calendar Attendant\*"
"\MSExchange Calendar Notifications Assistant\*"
"\MSExchange Calendar Repair Assistant\*"
"\MSExchange Connection Filtering Agent\*"
"\MSExchange Content Filter Agent\*"
"\MSExchange Control Panel\*"
"\MSExchange Conversations Transport Agent\*"
"\MSExchange Database(*)\*"
"\MSExchange Database ==> Instances(*)\*"
"\MSExchange Database ==> TableClasses(*)\*"
"\MSExchange Decryption Agent\*"
"\MSExchange Encryption Agent\*"
"\MSExchange Extensibility Agents(*)\*"
"\MSExchange FreeBusy Assistant\*"
"\MSExchange Inbound SMS Delivery Agent\*"
"\MSExchange Journal Report Decryption Agent\*"
"\MSExchange Journaling Agent\*"
"\MSExchange Junk E-mail Options Assistant\*"
"\MSExchange Log Search Service\*"
"\MSExchange Mail Submission(*)\*"
"\MSExchange Mailbox Replication Service\*"
"\MSExchange Mailbox Replication Service Per Mdb(*)\*"
"\MSExchange MailTips Service\*"
"\MSExchange Managed Folder Assistant\*"
"\MSExchange Message Tracking\*"
"\MSExchange Middle-Tier Storage(*)\*"
"\MSExchange Network Manager\*"
"\MSExchange NSPI RPC Client Connections\*"
"\MSExchange OWA\*"
"\MSExchange Prelicensing Agent\*"
"\MSExchange Protocol Analysis Agent\*"
"\MSExchange Protocol Analysis Background Agent\*"
"\MSExchange Provisioning\*"
"\MSExchange Recipient Cache(_total)\*"
"\MSExchange Recipient Filter Agent\*"
"\MSExchange Replica Seeder\*"
"\MSExchange Replication(_total)\*"
"\MSExchange Resource Booking\*"
"\MSExchange Rights Management\*"
"\MSExchange RMS Agents\*"
"\MSExchange RMS Decryption Agent\*"
"\MSExchange RpcClientAccess\*"
"\MSExchange RpcClientAccess Per Server(*)\*"
"\MSExchange Search Indexer\*"
"\MSExchange Search Indices(_total)\*"
"\MSExchange Secure Mail Transport(_total)\*"
"\MSExchange Sender Filter Agent\*"
"\MSExchange Sender Id Agent\*"
"\MSExchange Sharing Engine\*"
"\MSExchange Store Driver(*)\*"
"\MSExchange Store Interface(*)\*"
"\MSExchange Text Messaging\*"
"\MSExchange Throttling(*)\*"
"\MSExchange Throttling Service Client\*"
"\MSExchange TopN Words Assistant\*"
"\MSExchange Topology(*)\*"
"\MSExchange Transport Rules(*)\*"
"\MSExchange Update Agent\*"
"\MSExchangeAB\*"
"\MSExchangeAL(*)\*"
"\MSExchangeAutodiscover\*"
"\MSExchangeEdgeSync Synchronizer\*"
"\MSExchangeEdgeSync Topology\*"
"\MSExchangeFDS:GM(*)\*"
"\MSExchangeFDS:OAB(*)\*"
"\MSExchangeImap4\*"
"\MSExchangeIS\*"
"\MSExchangeIS Client(*)\*"
"\MSExchangeIS Mailbox(*)\*"
"\MSExchangeIS Public(*)\*"
"\MSExchangePop3\*"
"\MSExchangeTransport Batch Point(*)\*"
"\MSExchangeTransport Component Latency(*)\*"
"\MSExchangeTransport Configuration Cache(*)\*"
"\MSExchangeTransport Database(*)\*"
"\MSExchangeTransport Delivery Failures\*"
"\MSExchangeTransport DeliveryAgent\*"
"\MSExchangeTransport DSN(*)\*"
"\MSExchangeTransport Dumpster\*"
"\MSExchangeTransport IsMemberOfResolver(*)\*"
"\MSExchangeTransport Pickup(*)\*"
"\MSExchangeTransport Queues(*)\*"
"\MSExchangeTransport Resolver(*)\*"
"\MSExchangeTransport Routing(*)\*"
"\MSExchangeTransport ServerAlive(*)\*"
"\MSExchangeTransport Shadow Redundancy(*)\*"
"\MSExchangeTransport SMTPAvailability(*)\*"
"\MSExchangeTransport SMTPReceive(*)\*"
"\MSExchangeTransport SmtpSend(*)\*"
"\MSExchangeUMClientAccess(*)\*"
"\MSExchangeUMMessageWaitingIndicator(*)\*"
"\MSExchangeWS\*"
"\W3SVC_W3WP(*)\*"
"\WAS_W3WP(*)\*"
)
$script:roles += [string]"Full"
Write-Debug "Added Exchange 2010 Full Counters"
}
else{	
	#Common counter list for all roles
	$Counters = @(
"\.NET CLR Exceptions(*)\# of Exceps Thrown / sec"
"\.NET CLR LocksAndThreads(*)\Contention Rate / sec"
"\.NET CLR Memory(*)\% Time in GC"
"\.NET CLR Memory(*)\# Bytes in all Heaps"
"\.NET CLR Memory(*)\# Gen 0 Collections"
"\.NET CLR Memory(*)\# Gen 1 Collections"
"\.NET CLR Memory(*)\# Gen 2 Collections"
"\.NET CLR Memory(*)\# of Pinned Objects"
"\.NET CLR Memory(*)\Allocated Bytes/sec"
"\.NET CLR Memory(*)\Gen 0 heap size"
"\.NET CLR Memory(*)\Gen 1 heap size"
"\.NET CLR Memory(*)\Gen 2 heap size"
"\.NET CLR Memory(*)\Large Object Heap size"
"\.NET CLR Memory(*)\Promoted Memory from Gen 0"
"\.NET CLR Memory(*)\Promoted Memory from Gen 1"
"\LogicalDisk(*)\Avg. Disk Queue Length"
"\LogicalDisk(*)\Avg. Disk sec/Read"
"\LogicalDisk(*)\Avg. Disk sec/Write"
"\LogicalDisk(*)\Disk Reads/sec"
"\LogicalDisk(*)\Disk Transfers/sec"
"\LogicalDisk(*)\Disk Writes/sec"
"\LogicalDisk(*)\% idle time"
"\LogicalDisk(*)\Disk Read Bytes/sec"
"\LogicalDisk(*)\Disk Write Bytes/sec"
"\LogicalDisk(*)\Split IO/Sec"
"\Memory\*"
"\MSExchange ADAccess Caches(*)\Cache Hits/Sec"
"\MSExchange ADAccess Caches(*)\LDAP Searches/Sec"
"\MSExchange ADAccess Domain Controllers(*)\LDAP Read calls/Sec"
"\MSExchange ADAccess Domain Controllers(*)\LDAP Read Time"
"\MSExchange ADAccess Domain Controllers(*)\LDAP Search calls/Sec"
"\MSExchange ADAccess Domain Controllers(*)\LDAP Search Time"
"\MSExchange ADAccess Domain Controllers(*)\LDAP Searches timed out per minute"
"\MSExchange ADAccess Domain Controllers(*)\Long running LDAP operations/Min"
"\MSExchange ADAccess Domain Controllers(*)\Number of outstanding requests"
"\MSExchange ADAccess Local Site Domain Controllers(*)\LDAP Read calls/Sec"
"\MSExchange ADAccess Local Site Domain Controllers(*)\LDAP Read Time"
"\MSExchange ADAccess Local Site Domain Controllers(*)\LDAP Search calls/Sec"
"\MSExchange ADAccess Local Site Domain Controllers(*)\LDAP Search Time"
"\MSExchange ADAccess Local Site Domain Controllers(*)\LDAP Searches timed out per minute"
"\MSExchange ADAccess Local Site Domain Controllers(*)\Long running LDAP operations/Min"
"\MSExchange ADAccess Local Site Domain Controllers(*)\Number of outstanding requests"
"\MSExchange ADAccess Processes(*)\LDAP Read calls/Sec"
"\MSExchange ADAccess Processes(*)\LDAP Read Time"
"\MSExchange ADAccess Processes(*)\LDAP Search Time"
"\MSExchange ADAccess Processes(*)\LDAP Search calls/Sec"
"\MSExchange ADAccess Processes(*)\LDAP Timeout Errors/Sec"
"\MSExchange ADAccess Processes(*)\Long running LDAP operations/Min"
"\MSExchange ADAccess Processes(*)\Number of outstanding requests"
"\Netlogon(*)\*"
"\Network Interface(*)\Bytes Received/sec"
"\Network Interface(*)\Bytes Sent/sec"
"\Network Interface(*)\Bytes Total/sec"
"\Network Interface(*)\Current Bandwidth"
"\Network Interface(*)\Output Queue Length"
"\Network Interface(*)\Packets Outbound Errors"
"\Paging File(_Total)\% Usage"
"\PhysicalDisk(*)\Avg. Disk Queue Length"
"\PhysicalDisk(*)\Avg. Disk sec/Read"
"\PhysicalDisk(*)\Avg. Disk sec/Write"
"\PhysicalDisk(*)\% idle time"
"\PhysicalDisk(*)\Disk Reads/sec"
"\PhysicalDisk(*)\Disk Read Bytes/sec"
"\PhysicalDisk(*)\Disk Transfers/sec"
"\PhysicalDisk(*)\Disk Write Bytes/sec"
"\PhysicalDisk(*)\Disk Writes/sec"
"\PhysicalDisk(*)\Split IO/Sec"
"\Process(*)\*"
"\Processor(*)\*"
"\Processor Information(*)\*"
"\Redirector\*"
"\Server\*"
"\System\*"
"\TCPv4\*"
"\TCPv6\*"
)
Write-Debug "Added Common Counters"
}
# Add $Counters
if ($threads -eq $true)
{
	$Counters += [string]"\Thread(*)\*"
	Write-Debug "Added Threads Counters"
}
$Counters = $Counters | Sort-Object | Select-Object -Unique
$script:Counterlist = $Counters

if (!$full){
	if (!$GetServer.IsEdgeServer){
	#Add Store Interface Counters
	$StoreInterfaceCounters = @(
"\MSExchange Store Interface(*)\ConnectionCache active connections"
"\MSExchange Store Interface(*)\ConnectionCache num caches"
"\MSExchange Store Interface(*)\ConnectionCache out of limit creations"
"\MSExchange Store Interface(*)\ConnectionCache total capacity"
"\MSExchange Store Interface(*)\ExRPCConnection creation events"
"\MSExchange Store Interface(*)\ExRPCConnection disposal events"
"\MSExchange Store Interface(*)\ExRPCConnection outstanding"
"\MSExchange Store Interface(*)\ROP Requests outstanding"
"\MSExchange Store Interface(*)\RPC Latency average (msec)"
"\MSExchange Store Interface(*)\RPC Requests failed (%)"
"\MSExchange Store Interface(*)\RPC Requests outstanding"
"\MSExchange Store Interface(*)\RPC Requests sent/sec"
"\MSExchange Store Interface(*)\RPC Slow requests (%)"
)
	
	$Counters += $StoreInterfaceCounters
	}
#	if ($threads -eq $true)
#	{
#		$Counters += [string]"\Thread(*)\*"
#		Write-Debug "Added Threads Counters"
#	}
	if ($GetServer.IsMailboxServer -eq $true){
		$script:roles += [string]"Mbx"
		#MBX Counter list
		if ($Exchange2007){		
		$MBXCounterList = @(
"\MSExchange Assistants(*)\Average Event Processing Time In seconds"
"\MSExchange Assistants(*)\Average Event Queue Time in seconds"
"\MSExchange Assistants(*)\Average Mailbox Processing Time In seconds"
"\MSExchange Assistants(*)\Events in queue"
"\MSExchange Assistants(*)\Events Polled/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Reads/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Writes/sec"
"\MSExchange Database ==> Instances(*)\I/O Log Reads/sec"
"\MSExchange Database ==> Instances(*)\I/O Log Writes/sec"
"\MSExchange Database ==> Instances(*)\Log Generation Checkpoint Depth"
"\MSExchange Database ==> Instances(*)\Log Record Stalls/sec"
"\MSExchange Database ==> Instances(*)\Log Threads Waiting"
"\MSExchange Database ==> Instances(*)\Version buckets allocated"
"\MSExchange Database(Information Store)\Database Cache % Hit"
"\MSExchange Database(Information Store)\Database Cache Size (MB)"
"\MSExchange Database(Information Store)\Database Page Fault Stalls/sec"
"\MSExchange Database(Information Store)\I/O Database Reads Average Latency"
"\MSExchange Database(Information Store)\I/O Database Writes Average Latency"
"\MSExchange Database(Information Store)\Log Record Stalls/sec"
"\MSExchange Database(Information Store)\Log Threads Waiting"
"\MSExchange Database(Information Store)\Version buckets allocated"
"\MSExchange Replication(*)\ReplayQueueLength"
"\MSExchange Replication(*)\CopyQueueLength"
"\MSExchange Resource Booking\Average Resource Booking Processing Time"
"\MSExchange Resource Booking\Requests Failed"
"\MSExchange Search Indexer\Average Batch Latency"
"\MSExchange Search Indexer\Number of Databases Being Crawled"
"\MSExchange Search Indexer\Number of Databases Being Indexed"
"\MSExchange Search Indexer\Number of Indexed Databases Being Kept Up-to-Date by Notifications"
"\MSExchange Search Indices(*)\Age of the Last Notification Indexed"
"\MSExchange Search Indices(*)\Average Document Indexing Time"
"\MSExchange Search Indices(*)\Average Latency of RPCs Used to Obtain Content"
"\MSExchange Search Indices(*)\Average Latency of RPCs to get notifications"
"\MSExchange Search Indices(*)\Average Latency of RPCs During Crawling"
"\MSExchange Search Indices(*)\Full Crawl Mode Status"
"\MSExchange Search Indices(*)\Number of Create Notifications/sec"
"\MSExchange Search Indices(*)\Number of Items in a Notification Queue"
"\MSExchange Search Indices(*)\Number of Mailboxes Left to Crawl"
"\MSExchange Search Indices(*)\Number of Outstanding Batches"
"\MSExchange Search Indices(*)\Number of Outstanding Documents"
"\MSExchange Search Indices(*)\Number of Recently Moved Mailboxes Being Crawled"
"\MSExchange Search Indices(*)\Number of Retries"
"\MSExchange Search Indices(*)\Number of Update Notifications/sec"
"\MSExchange Search Indices(*)\Throttling Delay Value"
"\MSExchangeAL(_Total)\LDAP Results/sec"
"\MSExchangeAL(_Total)\LDAP Search calls"
"\MSExchangeAL(_Total)\LDAP Search calls/sec"
"\MSExchangeMailSubmission(*)\Hub Servers In Retry"
"\MSExchangeMailSubmission(*)\Successful Submissions Per Second"
"\MSExchangeMailSubmission(*)\Failed Submissions Per Second"
"\MSExchangeIS Client(*)\*"
"\MSExchangeIS Mailbox(*)\Folder opens/sec"
"\MSExchangeIS Mailbox(*)\Logon Operations/sec"
"\MSExchangeIS Mailbox(*)\Message Opens/sec"
"\MSExchangeIS Mailbox(*)\Slow FindRow Rate"
"\MSExchangeIS Mailbox(*)\Search Task Rate"
"\MSExchangeIS Mailbox(*)\Categorization Count"
"\MSExchangeIS Mailbox(_Total)\Active Client Logons"
"\MSExchangeIS Mailbox(_Total)\Client Logons"
"\MSExchangeIS Mailbox(_Total)\Local delivery rate"
"\MSExchangeIS Mailbox(_Total)\Messages Delivered/sec"
"\MSExchangeIS Mailbox(_Total)\Messages Queued For Submission"
"\MSExchangeIS Mailbox(_Total)\Messages Sent/sec"
"\MSExchangeIS Mailbox(_Total)\Messages Submitted/sec"
"\MSExchangeIS Public(_Total)\Active Client Logons"
"\MSExchangeIS Public(_Total)\Client Logons"
"\MSExchangeIS Public(_Total)\Messages Delivered/sec"
"\MSExchangeIS Public(_Total)\Messages Queued For Submission"
"\MSExchangeIS Public(_Total)\Messages Sent/sec"
"\MSExchangeIS Public(_Total)\Messages Submitted/sec"
"\MSExchangeIS\Active User Count"
"\MSExchangeIS\Client: Latency > 2 sec RPCs"
"\MSExchangeIS\Client: Latency > 5 sec RPCs"
"\MSExchangeIS\Client: Latency > 10 sec RPCs"
"\MSExchangeIS\Client: RPCs Failed"
"\MSExchangeIS\Client: RPCs Failed: Server Too Busy / sec"
"\MSExchangeIS\Slow QP Threads"
"\MSExchangeIS\Slow Search Threads"
"\MSExchangeIS\RPC Averaged Latency"
"\MSExchangeIS\RPC Client Backoff/sec"
"\MSExchangeIS\RPC Num. of Slow Packets"
"\MSExchangeIS\RPC Operations/sec"
"\MSExchangeIS\RPC Requests"
"\MSExchangeIS\Virus Scan Files Quarantined/sec"
"\MSExchangeIS\Virus Scan Files Scanned/sec"
"\MSExchangeIS\Virus Scan Messages Processed/sec"
"\MSExchangeIS\Virus Scan Queue Length"
"\MSExchangeIS\VM Largest Block Size"
"\MSExchangeIS\VM Total 16MB Free Blocks"
"\MSExchangeIS\VM Total Free Blocks"
"\MSExchangeIS\VM Total Large Free Block Bytes"
)
		Write-Debug "Added Exchange 2007 Mailbox Counters"
		}
	if ($Exchange2010){
	$MBXCounterList = @(
"\MSExchange Active Manager(*)\Database Mounted"
"\MSExchange Approval Assistant\Average Approval Assistant Processing Time"
"\MSExchange Approval Assistant\Last Approval Assistant Processing Time"
"\MSExchange Assistants - Per Assistant(*)\Average Event Processing Time In Seconds"
"\MSExchange Assistants - Per Assistant(*)\Average Event Queue Time In Seconds"
"\MSExchange Assistants - Per Assistant(*)\Elapsed Time Since Last Event Queued"
"\MSExchange Assistants - Per Assistant(*)\Events in Queue"
"\MSExchange Assistants - Per Assistant(*)\Events Processed/sec"
"\MSExchange Assistants - Per Assistant(*)\Handled Exceptions"
"\MSExchange Assistants - Per Database(*)\Average Event Processing Time In seconds"
"\MSExchange Assistants - Per Database(*)\Average Mailbox Processing Time In seconds"
"\MSExchange Assistants - Per Database(*)\Events in queue"
"\MSExchange Assistants - Per Database(*)\Events Polled/sec"
"\MSExchange Assistants - Per Database(*)\Mailboxes processed/sec"
"\MSExchange Calendar Attendant\Average Calendar Attendant Processing Time"
"\MSExchange Calendar Attendant\Requests Failed"
"\MSExchange Calendar Notifications Assistant\Average update processing latency (milliseconds)"
"\MSExchange Database ==> Instances(*)\Database Maintenance Duration"
"\MSExchange Database ==> Instances(*)\Defragmentation Tasks"
"\MSExchange Database ==> Instances(*)\Defragmentation Tasks Pending"
"\MSExchange Database ==> Instances(*)\I/O Database Reads Average Latency"
"\MSExchange Database ==> Instances(*)\I/O Database Reads (Attached)/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Reads (Recovery)/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Reads/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Writes Average Latency"
"\MSExchange Database ==> Instances(*)\I/O Database Writes (Attached)/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Writes (Recovery)/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Writes/sec"
"\MSExchange Database ==> Instances(*)\I/O Log Reads/sec"
"\MSExchange Database ==> Instances(*)\I/O Log Reads Average Latency"
"\MSExchange Database ==> Instances(*)\I/O Log Writes/sec"
"\MSExchange Database ==> Instances(*)\I/O Log Writes Average Latency"
"\MSExchange Database ==> Instances(*)\Log Bytes Write/sec"
"\MSExchange Database ==> Instances(*)\Log Checkpoint Maintenance Outstanding IO Max"
"\MSExchange Database ==> Instances(*)\Log Generation Checkpoint Depth"
"\MSExchange Database ==> Instances(*)\Log Record Stalls/sec"
"\MSExchange Database ==> Instances(*)\Log Threads Waiting"
"\MSExchange Database ==> Instances(*)\Sessions % Used"
"\MSExchange Database ==> Instances(*)\Version buckets allocated"
"\MSExchange Database(*)\Database Cache % Dehydrated"
"\MSExchange Database(*)\Database Cache % Hit" 
"\MSExchange Database(*)\Database Cache Size Effective (MB)"
"\MSExchange Database(*)\Database Cache Size Resident (MB)"
"\MSExchange Database(Information Store)\Database Cache % Hit"
"\MSExchange Database(Information Store)\Database Cache Size (MB)"
"\MSExchange Database(Information Store)\Database Page Fault Stalls/sec"
"\MSExchange Database(Information Store)\I/O Database Writes (Attached) Average Latency" 
"\MSExchange Database(Information Store)\I/O Database Writes (Recovery) Average Latency"
"\MSExchange Database(Information Store)\I/O Database Writes Average Latency"
"\MSExchange Database(Information Store)\I/O Log Writes Average Latency"
"\MSExchange Database(Information Store)\Log Record Stalls/sec"
"\MSExchange Database(Information Store)\Log Threads Waiting"
"\MSExchange Database(Information Store)\Version Buckets Allocated"
"\MSExchange FreeBusy Assistant\Average FreeBusy Assistant Processing Time"
"\MSExchange FreeBusy Assistant\Events processed by freebusy assistant (sec)"
"\MSExchange Junk E-mail Options Assistant\Recipients updated per second"
"\MSExchange Mail Submission(*)\Temporary Submission Failures/sec"
"\MSExchange Mailbox Replication Service Per Mdb(*)\Last Scan: Duration (msec)"
"\MSExchange Network Manager(*)\Avg Log Copy Latency (msec)"
"\MSExchange Network Manager(*)\Avg Seeding Latency (msec)"
"\MSExchange Network Manager(*)\Log Copy KB Received/Sec"
"\MSExchange Network Manager(*)\Log Copy KB Sent/Sec"
"\MSExchange Network Manager(*)\Seeder KB Received/Sec"
"\MSExchange Network Manager(*)\Seeder KB Sent/Sec"
"\MSExchange Replication(*)\Avg Log Copy Latency (msec)"
"\MSExchange Replication(*)\CopyQueueLength"
"\MSExchange Replication(_Total)\Log Copying is Not Keeping Up"
"\MSExchange Replication(*)\Log Generation Rate on Source (generations/sec)"
"\MSExchange Replication(*)\Log Replay Rate (generations/sec)"
"\MSExchange Replication(*)\ReplayQueueLength"
"\MSExchange Replication(_Total)\Log Replay is Not Keeping Up"
"\MSExchange Replication(_Total)\Log Copy KB/Sec"
"\MSExchange Resource Booking\Average Resource Booking Processing Time"
"\MSExchange Resource Booking\Requests Failed"
"\MSExchange Search Indexer\Average Batch Latency"
"\MSExchange Search Indexer\Number of Databases Being Crawled"
"\MSExchange Search Indexer\Number of Databases Being Indexed"
"\MSExchange Search Indexer\Number of Disabled Databases"
"\MSExchange Search Indexer\Number of Paused Databases"
"\MSExchange Search Indexer\Number of Indexed Databases Being Kept Up-to-Date by Notifications"
"\MSExchange Search Indices(*)\Age of the Last Notification Indexed"
"\MSExchange Search Indices(*)\Average Document Indexing Time"
"\MSExchange Search Indices(*)\Average Latency of RPCs to get notifications"
"\MSExchange Search Indices(*)\Average Latency of RPCs During Crawling"
"\MSExchange Search Indices(*)\Full Crawl Mode Status"
"\MSExchange Search Indices(*)\Number of Items in a Notification Queue"
"\MSExchange Search Indices(*)\Number of Mailboxes Left to Crawl"
"\MSExchange Search Indices(*)\Number of Outstanding Batches"
"\MSExchange Search Indices(*)\Number of Outstanding Documents"
"\MSExchange Search Indices(*)\Recent Average Latency of RPCs Used to Obtain Content"
"\MSExchange Search Indices(*)\Throttling Delay Value"
"\MSExchange Search Indices(*)\Time Since Last Notification Was Indexed"
"\MSExchange Search Indices(*)\Total Time Taken For Indexing Protected Messages"
"\MSExchange Search Indices(*)\Number of Create Notifications/sec"
"\MSExchange Search Indices(*)\Number of InTransit Mailboxes Being Indexed on this Destination Database"
"\MSExchange Search Indices(*)\Number of Retries"
"\MSExchange Search Indices(*)\Number of Update Notifications/sec"
"\MSExchange TopN Words Assistant\Time to Process Last Mailbox in Milliseconds"
"\MSExchange Transport Sync Manager\Failed Submissions"
"\MSExchangeAL(_Total)\LDAP Results/sec"
"\MSExchangeAL(_Total)\LDAP Search Calls"
"\MSExchangeAL(_Total)\LDAP Search Calls/sec"
"\MSExchangeIS Client(*)\*"
"\MSExchangeIS Mailbox(*)\Active RPC Thread Limit"
"\MSExchangeIS Mailbox(*)\Active RPC Threads"
"\MSExchangeIS Mailbox(*)\Exchange Search Slow First Batch"
"\MSExchangeIS Mailbox(*)\ExchangeSearch First Batch"
"\MSExchangeIS Mailbox(*)\ExchangeSearch Queries"
"\MSExchangeIS Mailbox(*)\ExchangeSearch Ten More"
"\MSExchangeIS Mailbox(*)\ExchangeSearch Zero Results Queries"
"\MSExchangeIS Mailbox(*)\Folder opens/sec"
"\MSExchangeIS Mailbox(*)\Last Query Time"
"\MSExchangeIS Mailbox(*)\Local delivery rate"
"\MSExchangeIS Mailbox(*)\Logon Operations/sec"
"\MSExchangeIS Mailbox(*)\Message Opens/sec"
"\MSExchangeIS Mailbox(*)\Messages Delivered/sec"
"\MSExchangeIS Mailbox(*)\Messages Queued For Submission"
"\MSExchangeIS Mailbox(*)\Messages Sent/sec"
"\MSExchangeIS Mailbox(*)\Messages Submitted/sec"
"\MSExchangeIS Mailbox(*)\Quarantined Mailbox Count"
"\MSExchangeIS Mailbox(*)\Mailbox Replication Read Connections"
"\MSExchangeIS Mailbox(*)\Mailbox Replication Write Connections"
"\MSExchangeIS Mailbox(*)\RPC Average Latency"
"\MSExchangeIS Mailbox(*)\RPC Average Latency (Client)"
"\MSExchangeIS Mailbox(*)\Search Task Rate"
"\MSExchangeIS Mailbox(*)\Slow FindRow Rate"
"\MSExchangeIS Mailbox(*)\Store Only Queries"
"\MSExchangeIS Mailbox(*)\Store Only Query Ten More"
"\MSExchangeIS Mailbox(_Total)\Active Client Logons"
"\MSExchangeIS Mailbox(_Total)\Client Logons"
"\MSExchangeIS Mailbox(_Total)\Delivery Blocked: Low Database Space"
"\MSExchangeIS Mailbox(_Total)\Delivery Blocked: Low Log Disk Space"
"\MSExchangeIS Public(_Total)\Active Client Logons"
"\MSExchangeIS Public(_Total)\Client Logons"
"\MSExchangeIS Public(_Total)\Messages Delivered/sec"
"\MSExchangeIS Public(_Total)\Messages Queued For Submission"
"\MSExchangeIS Public(_Total)\Messages Sent/sec"
"\MSExchangeIS Public(_Total)\Messages Submitted/sec"
"\MSExchangeIS Public(_Total)\Replication Receive Queue Size"
"\MSExchangeIS\% Connections"    
"\MSExchangeIS\% RPC Threads"    
"\MSExchangeIS\Active User Count"
"\MSExchangeIS\Async RPC Requests"
"\MSExchangeIS\CI QP Threads"
"\MSExchangeIS\Client: Latency > 2 sec RPCs"
"\MSExchangeIS\Client: Latency > 5 sec RPCs"
"\MSExchangeIS\Client: Latency > 10 sec RPCs"
"\MSExchangeIS\Client: RPCs Failed"
"\MSExchangeIS\Client: RPCs Failed: Server Too Busy/sec"
"\MSExchangeIS\Minimsg created for views/sec"
"\MSExchangeIS\Minimsg Msg table seeks/sec"
"\MSExchangeIS\MsgView Records Deleted/sec"
"\MSExchangeIS\MsgView Records Inserted/sec"
"\MSExchangeIS\MsgView table Create/sec"
"\MSExchangeIS\RPC Averaged Latency"
"\MSExchangeIS\RPC Client Backoff/sec"
"\MSExchangeIS\RPC Num. of Slow Packets"
"\MSExchangeIS\RPC Operations/sec"
"\MSExchangeIS\RPC Request Timeout Detected"
"\MSExchangeIS\RPC Requests"
"\MSExchangeIS\Slow QP Threads"
"\MSExchangeIS\Slow Search Threads"
"\MSExchangeIS\User Count"
"\MSExchangeIS\View Cleanup Categorization Index Deletions/sec"
"\MSExchangeIS\View Cleanup DVU Entry Deletions/sec"
"\MSExchangeIS\View Cleanup Restriction Index Deletions/sec"
"\MSExchangeIS\View Cleanup Search Index Deletions/sec"
"\MSExchangeIS\View Cleanup Sort Index Deletions/sec"
"\MSExchangeIS\View Cleanup Tasks Nullified/sec"
"\MSExchangeIS\View Cleanup Tasks/sec"
"\MSExchangeIS\Virus Scan Files Quarantined/sec"
"\MSExchangeIS\Virus Scan Files Scanned/sec"
"\MSExchangeIS\Virus Scan Messages Processed/sec"
"\MSExchangeIS\Virus Scan Queue Length"
"\MSExchangeIS\VM Largest Block Size"
"\MSExchangeIS\VM Total Free Blocks"
"\MSExchangeIS\VM Total Large Free Block Bytes"
"\MSExchangeIS\VM Total 16MB Free Blocks"
"\MSExchange Mail Submission(*)\Failed Submissions Per Second"
"\MSExchange Mail Submission(*)\Hub Servers In Retry"
"\MSExchange Mail Submission(*)\Hub Transport Servers Percent Active"
"\MSExchange Mail Submission(*)\Successful Submissions Per Second"
)
	Write-Debug "Added Exchange 2010 Mailbox Counters"
	}
		$Counters += $MBXCounterList
	#Add Extended Counters
	if ($StoreExtendedon)
	{
	$StoreExtended = @(
"\MSExchangeIS Mailbox(*)\ImportDeleteOpRate"
"\MSExchangeIS Mailbox(*)\SaveChangesMessageOpRate"
"\MSExchangeIS Mailbox(*)\SaveChangesAttachOpRate"
"\MSExchangeIS Mailbox(*)\FindRow operations/sec"
"\MSExchangeIS Mailbox(*)\Restrict Operations/sec"
"\MSExchangeIS Mailbox(*)\QueryPosition Operations/sec"
"\MSExchangeIS Mailbox(*)\SeekRow Operations/sec"
"\MSExchangeIS Mailbox(*)\SeekRowBookMark Operations/sec"
"\MSExchangeIS Mailbox(*)\QueryRowsOpRate"
"\MSExchangeIS Mailbox(*)\SetSearchCriteriaOpRate"
"\MSExchangeIS Mailbox(*)\GetSearchCriteriaOpRate"	
)
	Write-Debug "Added Exchange Store Extended Counters"
	}
	if ($ESEExtendedon -and $Exchange2007)
	{
	$ESEExtended = @(
"\MSExchange Database(*)\Database Cache % Clean"
"\MSExchange Database(*)\Database Cache % Available"
"\MSExchange Database(*)\Database Cache % Versioned"
"\MSExchange Database(*)\Threads Blocked/sec"
"\MSExchange Database(*)\Threads Blocked"
"\MSExchange Database ==> Instances(*)\FCB Asynchronous Scan/sec"
"\MSExchange Database ==> Instances(*)\FCB Asynchronous Purge/sec"
"\MSExchange Database ==> Instances(*)\FCB Cache % Hit"
"\MSExchange Database ==> Instances(*)\FCB Cache Allocated"
"\MSExchange Database ==> Instances(*)\FCB Cache Available"
"\MSExchange Database ==> Instances(*)\FCB Cache Maximum"
"\MSExchange Database ==> Instances(*)\FCB Cache Preferred"
"\MSExchange Database ==> Instances(*)\Database Pages Repeatedly Written/sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Average Log Bytes"
"\MSExchange Database ==> Instances(*)\Online Defrag Log Records/sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Pages Dirtied/sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Pages Preread/sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Pages Read/sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Pages Re-Dirtied/sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Pages Referenced/sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Pages Freed/Sec"
"\MSExchange Database ==> Instances(*)\Online Defrag Data Moves/Sec"
)
	Write-Debug "Added Exchange 2007 ESE Extended Counters"
	}
	if ($ESEExtendedon -and $Exchange2010)
	{
	$ESEExtended = @(
"\MSExchange Database(*)\Database Cache % Clean"
"\MSExchange Database(*)\Database Cache % Available"
"\MSExchange Database(*)\Database Cache % Resident"
"\MSExchange Database(*)\Database Cache % Versioned"
"\MSExchange Database(*)\Database Cache Size Target"
"\MSExchange Database(*)\Database Cache Lifetime"
"\MSExchange Database(*)\Threads Blocked/sec"
"\MSExchange Database(*)\Threads Blocked"
"\MSExchange Database ==> Instances(*)\FCB Asynchronous Scan/sec"
"\MSExchange Database ==> Instances(*)\FCB Asynchronous Purge/sec"
"\MSExchange Database ==> Instances(*)\FCB Cache % Hit"
"\MSExchange Database ==> Instances(*)\FCB Cache Allocated"
"\MSExchange Database ==> Instances(*)\FCB Cache Available"
"\MSExchange Database ==> Instances(*)\FCB Cache Maximum"
"\MSExchange Database ==> Instances(*)\FCB Cache Preferred"
"\MSExchange Database ==> Instances(*)\Log Checkpoint Maintenance Outstanding IO Max"
"\MSExchange Database ==> Instances(*)\Database Maintenance IO Reads/sec"
"\MSExchange Database ==> Instances(*)\Defragmentation Tasks Completed/sec"
"\MSExchange Database ==> Instances(*)\Database Pages Flushed (Checkpoint)/sec"
"\MSExchange Database ==> Instances(*)\Database Pages Flushed (Scavenge)/sec"
"\MSExchange Database ==> Instances(*)\Database Pages Repeatedly Written/sec"	
)
	Write-Debug "Added Exchange 2010 ESE Extended Counters"
	}
	$Counters += $StoreExtended
	$Counters += $ESEExtended
	}
	if ($GetServer.IsClientAccessServer -eq $true){
		$script:roles += [string]"Cas"
	# CAS Counter list
	if ($Windows2008 -or $Windows2008R2){
		$HTTPCounters = @(
"\HTTP Service Request Queues(*)\ArrivalRate"
"\HTTP Service Request Queues(*)\CurrentQueueSize"
"\HTTP Service Request Queues(*)\RejectionRate"
"\RPC/HTTP Proxy\Attempted RPC Load Balancing Broker Requests per Second"
"\RPC/HTTP Proxy\Attempted RPC Load Balancing Decisions per Second"
"\RPC/HTTP Proxy\Current Number of Incoming RPC over HTTP Connections"
"\RPC/HTTP Proxy\Current Number of Unique Users"
"\RPC/HTTP Proxy\Failed RPC Load Balancing Broker Requests per Second"
"\RPC/HTTP Proxy\Failed RPC Load Balancing Decisions per Second"
"\RPC/HTTP Proxy\RPC/HTTP Requests per Second"
"\RPC/HTTP Proxy\Number of Back-End Connection Attempts per Second"
"\RPC/HTTP Proxy\Number of Failed Back-End Connection Attempts per Second"
"\RPC/HTTP Proxy\Total Incoming Bandwidth from Back-EndServers"
"\RPC/HTTP Proxy\Total Outgoing Bandwidth to Back-EndServers"
)
	$Counters += $HTTPCounters
	Write-Debug "Added Windows 2008 HTTP Counters"
	}
	if ($Exchange2007){
		$CASCounterList = @(
"\ASP.NET\Application Restarts"
"\ASP.NET Applications(*)\Requests In Application Queue"
"\ASP.NET Applications(*)\Requests Executing"
"\ASP.NET\Request Execution Time"
"\ASP.NET\Request Wait Time"
"\ASP.NET\Requests Current"
"\ASP.NET\Requests Queued"
"\ASP.NET\Requests Rejected"
"\ASP.NET\Worker Process Restarts"
"\ASP.NET Apps v2.0.50727(*)\Requests In Application Queue"
"\ASP.NET Apps v2.0.50727(_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync)\Request Wait Time"
"\ASP.NET Apps v2.0.50727(_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync)\Requests Executing"
"\ASP.NET Apps v2.0.50727(_LM_W3SVC_1_ROOT_Microsoft-Server-ActiveSync)\Requests In Application Queue"
"\MSExchangeAutodiscover\Requests/sec"
"\MSExchangeImap4(_total)\Current Connections"
"\MSExchangePop3(_total)\Connections Current"
"\MSExchangePop3(_total)\DELE Rate"
"\MSExchangePop3(_total)\RETR Rate"
"\MSExchangePop3(_total)\UIDL Rate"
"\MSExchangeWS\Items Read/sec"
"\MSExchangeWS\Proxy average response time"
"\MSExchangeWS\Requests/sec"
"\MSExchange ActiveSync\Average Ping Time"
"\MSExchange ActiveSync\Average Request Time"
"\MSExchange ActiveSync\Busy Threads"
"\MSExchange ActiveSync\Heartbeat Interval"
"\MSExchange ActiveSync\Incoming Proxy Requests Total"
"\MSExchange ActiveSync\Ping Commands Dropped/sec"
"\MSExchange ActiveSync\Ping Commands Pending"
"\MSExchange ActiveSync\Requests Queued"
"\MSExchange ActiveSync\Requests/sec"
"\MSExchange ActiveSync\Sync Commands Pending"
"\MSExchange ActiveSync\Sync Commands/sec"
"\MSExchange ActiveSync\Wrong CAS Proxy Requests Total"
"\MSExchange ADAccess Domain Controllers(*)\LDAP Searches timed out per minute"
"\MSExchange Availability Service\Availability Requests (sec)"
"\MSExchange Availability Service\Average Number of Mailboxes Processed per Request"
"\MSExchange Availability Service\Average Time to Process a Cross-Forest Free Busy Request"
"\MSExchange Availability Service\Average Time to Process a Cross-Site Free Busy Request"
"\MSExchange Availability Service\Average Time to Process a Free Busy Request"
"\MSExchange Availability Service\Average Time to Process a Meeting Suggestions Request"
"\MSExchange Availability Service\Public Folder Queries (sec)"
"\MSExchange OWA\AS Queries Failure %"
"\MSExchange OWA\Average Search Time"
"\MSExchange OWA\Average Response Time"
"\MSExchange OWA\Current Proxy Users"
"\MSExchange OWA\Current Unique Users"
"\MSExchange OWA\Current Unique Users Light"
"\MSExchange OWA\Current Unique Users Premium"
"\MSExchange OWA\Failed Requests/sec"
"\MSExchange OWA\Store Logon Failure %"
"\MSExchange OWA\Logons/sec"
"\MSExchange OWA\Proxy Response Time Average"
"\MSExchange OWA\Proxy User Requests/sec"
"\MSExchange OWA\Proxy User Requests"
"\MSExchange OWA\Requests/sec"
"\Web Service(_Total)\Bytes Received/sec"
"\Web Service(_Total)\Bytes Sent/sec"
"\Web Service(_Total)\Bytes Total/sec"
"\Web Service(_Total)\Connection Attempts/sec"
"\Web Service(_Total)\Current Connections"
"\Web Service(_Total)\Get Requests/sec"
"\Web Service(_Total)\ISAPI Extension Requests/sec"
"\Web Service(_Total)\Other Request Methods/sec"
)
		Write-Debug "Added Exchange 2007 CAS Counters"
		}
	if ($Exchange2010){
		$CASCounterList = @(
"\ASP.NET Applications(*)\Requests In Application Queue"
"\ASP.NET Applications(*)\Requests Executing"
"\ASP.NET\Application Restarts"
"\ASP.NET\Request Execution Time"
"\ASP.NET\Request Wait Time"
"\ASP.NET\Requests Current"
"\ASP.NET\Requests Queued"
"\ASP.NET\Requests Rejected"
"\ASP.NET\Worker Process Restarts"
"\GALSync\Client reported total time used for Mailbox creation in milliseconds" 
"\MSExchange ActiveSync\Average Ping Time"
"\MSExchange ActiveSync\Average Request Time"
"\MSExchange ActiveSync\Busy Threads"
"\MSExchange ActiveSync\Incoming Proxy Requests Total"
"\MSExchange ActiveSync\Ping Commands Dropped/sec"
"\MSExchange ActiveSync\Ping Commands Pending"
"\MSExchange ActiveSync\Requests Queued"
"\MSExchange ActiveSync\Requests/sec"
"\MSExchange ActiveSync\Sync Commands Pending"
"\MSExchange ActiveSync\Sync Commands/sec"
"\MSExchange ActiveSync\Wrong CAS Proxy Requests Total"
"\MSExchange Availability Service\Average Time to Map External Caller to Internal Identity"
"\MSExchange Availability Service\Average Time to Process a Cross-Forest Free Busy Request"
"\MSExchange Availability Service\Average Time to Process a Cross-Site Free Busy Request"
"\MSExchange Availability Service\Average Time to Process a Federated Free Busy Request"
"\MSExchange Availability Service\Average Time to Process a Free Busy Request"
"\MSExchange Availability Service\Average Time to Process a Meeting Suggestions Request"
"\MSExchange Availability Service\Client Reported Failures - Total"
"\MSExchange Availability Service\Cross-Site Calendar Failures (sec)"
"\MSExchange Availability Service\Cross-Site Calendar Queries (sec)"
"\MSExchange Availability Service\Successful Client Reported Requests - Over 20 seconds"
"\MSExchange Control Panel\ASP.Net Request Failures"
"\MSExchange Control Panel\Explicit Sign-On Inbound Proxy Requests/sec"
"\MSExchange Control Panel\Explicit Sign-On Inbound Proxy Sessions/sec"
"\MSExchange Control Panel\Explicit Sign-On Outbound Proxy Requests/sec"
"\MSExchange Control Panel\Explicit Sign-On Outbound Proxy Sessions/sec"
"\MSExchange Control Panel\Explicit Sign-On Standard RBAC Requests/sec"
"\MSExchange Control Panel\Explicit Sign-On Standard RBAC Sessions/sec"
"\MSExchange Control Panel\Inbound Proxy Requests/sec"
"\MSExchange Control Panel\Inbound Proxy Sessions/sec"
"\MSExchange Control Panel\Outbound Proxy Requests - Average Response Time"
"\MSExchange Control Panel\Outbound Proxy Requests/sec"
"\MSExchange Control Panel\Outbound Proxy Sessions/sec"
"\MSExchange Control Panel\PowerShell Runspaces - Activations/sec"
"\MSExchange Control Panel\PowerShell Runspaces - Average Active Time"
"\MSExchange Control Panel\PowerShell Runspaces/sec"
"\MSExchange Control Panel\RBAC Sessions/sec"
"\MSExchange Control Panel\Requests - Activations/sec"
"\MSExchange Control Panel\Requests - Average Response Time"
"\MSExchange Control Panel\Web Service Request Failures"
"\MSExchange Mailbox Replication Service Per Mdb(*)\Active Moves: Moves in Completion State"
"\MSExchange Mailbox Replication Service Per Mdb(*)\Active Moves: Moves in Initial Seeding State"
"\MSExchange Mailbox Replication Service Per Mdb(_total)\Active Moves: Moves in Transient Failure State"
"\MSExchange Mailbox Replication Service Per Mdb(*)\Active Moves: Stalled Moves (Content Indexing)"
"\MSExchange Mailbox Replication Service Per Mdb(*)\Active Moves: Stalled Moves (Database Replication)"
"\MSExchange Mailbox Replication Service Per Mdb(*)\Active Moves: Stalled Moves Total"
"\MSExchange Mailbox Replication Service Per Mdb(*)\Active Moves: Total Moves"
"\MSExchange Mailbox Replication Service Per Mdb(_total)\Active Moves: Transfer Rate (KB/sec)"
"\MSExchange Mailbox Replication Service Per Mdb(*)\MDB Health: Content Indexing Lagging"
"\MSExchange Mailbox Replication Service Per Mdb(*)\MDB Health: Database Replication Lagging"
"\MSExchange Mailbox Replication Service Per Mdb(*)\MDB Health: Scan Failure"
"\MSExchange Mailbox Replication Service\Last Scan Duration (msec)"
"\MSExchange MailTips Service\GetMailTips Average Response Time for GroupMetrics Queries"
"\MSExchange MailTips Service\GetMailTips Average Response Time"
"\MSExchange MailTips Service\GetMailTipsConfiguration Average Response Time"
"\MSExchange MailTips Service\GetServiceConfiguration average response time"
"\MSExchange MailTips Service\MailTips Queries Answered Within One Second"
"\MSExchange MailTips Service\MailTips Queries Answered Within Ten Seconds"
"\MSExchange MailTips Service\MailTips Queries Answered Within Three Seconds"
"\MSExchange OWA\Average Response Time"
"\MSExchange OWA\Average Search Time"
"\MSExchange OWA\AS Queries Failure %"
"\MSExchange OWA\Current Proxy Users"
"\MSExchange OWA\Current Unique Users"
"\MSExchange OWA\Current Unique Users Light"
"\MSExchange OWA\Current Unique Users Premium"
"\MSExchange OWA\Failed Requests/sec"
"\MSExchange OWA\Store Logon Failure %"
"\MSExchange OWA\Logons/sec"
"\MSExchange OWA\Proxy Response Time Average"
"\MSExchange OWA\Proxy User Requests/sec"
"\MSExchange OWA\Proxy User Requests"
"\MSExchange OWA\Requests/sec"
"\MSExchange RpcClientAccess Per Server(*)\RPC Active Backend Connections (% of Limit)"
"\MSExchange RpcClientAccess Per Server(*)\RPC Average Latency (Backend)"
"\MSExchange RpcClientAccess Per Server(*)\RPC Average Latency (End To End) - Cached Mode"
"\MSExchange RpcClientAccess Per Server(*)\RPC Average Latency (End To End) - Online Mode"
"\MSExchange RpcClientAccess Per Server(*)\RPC Average Latency (End To End)"
"\MSExchange RpcClientAccess Per Server(*)\RPC Failed Backend Connections"
"\MSExchange RpcClientAccess\Active User Count"
"\MSExchange RpcClientAccess\Client: RPCs Failed"
"\MSExchange RpcClientAccess\Client: Latency > 10 sec RPCs"
"\MSExchange RpcClientAccess\Client: Latency > 2 sec RPCs"
"\MSExchange RpcClientAccess\Client: Latency > 5 sec RPCs"
"\MSExchange RpcClientAccess\Connection Count"
"\MSExchange RpcClientAccess\RPC Averaged Latency"
"\MSExchange RpcClientAccess\RPC Clients Bytes Read"
"\MSExchange RpcClientAccess\RPC Clients Bytes Written"
"\MSExchange RpcClientAccess\RPC Operations/sec"
"\MSExchange RpcClientAccess\RPC Packets/sec"
"\MSExchange RpcClientAccess\RPC Requests"
"\MSExchange RpcClientAccess\User Count"
"\MSExchange Sharing Engine\Average Folder Synchronization Time (in seconds)"
"\MSExchange Sharing Engine\Average Time to Request a Token for an External Authentication"
"\MSExchange Throttling Service Client(*)\Average request processing time."
"\MSExchange Throttling(*)\OverBudgetThreshold"
"\MSExchange Throttling(*)\Unique Budgets OverBudget"
"\MSExchange Throttling(*)\Users X Times OverBudget"
"\MSExchangeAB\NSPI Connections Current"
"\MSExchangeAB\NSPI Connections/sec"
"\MSExchangeAB\NSPI RPC Browse Requests Average Latency"
"\MSExchangeAB\NSPI RPC Requests"
"\MSExchangeAB\NSPI RPC Requests Average Latency"
"\MSExchangeAB\NSPI RPC Requests/sec"
"\MSExchangeAB\Referral RPC Requests Average Latency"
"\MSExchangeAB\Referral RPC Requests"
"\MSExchangeAB\Referral RPC Requests/sec"
"\MSExchangeAutodiscover\Requests/sec"
"\MSExchangeFDS:OAB(_total)\Download Task Queued"
"\MSExchangeFDS:OAB(_total)\Download Tasks Completed"
"\MSExchangeImap4(_total)\Active SSL Connections"
"\MSExchangeImap4(_total)\Average Command Processing Time (milliseconds)"
"\MSExchangeImap4(_total)\Connections Rate"
"\MSExchangeImap4(_total)\Current Connections"
"\MSExchangeImap4(_total)\Proxy Current Connections"
"\MSExchangeImap4(_total)\SearchFolder Creation Rate"
"\MSExchangePOP3(_total)\Active SSL Connections"
"\MSExchangePop3(_total)\Average Command Processing Time (milliseconds)"
"\MSExchangePop3(_total)\Connections Current"
"\MSExchangePop3(_total)\Connections Rate"
"\MSExchangePop3(_total)\DELE Rate"
"\MSExchangePop3(_total)\Proxy Current Connections"
"\MSExchangePop3(_total)\RETR Rate"
"\MSExchangePop3(_total)\UIDL Rate"
"\MSExchangeWS\Average Response Time"
"\MSExchangeWS\Items Read/sec"
"\MSExchangeWS\Proxy average response time"
"\MSExchangeWS\Requests/sec"
"\MSExchangeWS\Request rejections/sec"
"\W3SVC_W3WP(*)\*"  
"\WAS_W3WP(*)\*"
"\Web Service(_Total)\Bytes Received/sec"
"\Web Service(_Total)\Bytes Sent/sec"
"\Web Service(_Total)\Bytes Total/sec"
"\Web Service(_Total)\Connection Attempts/sec"
"\Web Service(_Total)\Current Connections"
"\Web Service(_Total)\ISAPI Extension Requests/sec"
"\Web Service(_Total)\Other Request Methods/sec"
)		
	Write-Debug "Added Exchange 2010 CAS Counters"
	}
		$Counters += $CASCounterList
	}
	if ($GetServer.IsHubTransportServer -eq $true){
		$script:roles += [string]"Hub"
		# HUB Counter list
		if ($Exchange2007){
		$HUBCounterList = @(
"\MSExchangeEdgeSync Job(*)\Edge objects added/sec"
"\MSExchangeEdgeSync Job\Edge objects deleted/sec"
"\MSExchangeEdgeSync Job(*)\Edge objects updated/sec"
"\MSExchangeEdgeSync Topology\Jobs waiting total"
"\MSExchangeEdgeSync Topology\SyncNow Edges not completed total"
"\MSExchangeEdgeSync Job\Scan jobs completed successfully total"
"\MSExchangeEdgeSync Job\Scan jobs failed because could not extend lock total"
"\MSExchangeEdgeSync Job\Scan jobs failed because of directory error total"
"\MSExchangeEdgeSync Job\Source objects scanned/sec"
"\MSExchangeEdgeSync Job\Target objects scanned/sec"
"\MSExchangeTransport Batch Point(*)\Batches waiting current"
"\MSExchangeTransport Dumpster\Dumpster Inserts/sec"
"\MSExchangeTransport Dumpster\Dumpster Item Count"
"\MSExchangeTransport Dumpster\Dumpster Size"
"\MSExchangeTransport Dumpster\Dumpster Deletes: Quota"
"\MSExchangeTransport Dumpster\Dumpster Mailbox Database Count"
"\MSExchangeTransport Dumpster\Dumpster Deletes/sec"
"\MSExchangeTransport Queues(_total)\Active Mailbox Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Active Remote Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Aggregate Delivery Queue Length (All Queues)"
"\MSExchangeTransport Queues(_total)\Items Completed Delivery Per Second"
"\MSExchangeTransport Queues(_total)\Items Completed Delivery Total"
"\MSExchangeTransport Queues(_total)\Items Queued for Delivery Per Second"
"\MSExchangeTransport Queues(_total)\Largest Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Messages Completed Delivery Per Second"
"\MSExchangeTransport Queues(_total)\Messages Completed Delivery Total"
"\MSExchangeTransport Queues(_total)\Messages Queued for Delivery Per Second"
"\MSExchangeTransport Queues(_total)\Poison Queue Length"
"\MSExchangeTransport Queues(_total)\Retry Remote Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Submission Queue Length"
"\MSExchangeTransport Queues(_total)\Unreachable Queue Length"
"\MSExchangeTransport SmtpReceive(_total)\Average bytes/message"
"\MSExchangeTransport SmtpReceive(_total)\Disconnections by Agents/second"
"\MSExchangeTransport SmtpReceive(_total)\Message Bytes Received/sec"
"\MSExchangeTransport SmtpReceive(_total)\Messages Received Total"
"\MSExchangeTransport SmtpReceive(_total)\Messages Received/sec"
"\MSExchangeTransport SmtpReceive(_total)\Tarpitting Delays Anonymous"
"\MSExchangeTransport SmtpSend(_total)\Average message bytes/message"
"\MSExchangeTransport SmtpSend(_total)\Average recipients/message"
"\MSExchangeTransport SmtpSend(_total)\Connections Current"
"\MSExchangeTransport SmtpSend(_total)\Message Bytes Sent/sec"
"\MSExchangeTransport SmtpSend(_total)\Messages Sent Total"
"\MSExchangeTransport SmtpSend(_total)\Messages Sent/sec"
"\MSExchangeTransport Resolver(_total)\Messages Chipped"
"\MSExchangeTransport Resolver(_total)\Messages Created"
"\MSExchange Connection Filtering Agent\Connections on IP Block List Providers /sec"
"\MSExchange Content Filter Agent\Messages Scanned Per Second"
"\MSExchange Database(edgetransport)\Database Cache % Available"
"\MSExchange Database(edgetransport)\Database Cache % Clean"
"\MSExchange Database(edgetransport)\Database Cache % Hit"
"\MSExchange Database(edgetransport)\Database Cache % Versioned"
"\MSExchange Database(edgetransport)\Database Cache Size (MB)"
"\MSExchange Database(edgetransport)\Database Cache Size Max"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Reads Average Latency"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Writes Average Latency"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Log Writes/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Log Reads/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Checkpoint Depth"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Generation Checkpoint Depth"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Generation Checkpoint Depth Max"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Version buckets allocated"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Reads/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Writes/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Record Stalls/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Threads Waiting"
"\MSExchange Extensibility Agents(*)\Average Agent Processing Time (sec)"
"\MSExchange Extensibility Agents(*)\Total Agent Invocations"
"\MSExchange Journaling Agent\Journal Reports Created/sec"
"\MSExchange Journaling Agent\Journaling Processing Time per Message"
"\MSExchange Journaling Agent\Users Journaled/sec"
"\MSExchange Recipient Filter Agent\Recipients Rejected by Recipient Validation/sec"
"\MSExchange Secure Mail Transport(_total)\Domain Secure Messages Sent"
"\MSExchange Sender Id Agent\Messages That Bypassed Validation/sec"
"\MSExchange Store Driver(_total)\Inbound: LocalDeliveryCallsPerSecond"
"\MSExchange Store Driver(_total)\Inbound: MessageDeliveryAttemptsPerSecond"
"\MSExchange Store Driver(_total)\Inbound: Recipients Delivered Per Second"
"\MSExchange Store Driver(_total)\Outbound: Submitted Mail Items Per Second"
"\MSExchange Topology(*)\Latest Exchange Topology Discovery Time in Seconds"
"\MSExchange Transport Rules(*)\Messages Evaluated/sec"
"\MSExchange Transport Rules(*)\Messages Processed/sec"
)
		Write-Debug "Added Exchange 2007 HUB Counters"
		}
	if ($Exchange2010){
		$HUBCounterList = @(
"\MSExchange Connection Filtering Agent\Connections on IP Block List Providers /sec"
"\MSExchange Content Filter Agent\Messages Scanned Per Second"
"\MSExchange Conversations Transport Agent\Average message processing time"

"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Reads Average Latency"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Writes Average Latency"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Log Writes/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Log Reads/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Checkpoint Depth"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Generation Checkpoint Depth"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Generation Checkpoint Depth Max"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Version buckets allocated"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Reads/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\I/O Database Writes/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Record Stalls/sec"
"\MSExchange Database ==> Instances(edgetransport/Transport Mail Database)\Log Threads Waiting"
"\MSExchange Extensibility Agents(*)\Average Agent Processing Time (sec)"
"\MSExchange Extensibility Agents(*)\Total Agent Invocations"
"\MSExchange Journaling Agent\Journal Reports Created/sec"
"\MSExchange Journaling Agent\Journaling Processing Time per Message"
"\MSExchange Journaling Agent\Journaling Processing Time"
"\MSExchange Journaling Agent\Users Journaled/sec"
"\MSExchange Log Search Service\Average search processing time"
"\MSExchange Message Tracking\Average Get-MessageTrackingReport Processing Time"
"\MSExchange Message Tracking\Average Search-MessageTrackingReport Processing Time"
"\MSExchange Message Tracking\Get-MessageTrackingReport Processing Time"
"\MSExchange Message Tracking\Search-MessageTrackingReport Processing Time"
"\MSExchange Recipient Filter Agent\Recipients Rejected by Recipient Validation/sec"
"\MSExchange Store Driver(_total)\Inbound: MessageDeliveryAttemptsPerSecond"
"\MSExchange Store Driver(_total)\Inbound: LocalDeliveryCallsPerSecond"
"\MSExchange Store Driver(_total)\Inbound: Recipients Delivered Per Second"
"\MSExchange Store Driver(_total)\Outbound: Submitted Mail Items Per Second"
"\MSExchange Text Messaging\Average text message delivery latency (milliseconds)"
"\MSExchange Throttling Service Client(*)\Percentage of Denied Submission Request."
"\MSExchange Transport Rules(*)\Messages Evaluated/sec"
"\MSExchange Transport Rules(*)\Messages Processed/sec"
"\MSExchangeTransport Component Latency(*)\Percentile99"
"\MSExchangeTransport DeliveryAgent(All Instances)\Average Bytes Per Message"
"\MSExchangeTransport DeliveryAgent(All Instances)\Average Messages Per Connection"
"\MSExchangeTransport DeliveryAgent(All Instances)\Connections Completed Per Second"
"\MSExchangeTransport DeliveryAgent(All Instances)\Connections Failed Per Second"
"\MSExchangeTransport DeliveryAgent(All Instances)\Message Bytes Sent Per Second"
"\MSExchangeTransport DeliveryAgent(All Instances)\Messages Delivered Per Second"
"\MSExchangeTransport Delivery Failures\*"
"\MSExchangeTransport DSN(_total)\Delay DSNs"
"\MSExchangeTransport Dumpster\Dumpster Deletes: Quota"
"\MSExchangeTransport Dumpster\Dumpster Deletes/sec"
"\MSExchangeTransport Dumpster\Dumpster Inserts/sec"
"\MSExchangeTransport Dumpster\Dumpster Item Count"
"\MSExchangeTransport Dumpster\Dumpster Mailbox Database Count"
"\MSExchangeTransport Dumpster\Dumpster Resubmit Jobs: Average Execution Time (sec)"
"\MSExchangeTransport Dumpster\Dumpster Resubmit Jobs: Average Request Latency (sec)"
"\MSExchangeTransport Dumpster\Dumpster Size"
"\MSExchangeTransport IsMemberOfResolver(Transport)\IsMemberOfResolver ResolvedGroups Cache Size Percentage"
"\MSExchangeTransport Queues(_total)\Active Mailbox Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Active Non-Smtp Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Active Remote Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Aggregate Delivery Queue Length (All Queues)"
"\MSExchangeTransport Queues(_total)\Largest Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Messages Completed Delivery Per Second"
"\MSExchangeTransport Queues(_total)\Messages Queued for Delivery Per Second"
"\MSExchangeTransport Queues(_total)\Messages Submitted Per Second"
"\MSExchangeTransport Queues(_total)\Poison Queue Length"
"\MSExchangeTransport Queues(_total)\Retry Mailbox Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Retry Non-Smtp Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Retry Remote Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Submission Queue Length"
"\MSExchangeTransport Queues(_total)\Unreachable Queue Length"
"\MSExchangeTransport Resolver(_total)\Messages Chipped"
"\MSExchangeTransport Resolver(_total)\Messages Created"
"\MSExchangeTransport SMTPAvailability(*)\% Activity"
"\MSExchangeTransport SMTPAvailability(*)\% Availability"
"\MSExchangeTransport SMTPAvailability(*)\% Failures Due To Active Directory Down"
"\MSExchangeTransport SMTPAvailability(*)\% Failures Due To Back Pressure"
"\MSExchangeTransport SMTPAvailability(*)\% Failures Due To IO Exceptions"
"\MSExchangeTransport SMTPAvailability(*)\% Failures Due To MaxInboundConnectionLimit"
"\MSExchangeTransport SmtpReceive(_total)\Average bytes/message"
"\MSExchangeTransport SmtpReceive(_total)\Disconnections by Agents/second"
"\MSExchangeTransport SmtpReceive(_total)\Message Bytes Received/sec"
"\MSExchangeTransport SmtpReceive(_total)\Messages Received/sec"
"\MSExchangeTransport SmtpSend(_total)\Average message bytes/message"
"\MSExchangeTransport SmtpSend(_total)\Average recipients/message"
"\MSExchangeTransport SmtpSend(_total)\Connections Current"
"\MSExchangeTransport SmtpSend(_total)\Message Bytes Sent/sec"
"\MSExchangeTransport SmtpSend(_total)\Messages Sent Total"
"\MSExchangeTransport SmtpSend(_total)\Messages Sent/sec"
)
	Write-Debug "Added Exchange 2010 HUB Counters"
	}
		$Counters += $HUBCounterList
	}
	if ($GetServer.IsEdgeServer -eq $true){
	$script:roles += [string]"Edge"
	#Edge Counter List
	if ($Exchange2007){
	$EdgeCounterList = @(
"\AD/AM(ADAM_MSExchange)\LDAP Searches/sec"
"\AD/AM(ADAM_MSExchange)\LDAP Writes/sec"
"\MSExchange Attachment Filtering\Messages Attachment Filtered"
"\MSExchange Attachment Filtering\Messages Filtered/sec"
"\MSExchange Content Filter Agent\Messages Deleted"
"\MSExchange Content Filter Agent\Messages Quarantined"
"\MSExchange Content Filter Agent\Messages Rejected"
"\MSExchange Content Filter Agent\Messages Scanned Per Second"
"\MSExchange Content Filter Agent\Messages that Bypassed Scanning"
"\MSExchange Content Filter Agent\Messages with SCL 0"
"\MSExchange Content Filter Agent\Messages with SCL 1"
"\MSExchange Content Filter Agent\Messages with SCL 2"
"\MSExchange Content Filter Agent\Messages with SCL 3"
"\MSExchange Content Filter Agent\Messages with SCL 4"
"\MSExchange Content Filter Agent\Messages with SCL 5"
"\MSExchange Content Filter Agent\Messages with SCL 6"
"\MSExchange Content Filter Agent\Messages with SCL 7"
"\MSExchange Content Filter Agent\Messages with SCL 8"
"\MSExchange Content Filter Agent\Messages with SCL 9"
"\MSExchange Database ==> Instances(*)\I/O Database Reads/sec"
"\MSExchange Database ==> Instances(*)\I/O Database Writes/sec"
"\MSExchange Database ==> Instances(*)\I/O Log Reads/sec"
"\MSExchange Database ==> Instances(*)\I/O Log Writes/sec"
"\MSExchange Database ==> Instances(*)\Log Generation Checkpoint Depth"
"\MSExchange Database ==> Instances(*)\Log Record Stalls/sec"
"\MSExchange Database ==> Instances(*)\Log Threads Waiting"
"\MSExchange Database ==> Instances(*)\Version buckets allocated"
"\MSExchange Database(edgetransport)\Database Cache Size (MB)"
"\MSExchange Extensibility Agents(*)\Average Agent Processing Time (sec)"
"\MSExchange Protocol Analysis Background Agent\Block Senders"
"\MSExchange Recipient Filter Agent\Recipients Rejected by Block List/sec"
"\MSExchange Recipient Filter Agent\Recipients Rejected by Recipient Validation/sec"
"\MSExchange Sender Filter Agent\Messages Filtered by Sender Filter/sec"
"\MSExchange Sender Id Agent\DNS queries/sec"
"\MSExchange Transport Rules(*)\Message Processed/sec"
"\MSExchange Transport Rules(*)\Messages Evaluated/sec"
"\MSExchangeTransport Queues(_total)\Active Remote Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Aggregate Delivery Queue Length (All Queues)"
"\MSExchangeTransport Queues(_total)\Largest Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Poison Queue Length"
"\MSExchangeTransport Queues(_total)\Retry Remote Delivery Queue Length"
"\MSExchangeTransport Queues(_total)\Submission Queue Length"
"\MSExchangeTransport Queues(_total)\Unreachable Queue Length"
)
	Write-Debug "Added Exchange 2007 Edge Counters"
	}
	$Counters += $EdgeCounterList
	}
	if ($GetServer.IsUnifiedMessagingServer -eq $true){
		$script:roles += [string]"Um"
		#UM Counter List
		if ($Exchange2007){
		$UMCounterList = @(
"\MSExchangeUMAvailability\Unhandled Exceptions per Second"
"\MSExchangeUMGeneral\Total Calls per Second"
"\MSExchangeUMGeneral\User Response Latency"
"\MSExchangeUMGeneral\Current Calls"
"\ASP.NET Apps v..(_LM_WSVC__Root_UnifiedMessaging)\Request Cutting"
"\ASP.NET Apps v..(_LM_WSVC__ROOT_UnifiedMessaging)\Requests Queued"
"\ASP.NET Apps v..(_LM_WSVC__ROOT_UnifiedMessaging)\Request Wait Time"
)
		Write-Debug "Added Exchange 2007 UM Counters"
		}
		if ($Exchange2010){
		$UMCounterList = @(
"\MSExchangeUMCallAnswer\Fetch Greeting Timed Out"
"\MSExchangeUMGeneral\% Successful Caller ID Resolutions"
"\MSExchangeUMGeneral\Current Calls"
"\MSExchangeUMGeneral\Current Voice Mail Calls"
"\MSExchangeUMGeneral\User Response Latency"
"\MSExchangeUMAvailability\% of Failed Mailbox Connection Attempts Over the Last Hour"
"\MSExchangeUMAvailability\% of Inbound Calls Rejected by UM Service Over the Last Hour"
"\MSExchangeUMAvailability\% of Inbound Calls Rejected by UM Worker Process over the Last Hour"
"\MSExchangeUMAvailability\% of Messages Successfully Processed Over the Last Hour"
"\MSExchangeUMAvailability\% of Partner Voice Message Transcription Failures Over the Last Hour"
"\MSExchangeUMAvailability\Call Answer Queued Messages"
"\MSExchangeUMAvailability\Direct Access Failures"
"\MSExchangeUMAvailability\Queued OCS User Event Notifications"
"\MSExchangeUMAvailability\Total Queued Messages"
"\MSExchangeUMAvailability\Unhandled Exceptions/sec"
"\MSExchangeUMCallAnswer\Calls Disconnected by Callers During UM Audio Hourglass"
"\MSExchangeUMPerformance\Operations over Six Seconds"
"\MSExchangeUMSubscriberAccess\Calls Disconnected by Callers During UM Audio Hourglass"
"\MSExchangeUMVoiceMailSpeechRecognition(en-us)\Average Confidence %"
"\MSExchangeUMVoiceMailSpeechRecognition(en-us)\Voice Messages Not Processed Because of Low Availability of Resource"
)
	Write-Debug "Added Exchange 2010 UM Counters"
	}
		$Counters += $UMCounterList
	}
} 

#Add custom counters if -CustomCounterPath is specified
if ($CustomCounterPath.Length -ne 0)
{
	if (!(test-path $CustomCounterPath))
	{	
		Write-Host "ERROR: Custom Counter File Path not found. Continuing without adding custom counters" -ForegroundColor Red
		Write-Host ""
	}
	else
	{
		Write-Host "Reading Custom Counter File..." -NoNewline
		$CustomCounters = Get-Content $CustomCounterPath
		$Counters += $CustomCounters
		Write-Host " COMPLETED"
		Write-Host ""
	}
}

# Remove duplicate counters if any....
	Write-Debug "Removing duplicate Counters"
	$script:CounterList = $Counters | Sort-Object | Select-Object -Unique
}
function WriteCounterConfig
{
	#Write list of performance counters to .config file for counter log creation.
	Write-Debug "Writing Counter Config file to disk"
	Out-File -FilePath ".\Exchange_Perfwiz.Config" -InputObject $CounterList -Force -Encoding "ascii"
}

function CheckifCollectionExists
{
#	Check if Existing Exchange_Perfwiz Data Collection exists
	Write-Debug "Checking if Existing Data Collector Exists"
	$QueryCollection = "logman query Exchange_Perfwiz -s $Servername"
	$CheckifExists = Invoke-Expression -Command $QueryCollection
		if ($Windows2003){$SearchString = "does not exist"}
		elseif ($Windows2008 -or $Windows2008R2){$SearchString = "Set was not found"}
		else{Write-Host "Incorrect Server version detected"}
	$cmd = Select-String -InputObject $CheckifExists -Pattern $SearchString -quiet
	if ($cmd -ne $true -and $quiet){
		Write-Host "Previous Exchange_Perfwiz collector found..." 
		Write-Host "Delete the existing Exchange_Perfwiz Data Collector? "
		Write-Host "Running quiet, assuming the removal of data collector "
		if ($quiet){StopAndDeleteCounter; return}

	}
	elseif ($cmd -ne $true){
		Write-Host "Previous Exchange_Perfwiz collector found..." 
		Write-Host "Delete the existing Exchange_Perfwiz Data Collector? " -NoNewline
			$answer = ConfirmAnswer
			if ($answer -eq "yes"){StopAndDeleteCounter; return}
			if ($answer -eq "no")
			{
				Write-Host "Start the existing Exchange_Perfwiz Data Collector? " -NoNewline
				$answer = ConfirmAnswer
				if ($answer -eq "yes")
				{
					Write-Host "Starting existing Exchange_Perfwiz Data Collector... " -NoNewline
					$QueryCollection = "logman query Exchange_Perfwiz -s $Server"
					$CheckifRunning = Invoke-Expression -Command $QueryCollection
					[string]$CheckStatus = $CheckifRunning -match "Status:"
					$RunningStatus = $CheckStatus.Contains("Running")
					if ($RunningStatus){
						Write-Host ""
						Write-Host "Exchange_Perfwiz Data collector already running..." -ForegroundColor Yellow
						Write-Host ""
						Exit
						}
					else{				
					$commandString = "logman start -n Exchange_Perfwiz -s $Server"
					$StartCounter = Invoke-Expression -Command $commandString 
					Write-Host "COMPLETED"
					Write-Host ""
					Exit}
				}	
				elseif ($answer -eq "no")
				{	
					if ($Exmon)
					{
						Write-Host ""
						Execute_Exmon
						Exit
					}
					else
					{
						Write-Host ""
						Exit
					}
				}
			}
			else
			{
				$answer = ConfirmAnswer
				if($answer -eq "no"){Write-Host ""; Exit}
			}
	}
	else{
		Write-Host "Existing Exchange_Perfwiz Data Collection not found. Creating New..." 
	}
}

function DeleteCounterConfig
{
	Write-Debug "Deleting Counter Config File"
	$Exists = Test-Path ".\Exchange_Perfwiz.Config"
	if ($Exists){Remove-Item ".\Exchange_Perfwiz.Config"}
}

function CreateCounter()
{
	#Create Counter Data Collection depending on role and switches passed
	Write-Debug "Create Data Collector"
	Write-Host "Creating Exchange_Perfwiz Data Collector.............. " -NoNewline
	#Set static Default sample interval (-si) and duration (-rf) if not specified.
	if (($interval -eq "")) {$interval = 30} else {$interval = $interval}
	if ($duration -eq "") {$duration = "08:00:00"} else {$duration = $duration}
	if ($filepath -eq "") {$filepath = "C:\Perflogs\"} elseif ($Filepath.EndsWith("\")) {$filepath = $filepath} else {$filepath = $filepath + "\"}
	foreach ($role in $roles){$rolenames += $role}
	
	if ($Windows2003)
	{
		if ($circular)
		{
			$commandString = "logman create counter -n Exchange_Perfwiz -cf Exchange_Perfwiz.Config -s $ServerName -f bincirc -max $maxsize -si " + $interval + " -o " + $filepath + $ServerName + "_" + $rolenames + "_Circular"
		}
		Else{
			#Windows 2003 (Removed duration since log roll fails to work with duration specified)
			$commandString = "logman create counter -n Exchange_Perfwiz -cf Exchange_Perfwiz.Config -s $ServerName -f bin -cnf -max $maxsize -si " + $interval + " -o " + $filepath + $ServerName + "_" + $rolenames + " -v MMDDHHMM"
		}
	}
	if ($Windows2008 -or $Windows2008R2)
	{
		if ($circular)
		{
			$commandString = "logman create counter -n Exchange_Perfwiz -cf Exchange_Perfwiz.Config -s $ServerName -f bincirc -max $maxsize -cnf -si " + $interval + " -o " + $filepath + $ServerName + "_" + $rolenames + "_Circular"
		}
		If ($Windows2008R2)
		{
			#Win2k8 R2 installed
			$commandString = "logman create counter -n Exchange_Perfwiz -cf Exchange_Perfwiz.Config -s $ServerName -f bin -cnf 0 -max $maxsize -si " + $interval + " -rf " + $duration + " -o " + $filepath + $ServerName + "_" + $rolenames
		}
		Elseif ($Windows2008)
		{
			#Windows 2008 (-max cannot be used due to OS bug for log rolling)
			if ($interval -lt 30) {$StopLimitDuration = "01:00:00"} else {$StopLimitDuration = "04:00:00"}
			$commandString = "logman create counter -n Exchange_Perfwiz -cf Exchange_Perfwiz.Config -s $ServerName -f bin -cnf $StopLimitDuration -si " + $interval + " -rf " + $duration + " -v MMDDHHMM -o " + $filepath + $ServerName + "_" + $rolenames
			if ($maxsize -ne 512)
			{
				$script:MaxSizeSuppressed = $true
				Write-Warning "Maxsize parameter supressed on Windows 2008 machines, using time interval instead"
			}
		}
	}
	
	# Add Begin and End times if passed
	if ($begin)
	{
		$addbegintime = " -b " + $begin
		$commandString += $addbegintime
	}
	if ($end)
	{
		$addendtime = " -e " + $end
		$commandString += $addendtime
	}
			
	#Invoke Command
	$CreateCounter = Invoke-Expression -command $commandString
	
	# Check to see if Invoke completed successfully
	if ($Createcounter -notmatch "The command completed successfully")
	{
		if($Createcounter -match "Access is denied")
		{
			Write-Host "ERROR" -foregroundcolor red
			Write-Host "Access is denied. Open the Exchange Management Shell using Run as Administrator" -foregroundcolor red
			exit
		}
		else
		{
			Write-Host "ERROR" -foregroundcolor red
			Write-Host $Createcounter -foregroundcolor red
			exit
		}
	}
	
	if ($MaxSizeSuppressed)
	{
		Write-Host "COMPLETED"
		Write-Warning "Maxsize parameter supressed on Windows 2008 machines, using time interval instead"
		Write-Host ""
	}
	else 
	{
		Write-Host "COMPLETED"
	}
	Write-Debug $commandString
	
	#Create Screen Output array
	$WriteOutput = @("")

	#Add Maxsize (Omit for Windows 2008 Non R2 versions
	if ($Windows2003 -or $Windows2008R2)
	{
		$AddOutput = "Maxsize: $maxsize MB"
		$WriteOutput += $AddOutput
	}

	#Add Interval
	$AddOutput ="Interval  (seconds): $interval"
	$WriteOutput += $AddOutput

	#Add Duration
	if ($Windows2003 -or $Windows2008 -or $Windows2008R2)
	{
		if ($circular)
		{
			$AddOutput = "Duration (hh:mm:ss): Circular logging enabled"
			$WriteOutput += $AddOutput
		}
		elseif($Windows2003)
		{
			$AddOutput = "Duration (hh:mm:ss): N/A"
			$WriteOutput += $AddOutput
		}
		else
		{
			$AddOutput = "Duration (hh:mm:ss): $duration"
			$WriteOutput += $AddOutput
		}
	}

	#Add Log Roll Duration for Windows 2008 Non R2 servers
	if($Windows2008)
		{
			$AddOutput = "Log Roll Duration (hh:mm:ss): $StopLimitDuration"
			$WriteOutput += $AddOutput
		}
	
	#Add Role and Data Location
	$AddOutput = @(
	"Counters for Role(s): $roles"
	"Data Location: $filepath"
	)
	$WriteOutput += $AddOutput

	#Add Extended counter text info
	if ((Get-ExchangeServer -Identity $Servername | where {$_.IsMailboxServer -eq $true}))
	{
		$StoreExtendedOnText = "Store Extended Counters: On"
		$StoreExtendedOffText = "Store Extended Counters: Off"
		$ESEExtendedOnText = "ESE Extended Counters: On"
		$ESEExtendedOffText = "ESE Extended Counters: Off"
	

		#Add Extended counter config to array
		if ($StoreExtendedOn){$WriteOutput += $StoreExtendedOnText}
		else {$WriteOutput += $StoreExtendedOffText}
		if ($ESEExtendedOn){$WriteOutput += $ESEExtendedOnText}
		else {$WriteOutput += $ESEExtendedOffText}
	}
		#Write Config info to screen
		foreach ($item in $WriteOutput){Write-Host $item}
		Write-Host ""

	#Cleanup
	if ($Windows2003 -and !$circular)
	{
		#DeleteCounterConfig
		PromptStartCollection
	}
	elseif ($begin -or $end)
	{
		DeleteCounterConfig
	}
	else
	{
		#DeleteCounterConfig
		PromptStartCollection
	}
	if ($exmon) {Execute_Exmon;}
}

function PromptStartCollection
{
	# Ask to start ExPerfwiz logging
	if ($quiet)
	{
		Write-Host "Starting Data Collector..." -NoNewline
		Start-Sleep 2
		$cmd = "logman start Exchange_Perfwiz -s $Servername"
		$Invokecmd = Invoke-Expression $cmd 
		$SearchString = "Cannot create a file when that file already exists"
		$CheckCmd = Select-String -InputObject $Invokecmd -Pattern $SearchString -quiet
		while(($CheckCmd = Select-String -InputObject $Invokecmd -Pattern $SearchString -quiet) -eq $true) 
		{
			Write-Host "." -NoNewline
			$Invokecmd = Invoke-Expression $cmd
			Start-Sleep 2
		}
		if ($Invokecmd -match "The command completed successfully")
		{
			Write-Host "COMPLETED"
		}
		else
		{
			Write-Host " FAILED" -ForegroundColor Red
			Write-Debug "$Invokecmd"
			Write-Host "Check the application event log for any errors" -ForegroundColor Red
		}

	}
	else
	{
		Write-Debug "Prompt to Start Collection"
		Write-Host "Start the Exchange_Perfwiz Data Collection now? " -NoNewline
		$answer = ConfirmAnswer
		if ($answer -eq "yes")
		{
			Write-Host "Starting Data Collector..." -NoNewline
			Start-Sleep 2
			$cmd = "logman start Exchange_Perfwiz -s $Servername"
			$Invokecmd = Invoke-Expression $cmd 
			$SearchString = "Cannot create a file when that file already exists"
			$CheckCmd = Select-String -InputObject $Invokecmd -Pattern $SearchString -quiet
			while(($CheckCmd = Select-String -InputObject $Invokecmd -Pattern $SearchString -quiet) -eq $true) 
			{
				Write-Host "." -NoNewline
				$Invokecmd = Invoke-Expression $cmd
				Start-Sleep 2
			}
			if ($Invokecmd -match "The command completed successfully")
			{
				Write-Host "COMPLETED"
			}
			else
			{
				Write-Host " FAILED" -ForegroundColor Red
				Write-Debug "$Invokecmd"
				Write-Host "Check the application event log for any errors" -ForegroundColor Red
			}
		}
		elseif($answer -eq "no"){return}
		Write-Host ""
	}
}

Function StopCollection
{
	Write-Debug "Stop Data Collection"
	Write-Host ""
	Write-Host "Stopping Exchange_Perfwiz Data Collector if running... " -NoNewline
	$commandString = "logman stop -n Exchange_Perfwiz -s $Servername"
	$Error.Clear()
	$StopCounter = Invoke-Expression -Command $commandString -ErrorAction SilentlyContinue
	if ($Error){Write-host "Error encountered"; exit}
	else {Write-Host "COMPLETED"; Write-Host ""}
	$CheckExmon = @(logman query -s $ServerName) -match "Exmon_Trace"
	$CheckifRunning = select-string -InputObject $CheckExmon -pattern "Running" -quiet
	if ($CheckifRunning)
	{
		Write-Host "Stopping Exmon Tracing... " -NoNewline
		$cmd = "logman stop -n Exmon_Trace -s $Servername"
		$StopExmon = Invoke-Expression -Command $cmd
		Write-Host "COMPLETED"
		Write-Host ""
	}
}

Function StopAndDeleteCounter
{
	Write-Debug "Stop and Delete Data Collector"
	Write-Host ""
	Write-Host "Stopping Exchange_Perfwiz Data Collector if running... " -NoNewline
	$commandString = "logman stop -n Exchange_Perfwiz -s $Servername"
	$StopCounter = Invoke-Expression -Command $commandString 
	Write-Host "COMPLETED"
	Start-Sleep -Seconds 5
	Write-Host "Deleting Exchange_Perfwiz Data Collector.............. " -NoNewline
	$commandString = "logman delete -n Exchange_Perfwiz -s $Servername"
	$DeleteCounter = Invoke-Expression -Command $commandString
	if ($DeleteCounter -notmatch "The command completed successfully")
	{
		if($DeleteCounter -match "Access is denied")
		{
			Write-Host "ERROR" -foregroundcolor red
			Write-Host "Access is denied. Open the Exchange Management Shell using Run as Administrator" -foregroundcolor red
			exit
		}
		else
		{
			Write-Host "ERROR" -foregroundcolor red
			Write-Host $DeleteCounter -foregroundcolor red
			exit
		}
	}
	Write-Host "COMPLETED"
	Start-Sleep -Seconds 2
}

Function DeleteCollection
{
	Write-Debug "Deleting Data Collector"
	Write-Host ""
	$QueryCollection = "logman query Exchange_Perfwiz -s $Servername"
	$CheckifExists = Invoke-Expression -Command $QueryCollection
		if ($Windows2003){$SearchString = "does not exist"}
		elseif ($Windows2008 -or $Windows2008R2){$SearchString = "Set was not found"}
	$cmd = Select-String -InputObject $CheckifExists -Pattern $SearchString -quiet
	if ($cmd -eq $true){
		Write-Host "Exchange_Perfwiz Data Collector not found"
		Write-Host ""
		Exit
	}
	Write-Host "Stopping Exchange_Perfwiz Data Collector if running... " -NoNewline
	$commandString = "logman stop -n Exchange_Perfwiz -s $Servername"
	$StopCounter = Invoke-Expression -Command $commandString 
	Write-Host "COMPLETED"
	Start-Sleep -Seconds 5
	Write-Host "Deleting Exchange_Perfwiz Data Collector.............. " -NoNewline
	$commandString = "logman delete -n Exchange_Perfwiz -s $Servername"
	$DeleteCounter = Invoke-Expression -Command $commandString
	if ($DeleteCounter -notmatch "The command completed successfully")
	{
		if($DeleteCounter -match "Access is denied")
		{
			Write-Host "ERROR" -foregroundcolor red
			Write-Host "Access is denied. Open the Exchange Management Shell using Run as Administrator" -foregroundcolor red
			exit
		}
		else
		{
			Write-Host "ERROR" -foregroundcolor red
			Write-Host $DeleteCounter -foregroundcolor red
			exit
		}
	}
	Write-Host "COMPLETED"
	Write-Host ""
	$CheckExmon = @(logman query -s $ServerName) -match "Exmon_Trace"
	if ($CheckExmon)
	{
		Write-Host "Stopping Exmon Tracing if running... " -NoNewline
		$cmd = "logman stop -n Exmon_Trace -s $Servername"
		$StopCounter = Invoke-Expression -Command $cmd
		Write-Host "COMPLETED"
		Write-Host "Deleting Exmon Tracing.............. " -NoNewline
		$cmd = "logman delete -n Exmon_Trace -s $Servername"
		$DeleteExmon = Invoke-Expression -Command $cmd
		Write-Host "COMPLETED"
		Write-Host ""
	}
}

Function QueryCollection 
{
	Write-Debug "Query Data Collector"
	Write-Host ""
	Write-Host "Dumping Exchange_Perfwiz Data Collector Information"
	Write-Host "==================================================="
	Invoke-Expression "logman query Exchange_Perfwiz -s $Servername"
	$CheckExmon = @(logman query -s $ServerName) -match "Exmon_Trace"
	if ($CheckExmon)
	{
		Write-Host ""
		Write-Host "Dumping Exmon_Trace Information"
		Write-Host "==============================="
		Invoke-Expression "logman query Exmon_Trace -s $Servername"
	}
}

Function StartCollection 
{
	Write-Debug "Start Data Collector"
	if ($GetOSVerMajor -eq $null){GetOSVersion}
	$QueryPerfCollection = "logman query Exchange_Perfwiz -s $Servername"
	$CheckifExists = Invoke-Expression -Command $QueryPerfCollection
	if ($Windows2003){$SearchString = "does not exist"}
	elseif ($Windows2008 -or $Windows2008R2){$SearchString = "Set was not found"}
	$cmd = Select-String -InputObject $CheckifExists -Pattern $SearchString -quiet
	if ($cmd -eq $true)
	{
		Write-Host ""
		Write-Host "Exchange_Perfwiz Data Collector not found"
		Write-Host ""
		Exit
	}
	#Check if running
	$CheckifRunning = Invoke-Expression -Command $QueryPerfCollection
	[string]$CheckStatus = $CheckifRunning -match "Status:"
	$RunningStatus = $CheckStatus.Contains("Running")
	
	#Reset duration to 8 hours. Windows2003 removes duration when collections are stopped. Applies to circular only
	if ($Windows2003 -and !$RunningStatus)
	{
#		$commandString = "logman update Exchange_Perfwiz -rf 08:00:00 -s $Servername"
#		$UpdateCounter = Invoke-Expression -Command $commandString
#		Write-Host ""
#		Write-Warning "Exchange_Perfwiz duration has been reset to 8 hours. If a different duration is needed,"
#		Write-Warning "please rerun Experfwiz.ps1 specifying the appropriate duration"
#		Write-Host ""
#		PromptStartCollection
		Write-Host ""
		Write-Host "Starting Exchange_Perfwiz Data Collector... " -NoNewline
		$Start = Invoke-Expression "logman start Exchange_Perfwiz -s $Servername"
		Write-Host "COMPLETED"
		Write-Host ""
	}
	elseif($RunningStatus)
	{
		Write-Host ""
		Write-Host "Exchange_Perfwiz Data Collector already running"
		Write-Host ""
	}
#	else
#	{
#		
#	}
	#Exmon Tracing
	#Check if Exmon Tracing exists
	$QueryExmonCollection = "logman query Exmon_Trace -s $Servername"
	$CheckifExists = Invoke-Expression -Command $QueryExmonCollection
	if ($Windows2003){$SearchString = "does not exist"}
	elseif ($Windows2008 -or $Windows2008R2){$SearchString = "Set was not found"}
	$cmd = Select-String -InputObject $CheckifExists -Pattern $SearchString -quiet
	if ($cmd -ne $true)
	{
		# Query Exmon State
		$QueryTrace = "logman query Exmon_Trace -s $ServerName"
		$CheckifRunning = Invoke-Expression -Command $QueryTrace
		[string]$CheckStatus = $CheckifRunning -match "Status:"
		$RunningStatus = $CheckStatus.Contains("Running")
		
		if ($Windows2003 -and !$RunningStatus)
		{
			$commandString = "logman update Exmon_Trace -rf 00:30:00 -s $Servername"
			$UpdateTrace = Invoke-Expression -Command $commandString
			Write-Warning "Exmon_Trace duration has been reset to 30 minutes. If a different duration is needed,"
			Write-Warning "please rerun Experfwiz.ps1 specifying the appropriate duration"
			Write-Host ""
			
			Write-Host "Start the Exmon_Trace now? " -NoNewline
			$answer = ConfirmAnswer
			if ($answer -eq "yes")
			{
				Write-Host "Starting Exmon_Trace..." -NoNewline
				$cmd = "logman start Exmon_Trace -s $Servername"
				$Invokecmd = Invoke-Expression $cmd 
				
				$SearchString = "The command completed successfully"
				$CheckCmd = Select-String -InputObject $Invokecmd -Pattern $SearchString -quiet
				if (!$Checkcmd)
				{
					Write-Host "FAILED"
					Write-Warning "Exmon Tracing failed to start. Check event log for further details"
				}
				else
				{
					Write-Host "COMPLETED"
					Write-Host ""
				}
			}
		}
		elseif($RunningStatus)
		{
			Write-Host "Exmon_Trace is already running"
			Write-Host ""
		}
		else
		{
			Write-Host "Starting Exmon Tracing... "	-NoNewline
			$cmd = "logman start -s $Servername Exmon_Trace"
			$StartExmon = Invoke-Expression -Command $Cmd
			Write-Host "COMPLETED"
			Write-Host ""
		}
		}
}

Function RemoteRegistry
{
	Write-Debug "Entering Remote Registry function"
	$regkey = "SYSTEM\CurrentControlSet\Control\SecurePipeServers\winreg"

	#Try to access remote registry
	&{
		$RegValue = GetValueFromRegistry $Server $regkey
	}
	#catch
	trap [SystemException] 
	{
		if ($_ -match "The network path was not found")
		{
			Write-Host ""
			Write-Host "Remote Host $server not accessible. Check to ensure the Remote Registry service is running and that you have the proper permissions." -ForegroundColor Red
			Write-Host ""
		}
		else
		{
			Write-Host $_.Exception.Message
		}
		exit
	}
}

Function Enable-ExtendedStoreCounters
{
	Write-Debug "Enable Extended Store Counters"
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Servername)         
	$regKey= $reg.OpenSubKey("System\CurrentControlSet\Services\MSExchangeIS\Performance",$true)           
	$regValue = $regkey.GetValue("Library") 
	if ($RegValue.Contains("mdbperf.dll"))
	{
		$regValue = $RegValue.Replace("mdbperf.dll", "mdbperfx.dll")
		$regkey.SetValue("Library",$regValue)
	}
	else
	{
		Write-Host ""
		Write-Warning "Store Extended Counters already enabled"
	}
}

Function Disable-ExtendedStoreCounters
{
	Write-Debug "Disable Extended Store Counters"
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Servername)         
	$regKey= $reg.OpenSubKey("System\CurrentControlSet\Services\MSExchangeIS\Performance",$true)           
	$regValue = $regkey.GetValue("Library") 
	if ($RegValue.Contains("mdbperfx.dll"))
	{
		$regValue = $RegValue.Replace("mdbperfx.dll", "mdbperf.dll")
		$regkey.SetValue("Library",$regValue)
	}
	else
	{
		Write-Host ""
		Write-Warning "Store Extended Counters already disabled"
	}
}

Function Enable-ExtendedESECounters
{
	Write-Debug "Enable Extended ESE Counters"
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Servername)         
	$regKey = $reg.OpenSubKey("System\CurrentControlSet\Services\ESE\Performance",$true)
	$regValue = $regkey.GetValue("Show Advanced Counters") 
	#Check if value exists
	$CheckValue = $regkey.GetValueNames()
	if ($CheckValue -match "Show Advanced Counters")
	{
		#Check if correct type is defined, if not DWORD, delete it
		if ($regkey.GetValuekind("Show Advanced Counters") -ne "DWORD")
		{
			$regkey.DeleteValue("Show Advanced Counters")
			$regkey.SetValue("Show Advanced Counters","1", "DWORD")
		}
		elseif ($RegValue -ne 1)
		{
			$regkey.SetValue("Show Advanced Counters","1", "DWORD")
		}
		else
		{
			Write-Host ""
			Write-Warning "ESE Extended Counters already enabled"
		}
	}
	else
	{
		$regkey.SetValue("Show Advanced Counters","1", "DWORD")
	}
}

Function Disable-ExtendedESECounters
{
	Write-Debug "Disable Extended ESE Counters"
	$reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Servername)         
	$regKey= $reg.OpenSubKey("System\CurrentControlSet\Services\ESE\Performance",$true)           
	$regValue = $regkey.GetValue("Show Advanced Counters") 
	if ($RegValue -ne 0)
	{
		$regkey.SetValue("Show Advanced Counters","0")
	}
	else
	{
		Write-Host ""
		Write-Warning "ESE Extended Counters already disabled"
	}
}

Function CheckifExtended
{
	Write-Debug "Check if Extended Counters are already enabled"
	if ((Get-ExchangeServer -Identity $ServerName | where {$_.IsMailboxServer -eq $true}))
	{
		$ESERegKey = "System\CurrentControlSet\Services\ESE\Performance"
		$ESEName = "Show Advanced Counters"
		$StoreRegKey = "System\CurrentControlSet\Services\MSExchangeIS\Performance"
		$StoreName = "Library"
		
		$ESEValue = GetValueFromRegistry $Server $ESERegKey $ESEName
		if ($ESEValue -eq 1){$Script:ESEExtendedOn = $true}
		
		#Get Store Value
		$StoreValue = GetValueFromRegistry $Server $StoreRegKey $StoreName
		if ($StoreValue.Contains("mdbperfx.dll")){$Script:StoreExtendedOn = $true}
	}
}
	
Function GetValueFromRegistry ([string]$Server, $regkey, $value) 
{
  $reg = [Microsoft.Win32.RegistryKey]::OpenRemoteBaseKey('LocalMachine', $Server)
  $regKey= $reg.OpenSubKey("$regKey")
  trap [SystemException] 
  {
	if ($_ -match "Requested registry access is not allowed")
	{
		Write-Host ""
		Write-Host "ERROR: Remote registry access denied. Make sure that the account you are logged on as has admministrative permissions on the server specified" -foregroundcolor Red
		Write-Host ""
		exit
	}
	else
	{
		Write-Host ""
		Write-host "ERROR: "$_.Exception.Message -foregroundcolor Red
		Write-Host ""
		exit
	}
  }
  $result = $regkey.GetValue("$value")
  return $result
  #Close the Reg Key
  $regkey.Close()
}

Function Execute_Exmon
{
	Write-Debug "Create Exmon Trace"
	Write-Host "        Enabling Exmon Tracing"
	Write-Host "======================================="
	#Set Exmon duration. Default 30 minutes if not specified
	if ($ExmonDuration -eq "") {$ExmonDuration = "00:30:00"} else {$ExmonDuration = $ExmonDuration}
	
	# Check if Exmon Trace already exists
	$CheckExmon = @(logman query -s $ServerName) -match "Exmon_Trace"
	
	
	if (!$CheckExmon)
    {
		$RunAsUser = read-host "Enter User Name that the Exmon Trace will run under (ie:Domain\Username)"
		$Exmoncmd = "logman create trace Exmon_Trace -p '{2EACCEDF-8648-453e-9250-27F0069F71D2}' -o $filepath$Servername-ExMon -s $Servername -bs 128 -rf " + $ExmonDuration + " -cnf " + "00:05:00" + " -u " + $RunAsUser + " *"
		# Create Exmon Trace
		Write-Debug $Exmoncmd
		Invoke-Expression -Command $Exmoncmd

		while (!($CheckifCreated = @(logman query -s $ServerName) -match "Exmon_Trace"))
		{
			Write-Host "Exmon Traced failed to create. Would you like to try creating it again? " -NoNewline
			$answer = ConfirmAnswer
			if ($answer -eq "yes")
			{
				Invoke-Expression -Command $Exmoncmd
			}
			if ($answer -eq "no")
			{
				Exit
			}
		}
    }
    else
    {
		Write-Host "Exmon_Trace already exists. Checking if already running"
		$CheckifRunning = select-string -InputObject $CheckExmon -pattern "Running" -quiet
		if ($CheckifRunning)
		{
			$cmd = "logman stop Exmon_Trace -s $Servername"
			$StopExmon = Invoke-Expression -Command $Cmd
			Start-Sleep 2
		}
		#Delete and recreate Exmon tracing
		Write-Host "Deleting and recreating Exmon_Trace"
		$cmd = "logman delete Exmon_Trace -s $Servername"
		$DeleteExmon = Invoke-Expression -Command $Cmd 
		# Create Exmon Trace
		$RunAsUser = read-host "Enter User Name that the Exmon Trace will run under (ie:Domain\Username)"
		$Exmoncmd = "logman create trace Exmon_Trace -p '{2EACCEDF-8648-453e-9250-27F0069F71D2}' -o $filepath$Servername-ExMon -s $Servername -bs 128 -rf " + $ExmonDuration + " -cnf " + "00:05:00" + " -u " + $RunAsUser + " *"
		Write-Debug $Exmoncmd
		Invoke-Expression -Command $Exmoncmd

		while (!($CheckifCreated = @(logman query -s $ServerName) -match "Exmon_Trace"))
		{
			Write-Host "Exmon Traced failed to create. Would you like to try creating it again? " -NoNewline
			$answer = ConfirmAnswer
			if ($answer -eq "yes")
			{
				Invoke-Expression -Command $Exmoncmd
			}
			if ($answer -eq "no")
			{
				Exit
			}
		}
		
    }
	if ($Windows2008 -or $Windows2008R2)
	{	
		$cmd = "logman start Exmon_Trace -s $Servername"
		$StartExmon = Invoke-Expression -Command $Cmd
	}
	Write-Host ""
}

Function ConfirmAnswer
{
	$Confirm = "" 
	while ($Confirm -eq "") 
	{ 
		switch (Read-Host "(Y/N)") 
		{ 
			"yes" {$Confirm = "yes"} 
			"no" {$Confirm = "No"} 
			"y" {$Confirm = "yes"} 
			"n" {$Confirm = "No"} 
			default {Write-Host "Invalid entry, please answer question again " -NoNewline} 
		} 
	} 
	return $Confirm 
}

# Function that returns true if the incoming argument is a help request
Function IsHelpRequest
{
	param($argument)
	return ($argument -eq "-?" -or $argument -eq "-help");
}

# Function that displays the help related to this script following
# the same format provided by get-help or <cmdletcall> -?
Function Usage
{
@"

NAME:
`tExPerfwiz.ps1

INFORMATION:
`tSets up Performance data collections for Exchange 2007/2010 servers

SYNTAX:
`tExperfwiz.ps1 [-begin <StringValue>] [-duration <StringValue>] [-end <StringValue>] [-filepath <StringValue>] 
`t[-interval <IntegerValue>] [-maxsize <IntegerValue>] 

PARAMETERS:

`t-begin
`t`tSpecifies when you would like the perfmon data capture to begin.
`t`tThe format must be specified as "01/00/0000 00:00:00"

`t-circular
`t`tTurns on circular logging to save on disk space. Negates default 
`t`tduration of 8 hours.

`t-delete
`t`tDeletes the currently running Perfwiz data collection.

`t-duration
`t`tSpecifies the overall duration of the data collection.
`t`tIf omitted, the default value is (08:00:00) or 8 hours.

`t-end
`t`tSpecifies when you would like the perfmon data capture to end.
`t`tThe format must be specified as "01/00/0000 00:00:00"

`t-EseExtendedOn
`t`tEnables Extended ESE performance counters.

`t-EseExtendedOff
`t`tDisables Extended ESE performance counters.

`t-ExMon
`t`tAdds Exmon Tracing to specified server

`t-ExMonDuration
`t`tSets Exmon trace duration. If not specified, 30 minutes is the 
`t`tdefault duration

`t-filepath
`t`tSpecifies the location to write the Data Collection files to.

`t-full
`t`tDefines a counter set that includes all Counters/instances. 

`t-interval
`t`tSpecifies the interval time between data samples
`t`tIf omitted, the default value is 30 seconds. To change the 
`t`tinterval to 5 seconds, set the value to 5

`t-maxsize
`t`tSpecifies the maximum size of blg file in MB. If omitted, the
`t`tdefault value is 512.

`t-query
`t`tQueries configuration information of previously created
`t`tExchange_Perfwiz Data Collector.

`t-server
`t`tCreates Exchange_Perfwiz data collector on remote server specified.
`t`tIf no server is specified, the local server is used

`t-start
`t`tStarts a previously created Exchange_Perfwiz data collection.

`t-stop
`t`tStops the currently running Perfwiz data collection.

`t-StoreExtendedOn
`t`tEnables Extended Store performance counters.

`t-StoreExtendedOff
`t`tDisables Extended Store performance counters.

`t-threads
`t`tSpecifies whether threads will be added to the data collection.
`t`tIf omitted, threads counters will not be added.

`t-webhelp
`t`tLaunches web help for script

EXAMPLES:
`t- Set duration to 4 hours, change interval to collect data every 5 seconds and set Data location to d:\Logs
`t  .\experfwiz.ps1 -duration 04:00:00 -interval 5 -filepath D:\Logs
`t
`t- Enables Data Collection to begin on January 1st 2010 at 8:00AM
`t  .\experfwiz.ps1 -begin "01/01/2010 08:00:00"
`t
`t- Add threads to the collection set
`t  .\experfwiz.ps1 -threads
`t
`t- Enables Performance Counter data and Exmon data collection
`t  .\experfwiz.ps1 -Exmon
`t
`t- Create collection for all counters/instances
`t  .\experfwiz.ps1 -full
`t
"@
}

# Start Main Processing of Script
# =================================================================================

if ($debug){$DebugPreference = "Continue"}

# Add servers to array
#$Servers = @($Server)

# Check for Usage Statement Request
$args | foreach { 
if (IsHelpRequest $_) { Usage; exit;}
}

#Check for correct params
if (($stop -or $threads -or $query -or $full -or $start -or $delete -or $circular -or $StoreExtendedOn -or $StoreExtendedOff -or $EseExtendedOn -or $EseExtendedOff -or $WebHelp -or $filepath -or $server -or $debug -or $Exmon -or $ExmonDuration -or $quiet) -or $args.count -eq 0){}
else{Write-Host "Incorrect switch entered. Please try again with a valid switch name" -ForegroundColor Red ;exit;}

#Get Exchange Server and OS Info
GetExServerInfo
GetOSVersion

#Param switches
if ($WebHelp)
{
	#Pulls up online help for script
	$ie = new-object -comobject "InternetExplorer.Application"  
	$ie.visible = $true  
	$ie.navigate("http://code.msdn.microsoft.com/ExPerfwiz/")
	exit
}

if ($begin -or $end)
{
	#Check if the format of the begin/end times are correct
	$CheckBegin = $begin | Select-String "^\d{2}\/\d{2}\/\d{4}[ ]\d{2}[:]\d{2}[:]\d{2}"
	$CheckEnd = $end | Select-String "^\d{2}\/\d{2}\/\d{4}[ ]\d{2}[:]\d{2}[:]\d{2}"
	if (($CheckBegin -eq $null -and $begin) -or ($CheckEnd -eq $null -and $end))
	{
		Write-Host ""
		Write-Host "Begin or enter time entered in wrong format. Ensure that the format is similar to `"01/00/0000 00:00:00`"" -ForegroundColor Red
		Write-Host ""
		exit
	}
}

if ($Duration -or $ExmonDuration)
{
	# Check for Duration correctness
	$CheckDuration = $duration | Select-String "^\d{2}[:]\d{2}[:]\d{2}"
	$CheckExmonDuration = $ExmonDuration | Select-String "^\d{2}[:]\d{2}[:]\d{2}"
#	if (($CheckDuration -eq $null) -or ($CheckExmonDuration -eq $null))
	if ($Duration -and ($CheckDuration -eq $null))
	{
		Write-Host ""
		Write-Host "Duration or ExmonDuration time entered in wrong format. Ensure that the format is similar to 00:00:00" -ForegroundColor Red
		Write-Host ""
		exit
	}
	if ($ExmonDuration -and ($CheckExmonDuration -eq $null))
	{
		Write-Host ""
		Write-Host "Duration or ExmonDuration time entered in wrong format. Ensure that the format is similar to 00:00:00" -ForegroundColor Red
		Write-Host ""
		exit
	}
}
if ($StoreExtendedOn){Enable-ExtendedStoreCounters}
if ($StoreExtendedOff){Disable-ExtendedStoreCounters}
if ($ESEExtendedOn){Enable-ExtendedESECounters}
if ($ESEExtendedOff){Disable-ExtendedESECounters}
if ($stop) {StopCollection; exit;}
if ($threads){[bool]$script:threads = $true;}
if ($query){QueryCollection; exit; }
if ($full) {[bool]$script:full = $true;}
if ($start) {StartCollection; exit;}	
if ($delete){DeleteCollection; exit}
#{
#	foreach ($Server in $Servers)
#	{
#		DeleteCollection
#		exit
#	}
#}
# Execute Functions

#foreach ($Server in $Servers)
#{
	
	IsAdmin
	RemoteRegistry
	CheckIfExtended
	CreateCounterList
	CheckifCollectionExists
	WriteCounterConfig
	CreateCounter
#}


# Set Debug Preference back to original
$DebugPreference = $oldDebugPreference 