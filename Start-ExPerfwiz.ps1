Function Start-ExPerfwiz {
        <#
 
    .SYNOPSIS
    Starts a data collector set

    .DESCRIPTION
    Starts a data collector set on the local server or a remote server.

    .PARAMETER Name
    The Name of the Data Collector set to start

    Default Exchange_Perfwiz

    .PARAMETER Server
    Name of the remote server to start the data collector set on.

    Default LocalHost

	.OUTPUTS
     Logs all activity into $env:LOCALAPPDATA\ExPefwiz.log file   
     
	.EXAMPLE
    Start the default data collector set on this server.

    Start-ExPerfwiz

    .EXAMPLE
    Start a collector set on another server.

    Start-ExPerfwiz -Name "My Collector Set" -Server RemoteServer-01

    #>
    [cmdletbinding()]
    param (
        [string]
        $Name = "Exchange_Perfwiz",
        
        [string]
        $Server = $env:ComputerName
            )
    
    Out-LogFile -string ("Starting ExPerfwiz: " + $server) 
    
    # Remove the experfwiz counter set
    [string]$logman = logman start -name $Name -s $server

    # Check if we have an error and throw and error if needed.
    If ([string]::isnullorempty(($logman | select-string "Error:"))) {
        Out-LogFile "ExPerfwiz Started" 
    }
    else {
        Out-LogFile "[ERROR] - Unable to Start Collector" 
        Out-LogFile $logman 
        Throw $logman
    }
}