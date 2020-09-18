Function Start-ExPerfwiz {
        <#
 
    .SYNOPSIS
    Starts a data collector set

    .DESCRIPTION
    Starts a data collector set on the local server or a remote server.

    .PARAMETER Name
    The Name of the Data Collector set to start

    Default ExPerfwiz

    .PARAMETER Server
    Name of the remote server to start the data collector set on.

    Default LocalHost

    .PARAMETER Quiet
    Suppresses output to the screen

	.OUTPUTS
     Logs all activity into $env:LOCALAPPDATA\ExPefwiz.log file   
     
	.EXAMPLE
    Start the default data collector set on this server.

    Start-ExPerfwiz

    .EXAMPLE
    Start a collector set on another server.

    Start-ExPerfwiz -Name "My Collector Set" -Server RemoteServer-01

    #>
    param (
        [string]
        $Name = "Experfwiz",
        
        [string]
        $Server = $env:ComputerName,

        [bool]
        $Quiet = $false
    )
    
    Out-LogFile -string ("Starting ExPerfwiz: " + $server) -quiet $Quiet
    
    # Remove the experfwiz counter set
    [string]$logman = logman start -name $Name -s $server

    # Check if we have an error and throw and error if needed.
    If ([string]::isnullorempty(($logman | select-string "Error:"))) {
        Out-LogFile "ExPerfwiz Started" -quiet $Quiet
    }
    else {
        Out-LogFile "[ERROR] - Unable to Start Collector" -quiet $Quiet
        Out-LogFile $logman -quiet $Quiet
        Throw $logman
    }
}