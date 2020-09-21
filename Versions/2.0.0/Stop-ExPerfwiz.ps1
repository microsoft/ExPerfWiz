Function Stop-ExPerfwiz {
    <#
 
    .SYNOPSIS
    Stop a data collector set.

    .DESCRIPTION
    Stops a data collector set on the local or remote server.

    .PARAMETER Name
    Name of the data collector set to stop.

    Default ExPerfwiz

    .PARAMETER Server
    Name of the server to stop the collector set on.

    Default LocalHost

    .PARAMETER Quiet
    Suppresses output to the screen

	.OUTPUTS
    Logs all activity into $env:LOCALAPPDATA\ExPefwiz.log file   
    
    .EXAMPLE
    Stop the default data collector set on the local server

    Stop-ExPerfwiz

    .EXAMPLE
    Stop a data colletor set on a remote server

    Stop-ExPerfwiz -Name "My Collector Set" -Server RemoteServer-01

    #>
    param (
        [string]
        $Name = "Experfwiz",

        [string]
        $Server = $env:ComputerName,

        [switch]
        $Quiet = $false
    )
    
    Out-LogFile -string ("Stopping ExPerfwiz: " + $server) -quiet $Quiet
    
    # Remove the experfwiz counter set
    [string]$logman = logman stop -name $Name -s $server

    # Check if we have an error and throw and error if needed.
    If ([string]::isnullorempty(($logman | select-string "Error:"))) {
        Out-LogFile "ExPerfwiz Stopped" -quiet $Quiet
    }
    else {
        Out-LogFile "[ERROR] - Unable to Stop Collector" -quiet $Quiet
        Out-LogFile $logman -quiet $Quiet
        Throw $logman
    }
}