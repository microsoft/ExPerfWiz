Function Remove-ExPerfwiz {
    <#
 
    .SYNOPSIS
    Removes data collector sets from perfmon

    .DESCRIPTION
    Used to remove data collector sets from perfmon.

    .PARAMETER Name
    Name of the Perfmon Collector set

    Default ExPerfwiz

    .PARAMETER Server
    Name of the server to remove the collector set from

    Default LocalHost
    
    .PARAMETER Quiet
    Suppresses output to the screen

    Default False

    .OUTPUTS
    Logs all activity into $env:LOCALAPPDATA\ExPefwiz.log file   
	
    .EXAMPLE
    Remove a collector set on the local machine

    Remove-ExPerfwiz -Name "My Collector Set"

    .EXAMPLE
    Remove a collect set on another server

    Remove-ExPerfwiz -Server RemoteServer-01


    #>

    param (

        [string]
        $Name = "Experfwiz",

        [string]
        $Server = $env:ComputerName,

        [switch]
        $Quiet = $false
    )
    
    Out-LogFile -string ("Removing Experfwiz for: " + $server) -quiet $Quiet
    
    # Remove the experfwiz counter set
    [string]$logman = logman delete -name $Name -s $server

    # Check if we have an error and throw and error if needed.
    If ([string]::isnullorempty(($logman | select-string "Error:"))) {
        Out-LogFile "ExPerfwiz removed" -quiet $Quiet
    }
    else {
        Out-LogFile "[ERROR] - Unable to remove Collector" -quiet $Quiet
        Out-LogFile $logman -quiet $Quiet
        Throw $logman
    }
}