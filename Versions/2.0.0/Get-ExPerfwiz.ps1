Function Get-ExPerfwiz {
        <#
 
    .SYNOPSIS
    Get information about a data collector set.

    .DESCRIPTION
    Gets information about a data collector set on the local or remote server.

    .PARAMETER Name
    Name of the Data Collector set

    Default ExPerfwiz

    .PARAMETER Server
    Name of the server

    Default LocalHost

    .PARAMETER Quiet
    Suppresses output to the screen

	.OUTPUTS
    Logs all activity into $env:LOCALAPPDATA\ExPefwiz.log file 
	
    .EXAMPLE
    Get info on the default collector set

    Get-ExPerfwiz

    .EXAMPLE
    Get info on a collector set on a remote server

    Get-ExPerfwiz -Name "My Collector Set" -Server RemoteServer-01

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
    $logman = logman query -name $Name -s $server
    $formatlogman = $logman -join "`n`r" | out-string

    # Check if we have an error and throw and error if needed.
    If ([string]::isnullorempty(($logman | select-string "Error:"))) {
        Out-LogFile $formatlogman -quiet $Quiet
    }
    else {
        Out-LogFile "[ERROR] - Unable to Get collector" -quiet $Quiet
        Out-LogFile $logman -quiet $Quiet
        Throw $logman
    }
}