Function Step-ExPerfwizSize {

    [cmdletbinding()]
    param (
        [string]
        $Name = "Exchange_Perfwiz",

        [string]
        $Server = $env:ComputerName,

        [int]
        $Size

    )

    # Step up the size of the perfwiz by 1
    $perfmon = Get-ExPerfwiz -Name $Name -Server $Server
    $newSize = $perfmon.maxsize + 1

    # increment the size
    [string]$logman = $null
    [string]$logman = logman update -name $Name -s $Server -max $newSize

    # If we find an error throw
    # Otherwise nothing
    if ($logman | select-string "Error:") {      
        Out-LogFile -string "[ERROR] - Problem stepping perfwize size:"
        Out-LogFile -string $logman
        Throw $logman
    }
    else {}
}