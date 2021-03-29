Function Get-ExPerfwiz {
    <#

    .SYNOPSIS
    Get information about a data collector set.

    .DESCRIPTION
    Gets information about a data collector set on the local or remote server.

    .PARAMETER Name
    Name of the Data Collector set

    Default Exchange_Perfwiz

    .PARAMETER Server
    Name of the server

    Default LocalHost

	.OUTPUTS
    Logs all activity into $env:LOCALAPPDATA\ExPefwiz.log file

    .EXAMPLE
    Get info on the default collector set

    Get-ExPerfwiz

    .EXAMPLE
    Get info on a collector set on a remote server

    Get-ExPerfwiz -Name "My Collector Set" -Server RemoteServer-01

    #>
    [cmdletbinding()]
    param (
        [string]
        $Name,

        [string]
        $Server = $env:ComputerName

    )

    Out-LogFile -string ("Getting ExPerfwiz: " + $server)


    # If no name was provided then we need to return all counters logman finds
    if ([string]::IsNullOrEmpty($Name)) {

        # Returns all found counter sets
        $logmanAll = logman query -s $server

        If (!([string]::isnullorempty(($logmanAll | select-string "Error:")))) {
            throw $logmanAll[-1]
        }

        # Process the string return into a set of counter names
        $i = -3
        [array]$perfLogNames = $null

        While (!($logmanAll[$i] | select-string "---")) {

            # pull the first 40 characters then trim and trailing spaces
            [array]$perfLogNames += $logmanAll[$i].substring(0, 40).trimend()
            $i--
        }

    }
    # If a name was provided put just that into the array
    else {
        [array]$perfLogNames += $Name
    }

    # Query each counter found in turn to get their details
    foreach ($counterName in $perfLogNames) {

        $logman = logman query $counterName -s $Server

        # Quick error check
        If (!([string]::isnullorempty(($logman | select-string "Error:")))) {
            throw $logman[-1]
        }

        # Convert the output of logman into an object
        $logmanObject = New-Object -TypeName PSObject

        foreach ($line in $logman) {

            $linesplit = $line.split(":").trim()

            switch (($linesplit)[0]) {
                'Name' {
                    # Skip the path to the perfmon inside the counter set
                    if ($linesplit[1] -like "*\*") {}
                    # Set the name and push it into a variable to use later
                    else {
                        $logmanObject | Add-Member -MemberType NoteProperty -Name $linesplit[0] -Value $linesplit[1] -Force
                        $logManName = $linesplit[1]
                    }
                }
                'Status' { $logmanObject | Add-Member -MemberType NoteProperty -Name $linesplit[0] -Value $linesplit[1] }
                'Root Path' {
                    if ($linesplit[1].contains("%")) {
                        $logmanObject | Add-Member -MemberType NoteProperty -Name "RootPath" -Value $linesplit[1]
                        $logmanObject | Add-Member -MemberType NoteProperty -Name "OutputPath" -Value $linesplit[1]
                    }
                    else {
                        $logmanObject | Add-Member -MemberType NoteProperty -Name "RootPath" -Value (Resolve-path ($linesplit[1] + ":" + $linesplit[2]))
                        $logmanObject | Add-Member -MemberType NoteProperty -Name "OutputPath" -Value (Join-path (($linesplit[1] + ":" + $linesplit[2])) ($env:ComputerName + "_" + $logManName))
                    }
                }
                'Segment' { $logmanObject | Add-Member -MemberType NoteProperty -Name $linesplit[0] -Value $linesplit[1] }
                'Schedules' { $logmanObject | Add-Member -MemberType NoteProperty -Name $linesplit[0] -Value $linesplit[1] }
                'Duration' { $logmanObject | Add-Member -MemberType NoteProperty -Name "Duration" -Value (New-TimeSpan -Seconds ([int]($linesplit[1].split(" "))[0])) }
                'Segment Max Size' { $logmanObject | Add-Member -MemberType NoteProperty -Name 'MaxSize' -Value (($linesplit[1].replace(" ", "")) / 1MB) }
                'Run As' { $logmanObject | Add-Member -MemberType NoteProperty -Name "RunAs" -Value $linesplit[1] }
                'Start Date' { $logmanObject | Add-Member -MemberType NoteProperty -Name "StartDate" -Value $linesplit[1] }
                'Start Time' { $logmanObject | Add-Member -MemberType NoteProperty -Name "StartTime" -Value ($line.split(" ")[-2] + " " + $line.split(" ")[-1]) }
                'End Date' { $logmanObject | Add-Member -MemberType NoteProperty -Name "EndDate" -Value $linesplit[1] }
                'Days' { $logmanObject | Add-Member -MemberType NoteProperty -Name "Days" -Value $linesplit[1] }
                'Type' { $logmanObject | Add-Member -MemberType NoteProperty -Name "Type" -Value $linesplit[1] }
                'Append' { $logmanObject | Add-Member -MemberType NoteProperty -Name "Append" -Value (Convert-OnOffBool($linesplit[1])) }
                'Circular' { $logmanObject | Add-Member -MemberType NoteProperty -Name "Circular" -Value (Convert-OnOffBool($linesplit[1])) }
                'Overwrite' { $logmanObject | Add-Member -MemberType NoteProperty -Name "Overwrite" -Value (Convert-OnOffBool($linesplit[1])) }
                'Sample Interval' { $logmanObject | Add-Member -MemberType NoteProperty -Name "SampleInterval" -Value (($linesplit[1].split(" "))[0]) }
                Default {}
            }

        }

        # Add customer PS Object type for use with formatting files
        $logmanObject.pstypenames.insert(0, 'Experfwiz.Counter')

        # Add each object to the return array
        $logmanObject
    }
}