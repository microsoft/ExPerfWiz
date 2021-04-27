#############################################################################################
# DISCLAIMER:																				#
#																							#
# THE SAMPLE SCRIPTS ARE NOT SUPPORTED UNDER ANY MICROSOFT STANDARD SUPPORT					#
# PROGRAM OR SERVICE. THE SAMPLE SCRIPTS ARE PROVIDED AS IS WITHOUT WARRANTY				#
# OF ANY KIND. MICROSOFT FURTHER DISCLAIMS ALL IMPLIED WARRANTIES INCLUDING, WITHOUT		#
# LIMITATION, ANY IMPLIED WARRANTIES OF MERCHANTABILITY OR OF FITNESS FOR A PARTICULAR		#
# PURPOSE. THE ENTIRE RISK ARISING OUT OF THE USE OR PERFORMANCE OF THE SAMPLE SCRIPTS		#
# AND DOCUMENTATION REMAINS WITH YOU. IN NO EVENT SHALL MICROSOFT, ITS AUTHORS, OR			#
# ANYONE ELSE INVOLVED IN THE CREATION, PRODUCTION, OR DELIVERY OF THE SCRIPTS BE LIABLE	#
# FOR ANY DAMAGES WHATSOEVER (INCLUDING, WITHOUT LIMITATION, DAMAGES FOR LOSS OF BUSINESS	#
# PROFITS, BUSINESS INTERRUPTION, LOSS OF BUSINESS INFORMATION, OR OTHER PECUNIARY LOSS)	#
# ARISING OUT OF THE USE OF OR INABILITY TO USE THE SAMPLE SCRIPTS OR DOCUMENTATION,		#
# EVEN IF MICROSOFT HAS BEEN ADVISED OF THE POSSIBILITY OF SUCH DAMAGES						#
#############################################################################################


# ============== Utility Functions ==============

# Writes output to a log file with a time date stamp
Function Write-Logfile {
    [cmdletbinding()]
    Param
    (
        [Parameter(Mandatory = $true)]
        [string]$String
    )

    # Get our log file path
    $LogFile = Join-path $env:LOCALAPPDATA ExPefwiz.log

    # Get the current date
    [string]$date = Get-Date -Format G

    # Build output string
    [string]$logstring = ( "[" + $date + "] - " + $string)

    # Write everything to our log file and the screen
    $logstring | Out-File -FilePath $LogFile -Append
    Write-Verbose  $logstring
}

Function Convert-OnOffBool {
    [cmdletbinding()]
    [OutputType([bool])]
    Param(
        [Parameter(Mandatory = $true)]
        [string]$tocompare
    )

    switch ($tocompare) {
        On { return $true }
        Default { return $false }
    }
}

# Create a scheduled task to run experfwiz
Function New-PerfWizScheduledTask {
    [CmdletBinding()]
    param (

        [String]
        $Name,

        [string]
        $StartTime,

        [string]
        $Server
    )

    # Going to use schtasks.exe for this instead of New-ScheduledTask
    # New-ScheduledTask is not support on 2012R2 it appears; only shows in Win10 or 2016+
    Write-Logfile -String "Creating Scehduled Task $Name"
    # Create the task ... 2>&1 used to redirect schtasks stderr output to stdopt so it can be processed in PS
    $TaskResult = SCHTASKS /Create /S $Server /RU "SYSTEM" /SC DAILY /TN $Name /TR "logman.exe start $name" /ST $starttime /F 2>&1

    # Check if we have an error
    if (![string]::IsNullOrEmpty(($TaskResult | select-string "Error:"))) {
        Write-Logfile -string "[ERROR] - Creating Scheduled Task"
        Write-Logfile -string [string]$TaskResult
        Throw "UNABLE TO CREATE SCHEDULED TASK: $TaskResult"

    }

}

Function Remove-PerfWizScheduledTask {
    [CmdletBinding()]
    param (
        [String]
        $Name,

        [string]
        $Server
    )

    # Remove scheduled task using schtasks
    $taskRemove = SCHTASKS /DELETE /S $server /TN $name /F 2>&1

    # if success record that
    if (![string]::IsNullOrEmpty(($taskRemove | select-string "SUCCESS:"))) {
        Write-Logfile -string "Removed task $name from server $server"
    }
    # Check if we didn't find the task and record that
    elseif (![string]::IsNullOrEmpty(($taskRemove | select-string "The system cannot find the file specified"))) {
        Write-Logfile -string "Task $name not found on server $server"
    }
    # all other cases throw error
    else {
        Write-Logfile -string "[ERROR] - Removing Scheduled Task"
        Write-Logfile -string [string]$TaskResult
        Throw "UNABLE TO REMOVE SCHEDULED TASK: $TaskResult"

    }
}

Function Get-PerfWizScheduledTask {
    [CmdletBinding()]
    param (
        [string]
        $Name,

        [string]
        $Server
    )

    $taskFound = SCHTASKS /QUERY /TN $name /S $Server /FO csv 2>&1

    # If we found a task with that name return True
    if (![string]::IsNullOrEmpty(($taskFound | select-string "Taskname"))) {
        return true
    }
    # if we didn't find a task with that name return false
    elseif (![string]::IsNullOrEmpty(($taskFound | select-string "The system cannot find the file specified."))) {
        return false
    }
    # If we got an error throw
    else {
        Write-Logfile -string "[ERROR] - Getting Scheduled Task"
        Write-Logfile -string [string]$TaskResult
        Throw "UNABLE TO GET SCHEDULED TASK: $TaskResult"
    }
}

Function Get-ExperfwizUpdate {
    if((Get-Module experfwiz -ListAvailable |Sort-Object -Property version -Descending)[0].version -lt (Find-Module experfwiz -ErrorAction SilentlyContinue).version){
        Write-Warning "Newer Version of Experfwiz avalible from the gallery.  Please run Update-Module Experfwiz"
        Write-Logfile -string "[WARNING] - New Version of Experfwiz Found"
    }
}