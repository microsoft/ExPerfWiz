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
    $logstring | Out-File -FilePath $LogFile -Append -Confirm:$false
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

Function Get-ExperfwizUpdate {
    
    # Need to suppress error output here
    $CEAP = $ErrorActionPreference
    $ErrorActionPreference = 'SilentlyContinue'

    # Get the latest version of the module in the gallery
    $Request = Invoke-WebRequest -Uri https://www.powershellgallery.com/packages/experfwiz -TimeoutSec 10

    #Set error action back it back to what is was
    $ErrorActionPreference = $CEAP

    # If we get back a 200 then compare the versions otherwise just move on
    If ($Request.statuscode -eq 200) {
        # Pull the latest version off the response URI
        [Version]$Version = $Request.BaseResponse.ResponseUri.Segments[-1]
        Write-Logfile ("Found Version: " + $Version)

        # Compare the versions
        if ((Get-Module experfwiz | Sort-Object -Property version -Descending)[0].version -lt $Version) {
            Write-Warning "Newer Version of Experfwiz avalible from the gallery.  Please run Update-Module Experfwiz"
            Write-Logfile -string "[WARNING] - New Version of Experfwiz Found"
        }
        else {Write-Logfile "Version is Current"}
    }
    else { Write-Logfile ("Update Status Code: " + $Request.statuscode) }
}