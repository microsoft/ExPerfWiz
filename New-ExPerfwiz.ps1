Function New-ExPerfwiz {
    <#
 
    .SYNOPSIS
    Creates a data collector set for investigating performance related issues.

    .DESCRIPTION
    Creates a performance monitor data collector set from an XML template for the purpose of investigating server performance issues.

    Allows for configuration of the counter set at the time of running the creation command.
        
    .PARAMETER Circular
    Enabled or Disable circular logging
    
    Default is false (Disabled)

    .PARAMETER Duration
    Sets how long should the performance data be collected
    Provided in time span format hh:mm:ss

    Default is 8 hours (08:00:00)

    .PARAMETER FolderPath
    Output Path for performance logs.
    The folder path should exist.

    This paramater is required.

    .PARAMETER Interval
    How often the performance data should be collected.

    Default is 5s (5)

    .PARAMETER MaxSize
    Maximum size of the perfmon log in MegaBytes
    Valid ranges are 100 - 1024

    Default is 250mb (250)

    .PARAMETER Name
    The name of the data collector set
    
    Default is ExPerfwiz

    .PARAMETER Server
    Name of the server where the perfmon collector should be created

    Default is Localhost 

    .PARAMETER StartOnCreate
    Starts the counter set as soon as it is created

    Default is False

    .PARAMETER Template
    XML perfmon template file that should be loaded to create the data collector set.

    Default is to prompt to select a Template from the XMLs provided with this module.

    .PARAMETER Threads
    Includes threads in the counter set.
    *** Including Threads significantly increase the size of perfmon data ***

    Default is False
    

    .OUTPUTS
    Creates a data collector set in Perfmon based on the provided XML file

    Logs all activity into $env:LOCALAPPDATA\ExPefwiz.log file    
	
    .EXAMPLE
    Create a standard ExPerfwiz data collector for troubleshooting performane issues on the local machine.

    New-ExPerfwiz -FolderPath C:\PerfData

    This will prompt the end user to select a template from the provided set and create a default data collector set using that Template.
    The perfmon data will be stored in the C:\PerfData folder

    .EXAMPLE
    Create a custom ExPefwiz data collector on the local machine from a custom template

    New-ExPerfwiz -Name "My Collector" -Duration "01:00:00" -Interval 1 -MaxSize 500 -Template C:\Temp\MyTemplate.xml -Circular $true -Threads $True

    Creates a collector named "My Collector" From the template MyTemplate.xml.
    Circular logging will be enabled along with Threads.
    When started the collector will run for 1 hour.
    It will have a maximum file size of 500MB

    .EXAMPLE
    Create an ExPerfwiz data collector on another server

    New-ExPerfwiz -FolderPath C:\temp\experfwiz -Server OtherServer-01

    Will prompt for template to use.
    Will create a perfmon counter set on the remove server OtherServer-01 with the output folder being C:\temp\experfwiz on that server

    #>

    ### Creates a new experfwiz collector
    [cmdletbinding()]
    Param(
        [bool]
        $Circular = $false,
    
        [timespan]
        $Duration = [timespan]::Parse('8:00:00'),
    
        [Parameter(Mandatory = $true, HelpMessage = "Please provide a valid folder path for output")]
        [string]
        $FolderPath,
    
        [int]
        $Interval = 5,
    
        [ValidateRange(256, 4096)]
        [int]
        $MaxSize = 256,

        [string]
        $Name = "ExPerfwiz",

        [string]
        $Server = $env:ComputerName,

        [switch]
        $StartOnCreate,

        [string]
        $Template,

        [switch]
        $Threads = $false
        <#

        Commenting this out for now ... need to implement the task creation process first

        [string]
        $StartDate,

        [string]
        $EndDate,

        [string]
        $StartTime
        #>
    )

    ### Validate Date Time ###
    $ErrorActionPreference = "Stop"
    if (![string]::IsNullOrEmpty($StartDate)) { $DateToStart = Get-Date $StartDate }
    
    if (![string]::IsNullOrEmpty($EndDate)) { $DateToEnd = Get-Date $EndDate }
    
    if (![string]::IsNullOrEmpty($StartTime)) { $TimeToStart = Get-Date $StartTime }
    $ErrorActionPreference = "Continue"

    
    ### Validate Template ###

    # Build path to templates
    $templatePath = join-path (split-path (Get-Module experfwiz).path -Parent) Templates

    # If no template provided then we need to ask the end user for which one to use
    While ([string]::IsNullOrEmpty($Template)) {
        Out-LogFile -string ("Searching template path: " + $templatePath) 
        $templatesToChoose = Get-ChildItem -Path $templatePath  -Filter *.xml
        Write-Output "`nPlease choose a Template:"

        # Setup counters
        $i = 0

        # Go thru each of the xml templates we found
        Foreach ($file in $templatesToChoose) {
            $i++
            Write-Output ($i.tostring() + "> " + $file.name)        
        }

        # Get the selection from the user
        $selection = Read-Host ("`nChoose Template (1-" + $i + ")")

        # Put the selected xml into template
        $Template = $templatesToChoose[($selection - 1)].FullName
    }

    # Test the template path and log it as good or throw an error
    If (Test-Path $Template) {
        Out-LogFile -string ("Using Template:" + $Template) 
    }
    Else {
        Throw "Cannot find template xml file provided.  Please provide a valid Perfmon template file."
    }

    ### Manipulate Template ###

    # Load the provided template
    [xml]$XML = Get-Content $Template

    # Set Output Location
    $XML.DataCollectorSet.OutputLocation = $FolderPath
    $XML.DataCollectorSet.RootPath = $FolderPath

    # Set Duration
    $XML.DataCollectorSet.SegmentMaxDuration = [string]$Duration.TotalSeconds

    # Set Max File size
    $XML.DataCollectorSet.SegmentMaxSize = [string]$MaxSize

    # Circular logging state
    $XML.DataCollectorSet.PerformanceCounterDataCollector.LogCircular = [string]([int]$Circular * -1)

    # Sample Interval
    $XML.DataCollectorSet.PerformanceCounterDataCollector.SampleInterval = [string]$Interval

    ## Implement schedule
    ## Scenarios:
    ## 1) Start time with Duration
    ## 2) Start time with end time
    ## 3) No Start time with end time -- Invalid
    ## 4) Duration with no time

    # No start time value set so we are just using duration (Scenario 4)
    if ([string]::IsNullOrEmpty($StartDate)) {
        # Make sure the schedule is turned off
        $XML.DataCollectorSet.SchedulesEnabled = "0"
        Out-LogFile -string ("Setting scehdule to disabled") 
    }
    # Need to set the start / end time (Scenario 1 & 2)
    else {

        # Set the start date
        $xml.DataCollectorSet.Schedule.StartDate = [string](Get-Date $DateToStart -format "d")

        # If we have an end date set it
        if ([string]::IsNullOrEmpty($EndDate)) {}
        else {
            $xml.DataCollectorSet.Schedule.EndDate = [string](Get-Date $DateToEnd -format "d")
        }

        # Set the start time
        if ([string]::IsNullOrEmpty($StartTime)) {
            $xml.DataCollectorSet.Schedule.StartTime = [string](Get-Date 00:00:00 -format "t")
        }
        else {
            $xml.DataCollectorSet.Schedule.StartTime = [string](Get-Date $TimeToStart -format "t")
        }
    }    

    # If -threads is specified we need to add it to the counter set
    If ($Threads) {

        Out-LogFile -string "Adding threads to counter set" 

        # Create and set the XML element
        $threadCounter = $XML.CreateElement("Counter")
        $threadCounter.InnerXml = "\Thread(*)\*"

        # Add the XML element
        $XML.DataCollectorSet.PerformanceCounterDataCollector.AppendChild($threadCounter)

    }
    else {}

    # Write the XML to disk
    $xmlfile = Join-Path $env:TEMP ExPerfwiz.xml
    Out-LogFile -string ("Writing Configuration to: " + $xmlfile) 
    $XML.Save($xmlfile)
    Out-Logfile -string ("Importing Collector Set " + $xmlfile + " for " + $server) 
    
    # Import the XML with our configuration
    [string]$logman = logman import -xml $xmlfile -name $Name -s $server
    
    # Check if we generated and error on import
    If ([string]::isnullorempty(($logman | select-string "Error:"))) {
        Out-LogFile -string "Experfwiz imported." 
    }
    else {
        Out-LogFile -string "[ERROR] - Problem importing perfwiz:" 
        Out-LogFile -string $logman 
        Throw $logman
    }    

    # Need to start the counter set if asked to do so
    If ($StartOnCreate) {
        Start-ExPerfwiz -server $Server -Name $Name 
    }
    else {}

}