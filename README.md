# About ExPerfWiz
ExPerfWiz is a PowerShell based module to help automate the collection of performance data on Exchange 2010, 2013, 2016 and 2019 servers.  Supported operating systems are Windows 2008, 2008 R2, 2012, 2012 R2, 2016 and 2019 Core and Standard.

# Recommended Install
Install the powershell module from the powershell gallery using `Instal-Module ExPerfwiz` from the server where data will be gathered.

Additional information can be found on how to install powershell modules here:
https://docs.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-7

# Manual Install
If the server where data is being gathered doesn't have access to the internet directly then either:

* Install the module to another machine in the org that does have access and use the -server switch
* Download the module from https://github.com/microsoft/ExPerfWiz/releases
* Extract the Zip file to a known location
* From powershell in the extracted path run `Import-Module experfwiz.psd1`
* Change to a different directory than the one with the psd1 file

# How to use
The Module version of Experfwiz provides the following cmdlets to manage the Data Gathering.

### `Get-ExPerfwiz`
Gets ExPerfwiz data collector sets

Switch | Description|Default
-------|-------|-------
Name|Name of the Data Collector set|Experfwiz
Server|Name of the Server |Local Machine
Quiet|Suppress screen Output|False


### `New-ExPerfwiz`
Creates an ExPerfwiz data collector set

Switch | Description|Default
-------|-------|-------
Circular| Enabled or Disable circular logging|Disabled
Duration| How long should the performance data be collected|08:00:00
FolderPath|Output Path for performance logs|NA
Interval|How often the performance data should be collected.|5s
MaxSize|Maximum size of the perfmon log in MegaBytes|250Mb
Name|The name of the data collector set|ExPerfwiz
Quiet|Suppresses screen Output|False
Server|Name of the server where the perfmon collector should be created|Local Machine
StartOnCreate|Starts the counter set as soon as it is created|False
Template| XML perfmon template file that should be loaded to create the data collector set.|Prompt
Threads|Includes threads in the counter set.|False


### `Remove-ExPerfwiz`
Removes an ExPerfwiz data collector set

Switch | Description|Default
-------|-------|-------
Name|Name of the Perfmon Collector set|ExPerfwiz
Server|Name of the server to remove the collector set from|Local Machine
Quiet|Suppresses output to the screen|False

### `Start-ExPerfwiz`
Starts an ExPerfwiz data collector set

Switch | Description|Default
-------|-------|-------
Name|The Name of the Data Collector set to start|ExPerfwiz
Server|Name of the remote server to start the data collector set on.|Local Machine
Quiet|Suppresses output to the screen|False

### `Stop-ExPerfwiz`
Stops an ExPerfwiz data collector set

Switch | Description|Default
-------|-------|-------
Name|Name of the data collector set to stop.|ExPerfwiz
Server|Name of the server to stop the collector set on.|Local Machine
Quiet|Suppresses output to the screen|False

# Example Usage

### Default usage for data gathering

  `New-ExPerfwiz -FolderPath C:\experfwiz`

  When prompted pick the template for the version of Exchange that is being used


### Stop Data Collection

  `Stop-ExPerfwiz`


### Collect data from another server

  `New-ExPerfwiz -FolderPath C:\experfwiz -server RemoteExchServer`

  When prompted pick the template for the version of Exchange being used

# Important Notes
* The default duration is 8 hours to save on disk space meaning that the data collection will stop after 8 hours.

* Using -Threads should only be done if needed to troubleshoot the issue.  It will SIGNIFICANTLY increase the size of the resulting perfmon files.

* If there are any questions about using PowerShell Gallery please see: https://docs.microsoft.com/en-us/powershell/scripting/gallery/overview?view=powershell-7 

# Contributing

This project welcomes contributions and suggestions.  Most contributions require you to agree to a
Contributor License Agreement (CLA) declaring that you have the right to, and actually do, grant us
the rights to use your contribution. For details, visit https://cla.microsoft.com.

When you submit a pull request, a CLA-bot will automatically determine whether you need to provide
a CLA and decorate the PR appropriately (e.g., label, comment). Simply follow the instructions
provided by the bot. You will only need to do this once across all repos using our CLA.

This project has adopted the [Microsoft Open Source Code of Conduct](https://opensource.microsoft.com/codeofconduct/).
For more information see the [Code of Conduct FAQ](https://opensource.microsoft.com/codeofconduct/faq/) or
contact [opencode@microsoft.com](mailto:opencode@microsoft.com) with any additional questions or comments.
