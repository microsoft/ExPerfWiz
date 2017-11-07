# Download
Use this link: https://github.com/Microsoft/ExPerfWiz/blob/master/ExPerfwiz.zip?raw=true

# About ExPerfWiz
ExPerfWiz is a PowerShell based script to help automate the collection of performance data on Exchange 2007, 2010, 2013 and 2016 servers.  Supported operating systems are Windows 2003, 2008, 2008 R2, 2012, 2012 R2 and 2016.

# Important Notes
* The default duration is 8 hours to save on disk space meaning that the data collection will stop after 8 hours. If you should need a longer duration, please review the switches below for the best possible configuration that meets your needs.
The below table outlines what parameters ExPerfWiz can accept.

* Before running this script, you must do one or both of the following:

  - Set PowerShell’s execution policy to Unrestricted using (Set-ExecutionPolicy Unrestricted)
  
  - Files downloaded from the internet using Internet Explorer are automatically blocked from running. Follow the below steps to Unblock this script from running.
  
    - Save the script file on your computer.
 
    - Click Start, click My Computer, and navigate to the saved script file.
 
    - Right-click the script file, and then click "Properties."
 
    - Click "Unblock."
 
# Usage Examples

  - Set duration to 4 hours, change interval to collect data every 5 seconds and set data location to d:\Logs
  
    *.\experfwiz.ps1 -duration 04:00:00 -interval 5 -filepath D:\Logs*

  - Stop data collection
  
    *.\experfwiz.ps1 -stop*

  - Enables Perf Data collection on remote server MBXServer with interval set to 5 seconds and set Data location to d:\Logs
  
    *.\experfwiz.ps1 -server MBXServer -interval 5 -filepath D:\Logs*

  - Enables Perf Data collection on the local server, enabled Exmon data collection with a duration of 1 hour. Note that new ETL files are created every 5 minutes. This is hardcoded and cannot be changed.
  
    *.\experfwiz.ps1 -Exmon -exmonduration 01:00:00*

# Parameters

Parameter | Description
--------- | -----------
-help or -? | Provides help regarding the overall usage of the script
-begin | Specifies when you would like the ExPerfWiz data capture to begin.  The format must be specified as: “01/00/0000 00:00:00”
-circular | Turns on circular logging to save on disk space. Negates default duration of 8 hours
-ConvertToCsv | Converts existing BLG files to CSV. Must include –filepath (to BLG files). This can be run from any machine with PowerShell.
-CsvFilepath | Path where converted CSV files should be stored
-debug | Runs in debug mode to help troubleshoot runtime issues.
-delete | Deletes the currently running ExPerfWiz data collection
-duration | Specifies the overall duration of the data collection. If omitted, the default value is (08:00:00) or 8 hours
-end | Specifies when you would like the ExPerfWiz data capture to end.  The format must be specified as: “01/00/0000 00:00:00”
-EseExtendedOn | Enables Extended ESE performance counters. Currently not supported with Exchange 2013.
-EseExtendedOff | Disables Extended ESE performance counters. Currently not supported with Exchange 2013.
-ExMon | Adds ExMon tracing to specified server.  Exchange 2013 support added in 1.4.2.
-ExMonDuration | Sets ExMon trace duration.  If not specified, 30 minutes is the default duration.  Exchange 2013 support added in 1.4.2.
 -ExmonOnly | Only collect ExMon capture - do NOT collect Performance data. Version 1.4.5+
-filepath | Sets the directory location of where the BLG file will be stored. Default Location is C:\Perflogs.
-interval | Specifies the interval time between data samples. If omitted, the default value is 5 seconds.  NOTE: Exchange 2013 and Server 2012 introduced a large number of counters that were not available in previous version of Exchange/Windows.  Because of this, using a value of less than 5 will result in very large files, very quickly.  Please be sure the storage location has enough space.
-maxsize | Specifies the maximum size of BLG file in MB. If omitted, the default value is 512. NOTE: Starting with v1.4, Exchange 2013/2016 defaults to 1024MB
-nofull | Run ExPerfWiz in role-based mode.  This will collect counters based on the roles installed.  Currently not supported with Exchange 2013/2016.
-query | Queries configuration information of previously created Exchange_Perfwiz Data Collector
 -quiet | Silently run ExPerfWiz (no prompts).
-server | Creates ExPerfWiz data collector on remote server specified.  If no server is specified, the local server will be used.
-skipUpdateCheck | Skips the default behavior of ExPerfWiz to check for an update.
-start | Starts a previously created ExPerfWiz data collection
-stop | Stops the currently running ExPerfWiz data collection.  This is useful if you need to stop ExPerfWiz before the configured duration is met.  For example, the default duration is 8 hours, however you may want to stop the collection process after only 4 hours.
-StoreExtendedOn | Enables Extended Store performance counters. Currently not supported with Exchange 2013.
-StoreExtendedOff | Disables Extended Store performance counters. Currently not supported with Exchange 2013.
-threads | Specifies whether threads will be added to the data collection. If omitted, threads counters will not be added to the collection
-webhelp | Launches web help for script

# Known Issues
  - If you had downloaded version 1.4.6 between 6/29/17 and 6/30/17, you may have a corrupt file that shows "Error.  The command completed successfully."  If so, download the latest version again using the link above.
  - Won't work on Windows 2003 if default system is something other than English (Get off 2003!)
  - Other bugs with Windows 2003 (Again, get off 2003!)

# Change Log
  - 11/7/17 1.4.7.3 (brenle)
    - Fixed bug for running on Windows Server 2016
  - 8/23/17 1.4.7.2 (brenle)
    - Changed default interval to 5 seconds
  - 7/10/17 1.4.7.1 (brenle)
    - Fixed Windows 2016 bug (thanks to shaneto)
    - Fixed quiet switch bug
    - Improved update check process and reliability
  - 7/5/17 1.4.7 (brenle)
    - Added Windows 2016 support.
    - Skip update check when -start and -stop parameters are used
  - 6/29/17 1.4.6 (brenle)
    - Added auto check for update on run (Powershell 3+)

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
