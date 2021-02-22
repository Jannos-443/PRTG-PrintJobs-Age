# PRTG-PrintJobs-Age
# About

## Project Owner:

Jannos-443

## Project Details

Monitors pending Print Jobs older than x minutes.
Sensor message shows Printername and Job Owner.

Default Values:
- Age = 1 (minutes)

## HOW TO

1. Place "PRTG-PrintJobs-Age.ps1" under "C:\Program Files (x86)\PRTG Network Monitor\Custom Sensors\EXE"

2. Create new Sensor 
   - EXE/Script = PRTG-PrintJobs-Age.ps1
   - Parameter = "-ComputerName %host -Age 5"

3. Change "Value" Channel >> "Lookups and Limits" to "Enable alerting based on limits" 
   - This example shows Warning for 1 hanging Job and Error for more than 1 Job.
     - Upper Error Limit 1
     - Upper Warning Limit 0,1
![PRTG-PrintJobs-Age](media/Sensor-Limit-Channel.png)
![PRTG-PrintJobs-Age](media/Sensor-Limit.png)

4. Set the "$IgnorePattern" or "$IgnoreScript" parameter to Exclude PrinterQueues

## Examples
![PRTG-PrintJobs-Age](media/Print_Limit_OK.png)
![PRTG-PrintJobs-Age](media/Print_Limit_Warning.png)

Printer exceptions
------------------
You can either use the **parameter $IgnorePattern** to exclude a service on sensor basis, or set the **variable $IgnoreScript** within the script. Both variables take a regular expression as input to provide maximum flexibility. These regexes are then evaluated againt the **PrinterQueue Name**

By default, the $IgnoreScript varialbe looks like this:

```powershell
$IgnoreScript = '^(TestExclude123|TestExcludeWildcard.*)$'
```

For more information about regular expressions in PowerShell, visit [Microsoft Docs](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions).

".+" is one or more charakters
".*" is zero or more charakters
