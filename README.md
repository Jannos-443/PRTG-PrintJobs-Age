# PRTG-PrintJobs-Age
# About

## Project Owner:

Jannos-443

## Project Details

Monitors pending Print Jobs older than x minutes.
Sensor message shows PrinterQueue and Job Owner.

Default Values:
- Age = 1 (minutes)

## HOW TO

1. Place `PRTG-PrintJobs-Age.ps1` under `C:\Program Files (x86)\PRTG Network Monitor\Custom Sensors\EXEXML`

2. Create new Sensor 

   | Settings | Value |
   | --- | --- |
   | EXE/Script Advanced | PRTG-PrintJobs-Age.ps1 |
   | Parameters | -ComputerName %host -Age 5 |
   | Security Context | Use Windows credentials of parent device" or use "-Username" and "-Password" |
   
3. Set the "$IgnorePattern" or "$IgnoreScript" parameter to Exclude PrinterQueues

<br>

## Non Domain or IP

If you connect to **Computers by IP** or to **not Domain Clients** please read [Microsoft Docs](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_remote_troubleshooting?view=powershell-7.1#how-to-use-an-ip-address-in-a-remote-command)

you maybe have to add the target to the TrustedHosts on the PRTG Probe and use explicit credentials.

example (replace all currenty entries): 

    Set-Item -Path WSMan:\localhost\Client\TrustedHosts -Value "ServerIP1,ServerIP2,ServerHostname1"

example want to and and not replace the list:
    
    $curValue = (Get-Item wsman:\localhost\Client\TrustedHosts).value
    Set-Item WSMan:\localhost\Client\TrustedHosts -Value "$curValue,NewServer3.test.com"
    
exmaple PRTG parameter with explicit credentials:
    
    -ComputerName "%host" -Username "%windowsuser" -Password "%windowspassword" -Age 1

## Examples
![PRTG-PrintJobs-Age](media/PrintJobs_OK.png)
![PRTG-PrintJobs-Age](media/PrintJobs_Warning.png)

Exceptions
------------------
You can either use the **parameter $IgnorePattern** to exclude a Printer on sensor basis, or set the **variable $IgnoreScript** within the script. Both variables take a regular expression as input to provide maximum flexibility. These regexes are then evaluated againt the **PrinterQueue Name**

By default, the $IgnoreScript varialbe looks like this:

```powershell
$IgnoreScript = '^(TestExclude123|TestExcludeWildcard.*)$'
```

For more information about regular expressions in PowerShell, visit [Microsoft Docs](https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions).

".+" is one or more charakters
".*" is zero or more charakters
