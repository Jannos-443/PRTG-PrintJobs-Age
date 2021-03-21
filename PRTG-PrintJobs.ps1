<#
    .SYNOPSIS
    Monitors Pending Print Jobs from one Server

    .DESCRIPTION
    Using WMI this script searches for pending Print Jobs.
    Exceptions can be made within this script by changing the variable $IgnoreScript. This way, the change applies to all PRTG sensors 
    based on this script. If exceptions have to be made on a per sensor level, the script parameter $IgnorePattern can be used.

    Copy this script to the PRTG probe EXEXML scripts folder (${env:ProgramFiles(x86)}\PRTG Network Monitor\Custom Sensors\EXEXML)
    and create a "EXE/Script Advanced" sensor. Choose this script from the dropdown and set at least:

    + Parameters: -ComputerName %host
    + Security Context: Use Windows credentials of parent device

    .PARAMETER ComputerName
    The hostname or IP address of the Windows machine to be checked. Should be set to %host in the PRTG parameter configuration.

    .PARAMETER IgnorePattern
    Regular expression to describe the PrinterName + Jobs for Exampe "Printer 100, 12" where 12 is the JobID
     
      Example: ^(DT_IT_B10_P107, 238|TestPrinter123)$

      Example2: ^(Test123.*|TestPrinter555)$ excluded Test12345 und alles mit 

    #https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.1

    .PARAMETER UserName
    Provide the Windows user name to connect to the target host via WMI. Better way than explicit credentials is to set the PRTG sensor
    to launch the script in the security context that uses the "Windows credentials of parent device".

    .PARAMETER Password
    Provide the Windows password for the user specified to connect to the target machine using WMI. Better way than explicit credentials is to set the PRTG sensor
    to launch the script in the security context that uses the "Windows credentials of parent device".

    .EXAMPLE
    Sample call from PRTG EXE/Script Advanced
    PRTG-PrintJobs.ps1 -ComputerName %host -Age 5

    .NOTES
    This script is based on the sample by Paessler (https://kb.paessler.com/en/topic/67869-auto-starting-services) and debold (https://github.com/debold/PRTG-WindowsServices)

    Author:  Jannos-443
    https://github.com/Jannos-443/PRTG-PrintJobs
#>
param(
    [string]$ComputerName = "",
    [string]$IgnorePattern = "",
    [int]$Age = "1", # Jobs is older than x minutes
    [string]$UserName = "",
    [string]$Password = ""
)


#Catch all unhandled Errors
trap{
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>$($_.ToString() - $($_.ScriptStackTrace))</text>"
    Write-Output "</prtg>"
    Exit
}


if ($ComputerName -eq "") {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>You must provide a computer name to connect to</text>"
    Write-Output "</prtg>"
    Exit
}

# Error if there's anything going on
$ErrorActionPreference = "Stop"

# Generate Credentials Object, if provided via parameter
if ($UserName -eq "" -or $Password -eq "") {
   $Credentials = $null
} else {
    $SecPasswd  = ConvertTo-SecureString $Password -AsPlainText -Force
    $Credentials= New-Object System.Management.Automation.PSCredential ($UserName, $secpasswd)
}

$WmiClass = "win32_printjob"
$old = (Get-Date).AddMinutes(-$Age)

# Get list of Jobs that are older than the Age.
try {
    if ($null -eq $Credentials) {
        $PrintJobs = Get-WmiObject -class $WmiClass -namespace "root\CIMV2" -ComputerName $ComputerName | Where-Object {$_.ConvertToDateTime($_.timesubmitted) -lt "$old"}
    } else {
        $PrintJobs = Get-WmiObject -class $WmiClass -namespace "root\CIMV2" -ComputerName $ComputerName -Credential $Credentials | Where-Object {$_.ConvertToDateTime($_.timesubmitted) -lt "$old"}  
    }
} catch {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error connecting to $ComputerName ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
}

# hardcoded list that applies to all hosts
$IgnoreScript = '^(TestIgnore)$' 

#Remove Ignored Printer
if ($IgnorePattern -ne "") {
    $PrintJobs = $PrintJobs | where {$_.Name -notmatch $IgnorePattern}  
}

if ($IgnoreScript -ne "") {
    $PrintJobs = $PrintJobs | where {$_.Name -notmatch $IgnoreScript}  
}


$Count = ($PrintJobs | Measure-Object).Count
$Error = ""

$xmlOutput = '<prtg>'

#Check if pending Jobs exists
if($Count -ge 1)
    {
    foreach($PrintJob in $PrintJobs)
        {
        if($PrintJob.Name -like "*,*"){
            $PName = $PrintJob.Name.Substring(0,$PrintJob.Name.IndexOf(","))
            }
        else{
            $PName = $PrintJob.Name
            }
        $Error += "Printer=$($PName) Owner=$($PrintJob.owner); "
          
        }

    $xmlOutput = $xmlOutput + "<text>$($count) PrintJob(s) older $($age) min; $($Error)</text>"
    }

else{
    $xmlOutput = $xmlOutput + "<text>No PrintJobs older than $($Age) min.</text>"
    }


$xmlOutput = $xmlOutput + "<result>
        <channel>Jobs pending</channel>
        <value>$count</value>
        <unit>Count</unit>
        <limitmode>1</limitmode>
        <LimitMaxError>1.5</LimitMaxError>
        <LimitMaxWarning>0.1</LimitMaxWarning>
        </result>"

$xmlOutput = $xmlOutput + "</prtg>"

$xmlOutput
