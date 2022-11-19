<#
    .SYNOPSIS
    Monitors Pending Print Jobs from one Server

    .DESCRIPTION
    Using WMI this script searches for pending Print Jobs.
    Exceptions can be made within this script.

    Copy this script to the PRTG probe EXEXML scripts folder (${env:ProgramFiles(x86)}\PRTG Network Monitor\Custom Sensors\EXEXML)
    and create a "EXE/Script Advanced" sensor. Choose this script from the dropdown and set at least:

    + Parameters: -ComputerName %host
    + Security Context: Use Windows credentials of parent device

    .PARAMETER ComputerName
    The hostname or IP address of the Windows machine to be checked. Should be set to %host in the PRTG parameter configuration.

    .PARAMETER IncludeName
    Regular expression to describe the PrinterName + Jobs for Exampe "Printer 100, 12" where 12 is the JobID
     
      Example: ^(DT_IT_B10_P107, 238|TestPrinter123)$

      Example2: ^(Test123.*|TestPrinter555)$ excludes TestPrinter555 and any Printer starting with Test123

    #https://docs.microsoft.com/en-us/powershell/module/microsoft.powershell.core/about/about_regular_expressions?view=powershell-7.1

    .PARAMETER ExcludeName
    see IncludeName

    .PARAMETER IncludeUser
    see IncludeName

    .PARAMETER ExcludeUser
    see IncludeName

    .PARAMETER UserName
    Provide the Windows user name to connect to the target host via WMI. Better way than explicit credentials is to set the PRTG sensor
    to launch the script in the security context that uses the "Windows credentials of parent device".

    .PARAMETER Password
    Provide the Windows password for the user specified to connect to the target machine using WMI. Better way than explicit credentials is to set the PRTG sensor
    to launch the script in the security context that uses the "Windows credentials of parent device".
    
    .PARAMETER HttpPush
    enables HTTP Push in the Sensor (requires HttpToken, HttpServer and HttpPort)

    .PARAMETER HttpToken
    Set your HTTP Push Sensor Token (Token is available in the Sensor Settings after creating the Sensor)

    .PARAMETER HttpServer
    Set the Target HTTP Push Server (YourPRTGServer FQDN)

    .PARAMETER HttpPort
    Use this parameter if you need to use a HTTP Push Port other than 5050

    .PARAMETER HttpPushUseSSL
    Use this parameter to set the HTTP Push to use HTTPS

    .EXAMPLE
    Sample call from PRTG EXE/Script Advanced
    PRTG-PrintJobs.ps1 -ComputerName %host

    Sample call from Task Scheduler on Remote Computer
    -Command "& 'D:\Powershell\PRTG-PrintJobs.ps1' -ComputerName 'localhost' -HttpPush -HttpServer 'YourPRTGServer' -HttpPort '5050' -HttpToken 'YourHTTPPushToken'"

    .NOTES
    This script is based on the sample by Paessler (https://kb.paessler.com/en/topic/67869-auto-starting-services) and debold (https://github.com/debold/PRTG-WindowsServices)

    Author:  Jannos-443
    https://github.com/Jannos-443/PRTG-PrintJobs
#>
param(
    [string]$ComputerName = "",
    [string]$IncludeName = '',
    [string]$ExcludeName = '',
    [string]$IncludeUser = '',
    [string]$ExcludeUser = '',
    [string]$UserName = "",
    [string]$Password = "",
    [switch] $HttpPush,             #enables http push, usefull if you want to run the Script on the target Server to reduce remote Permissions
    [string] $HttpToken,            #http push token
    [string] $HttpServer,           #http push prtg server hostname
    [string] $HttpPort = "5050",    #http push port (default 5050)
    [switch] $HttpPushUseSSL        #use https for http push
)

#Catch all unhandled Errors
trap{
    if($session -ne $null)
        {
        Remove-CimSession -CimSession $session -ErrorAction SilentlyContinue
        }
    $Output = "line:$($_.InvocationInfo.ScriptLineNumber.ToString()) char:$($_.InvocationInfo.OffsetInLine.ToString()) --- message: $($_.Exception.Message.ToString()) --- line: $($_.InvocationInfo.Line.ToString()) "
    $Output = $Output.Replace("<","")
    $Output = $Output.Replace(">","")
    $Output = $Output.Replace("#","")
    Write-Output "<prtg>"
    Write-Output "<error>1</error>"
    Write-Output "<text>$($Output)</text>"
    Write-Output "</prtg>"
    Exit
}

# Error if there's anything going on
$ErrorActionPreference = "Stop"


if ($ComputerName -eq "") {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>You must provide a computer name to connect to</text>"
    Write-Output "</prtg>"
    Exit
}

# Generate Credentials Object, if provided via parameter
try{
    if($UserName -eq "" -or $Password -eq "") 
        {
        $Credentials = $null
        }
    else 
        {
        $SecPasswd  = ConvertTo-SecureString $Password -AsPlainText -Force
        $Credentials= New-Object System.Management.Automation.PSCredential ($UserName, $secpasswd)
        }
} catch {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error Parsing Credentials ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
}

$WmiClass = "win32_printjob"

# Get list of Jobs that are older than the Age.
try 
    {
    if ($null -eq $Credentials) 
        {
        $PrintJobs = Get-CimInstance -Namespace "root\CIMV2" -ClassName $WmiClass -ComputerName $ComputerName
        } 
    
    else 
        {
        $session = New-CimSession -ComputerName $ComputerName -Credential $Credentials
        $PrintJobs = Get-CimInstance -Namespace "root\CIMV2" -ClassName $WmiClass -CimSession $session
        Start-Sleep -Seconds 1
        Remove-CimSession -CimSession $session
        }

    } 
catch 
    {
    Write-Output "<prtg>"
    Write-Output " <error>1</error>"
    Write-Output " <text>Error connecting to $ComputerName ($($_.Exception.Message))</text>"
    Write-Output "</prtg>"
    Exit
    }

#Remove Ignored Printer
if ($IncludeName -ne "") 
    {
    $PrintJobs = $PrintJobs | Where-Object {$_.Name -match $IncludeName}  
    }

if ($IncludeUser -ne "") 
    {
    $PrintJobs = $PrintJobs | Where-Object {$_.owner -match $IncludeUser}  
    }

if ($ExcludeName -ne "") 
    {
    $PrintJobs = $PrintJobs | Where-Object {$_.Name -notmatch $ExcludeName}  
    }

if ($ExcludeUser -ne "") 
    {
    $PrintJobs = $PrintJobs | Where-Object {$_.owner -notmatch $ExcludeUser}  
    }

#Select oldest Job
if ($PrintJobs)
    {
    $OldestJob = ((Get-Date) - ($PrintJobs | Sort-Object -Property TimeSubmitted | Select-Object -Property TimeSubmitted -ExpandProperty TimeSubmitted -First 1)).TotalSeconds
    $OldestJob = [math]::Round($OldestJob)
    if($OldestJob -gt 42163632000 )
        {
        $OldestJob = 42163632000
        }
    }
else {
    $OldestJob = 0
    }

$count = ($PrintJobs | Measure-Object).count

$ErrorText = ""

$xmlOutput = '<prtg>'

#Check if pending Jobs exists
if($PrintJobs)
    {
    foreach($PrintJob in $PrintJobs)
        {
        if($PrintJob.Name -like "*,*"){
            $PName = $PrintJob.Name.Substring(0,$PrintJob.Name.LastIndexOf(","))
            }
        else{
            $PName = $PrintJob.Name
            }
        $ErrorText += "Printer=`"$($PName)`" Owner=`"$($PrintJob.owner)`"; "
          
        }

    $xmlOutput += "<text>$($count) PrintJob(s): $($ErrorText)</text>"
    }

else{
    $xmlOutput += "<text>No PrintJobs pending.</text>"
    }


$xmlOutput += "<result>
        <channel>pending PrintJobs</channel>
        <value>$count</value>
        <unit>Count</unit>
        </result>
        <result>
        <channel>oldest PrintJob</channel>
        <value>$OldestJob</value>
        <unit>TimeSeconds</unit>
        <limitmode>1</limitmode>
        <LimitMaxError>120</LimitMaxError>
        <LimitMaxWarning>60</LimitMaxWarning>
        </result>"

$xmlOutput += "</prtg>"

#region: Http Push
if($httppush)
    {
    if($HttpPushUseSSL)
        {$httppushssl = "https"}
    else
        {$httppushssl = "http"}

    Add-Type -AssemblyName system.web

    $Answer=Invoke-Webrequest -method "GET" -URI ("$($httppushssl)://$($httpserver):$($httpport)/$($httptoken)?content=$(([System.Web.HttpUtility]::UrlEncode($xmloutput)))") -usebasicparsing

    if ($answer.Statuscode -ne 200)
        {
        Write-Output "<prtg>"
        Write-Output "<error>1</error>"
        Write-Output "<text>http push failed</text>"
        Write-Output "</prtg>"
        Exit
        }
    }
#endregion

#finish Script - Write Output

Write-Output $xmlOutput