<#
System requirements
PSVersion: 5.1.19041.3803
PSEdition: Desktop
PSCompatibleVersions: 1.0, 2.0, 3.0, 4.0, 5.0, 5.1.19041.3803
BuildVersion: 10.0.19041.3803
CLRVersion: 4.0.30319.42000
WSManStackVersion: 3.0
PSRemotingProtocolVersion: 2.3
SerializationVersion: 1.1.0.1

About Script : 
Author : Fardin Barashi
Title : Microsoft 365 Endpoints checker
Description : The Office 365 IP Address and URL web service assists users in better identifying and distinguishing Office 365 network traffic, making it easier to assess, configure, and stay updated on changes. This REST-based web service replaces the previously downloadable XML files, which were phased out on October 2, 2018  This script checks after which ip-adresses that M365 Endpoints is needed. 
Documentation is available at http://aka.ms/ipurlws
ipadress m365 ( https://learn.microsoft.com/en-us/microsoft-365/enterprise/urls-and-ip-address-ranges?view=o365-worldwide )

Version : 1
Release day : 2023-12-20
Github Link  : https://github.com/fardinbarashi
News : 

#>

#----------------------------------- Settings ------------------------------------------
# Transcript
$ScriptName = $MyInvocation.MyCommand.Name
$LogFileDate = (Get-Date -Format yyyy/MM/dd/HH.mm.ss)
$TranScriptLogFile = "$PSScriptRoot\Logs\$ScriptName - $LogFileDate.Txt" 
$StartTranscript = Start-Transcript -Path $TranScriptLogFile -Force
Get-Date -Format "yyyy/MM/dd HH:mm:ss"
Write-Host ".. Starting TranScript"

# Error-Settings
$ErrorActionPreference = 'Continue'

# Uri
$clientRequestId = [GUID]::NewGuid().Guid
$uri = "https://endpoints.office.com/endpoints/worldwide?NoIPv6=true&clientRequestId=$clientRequestId"


#----------------------------------- Start Script ------------------------------------------


# Section 1 : Connect to Uri
$Section = "Section 1 : Connect to $uri"
Try
{ # Start Try, $Section
 Get-Date -Format "yyyy/MM/dd HH:mm:ss"
 Write-Host $Section... "0%" -ForegroundColor Yellow
 $EndpointSets = Invoke-RestMethod -Uri ($uri) -Verbose
         If ( $EndpointSets -Eq $Null  ) 
          { # Start If, $EndpointSets is empty or null
            Write-host ""
            Get-Date -Format "yyyy/MM/dd HH:mm:ss"
            Write-Host "ERROR on $Section" -ForegroundColor Red
            Write-Host "$EndpointSets is empty or null" -ForegroundColor Red
            Write-Host "Stopping Transcript and Script!" -ForegroundColor Red
            Stop-Transcript
            Exit
          } # End If, $EndpointSets is empty or null
          Else
          { # Start Else, $EndpointSets is Not empty or null
            Write-host ""
            Get-Date -Format "yyyy/MM/dd HH:mm:ss"
            Write-Host $Section... "100%" -ForegroundColor Green

            # Optimize
            Write-host ""
            Write-host "-----------"
            Write-host ""
            Write-Host "Getting all the Optimize Endpoints" 
            $Optimize = $endpointSets | Where-Object { $_.category -eq "Optimize" }
            $OptimizeIpsv4 = $Optimize.ips | Where-Object { ($_).contains(".") } | Sort-Object -Unique     
            Write-Host $OptimizeIpsv4

            Write-host ""
            Write-Host "Saving all the Optimize Endpoints to a textfile" 
            $optimizeIpsv4 | out-file -FilePath "$PSScriptRoot\Files\Microsoft Optimize Endpoints $LogFileDate.Txt" -Verbose -Encoding utf8 -Force

            # Allow
            Write-host ""
            Write-host "-----------"
            Write-host ""
            Write-Host "Getting all the Allow Endpoints" 
            $Allow = $endpointSets | Where-Object { $_.category -eq "Allow" }
            $AllowIpsv4 = $Allow.ips | Where-Object { ($_).contains(".") } | Sort-Object -Unique
            Write-Host $AllowIpsv4

            Write-host ""
            Write-Host "Saving all the Allow Endpoints to a textfile"            
            $AllowIpsv4 | out-file -FilePath "$PSScriptRoot\Files\Microsoft Allow Endpoints $LogFileDate.Txt" -Verbose -Encoding utf8 -Force

            # Default
            Write-host ""
            Write-host "-----------"
            Write-host ""
            Write-Host "Getting all the Default Endpoints" 
            $Default = $endpointSets | Where-Object { $_.category -eq "Default" }
            $DefaultIpsv4 = $Allow.ips | Where-Object { ($_).contains(".") } | Sort-Object -Unique
            Write-Host $DefaultIpsv4
            
            Write-host ""
            Write-Host "Saving all the Default Endpoints to a textfile"            
            $DefaultIpsv4 | out-file -FilePath "$PSScriptRoot\Files\Microsoft Default Endpoints $LogFileDate.Txt" -Verbose -Encoding utf8 -Force

            Write-host ""
            Write-host "-----------"

          } # End Else, $EndpointSets is Not empty or null

 # Run Query

 Write-Host ""
} # End Try

Catch
{ # Start Catch
Get-Date -Format "yyyy/MM/dd HH:mm:ss"
 Write-Host "ERROR on $Section" -ForegroundColor Red
 # Get-Errors
 Get-Error
 Write-Warning $Error[0]
 Write-Warning $Error[0].CategoryInfo
 Write-Host ""
 Write-Warning $Error[0].InvocationInfo
 Write-Host ""
 Write-Warning $Error[0].Exception
 Write-Host ""
 Write-Warning $Error[0].FullyQualifiedErrorId
 Write-Host ""
 Write-Warning $Error[0].PipelineIterationInfo
 Write-Host ""
 Write-Warning $Error[0].ScriptStackTrace
 Write-Host ""
 Write-Warning $Error[0].TargetObject
 Write-Host ""
 Write-Warning $Error[0].PSMessageDetails
 Write-Host "Stopping Transcript and Script!" -ForegroundColor Red
 Stop-Transcript
 Exit
} # End Catch

#----------------------------------- End Script ------------------------------------------
Stop-Transcript
