###############################################################################################################
# Language     :  PowerShell 4.0
# Filename     :  Import-WifiProfiles.ps1
# Author       :  stra
# Description  :  Import WiFi profiles asynchronously using RunspacePool
###############################################################################################################

<#
    .SYNOPSIS
    Imports WiFi profiles asynchronously using RunspacePool.
    
    .DESCRIPTION
    This script retrieves WiFi profiles from the current system using RunspacePool to execute the retrieval
    in parallel, enhancing performance especially when dealing with multiple profiles.
    
    .EXAMPLE
    .\Import-WifiProfiles.ps1
    
    Runs the script to import WiFi profiles asynchronously.
    
    .LINK
    https://github.com/RangO1972/PowerShell-WiFi-Profiles-Exporter
#>

[CmdletBinding()]
param(
    # Number of concurrent threads
    [Parameter(
        Position = 0,
        HelpMessage = 'Maximum number of threads at the same time (Default=10)'
    )]
    [int]$Threads = 10
)

Begin{
    # Get the hostname (computer name)
    $hostname = $env:COMPUTERNAME

    # Get the current date in the format yyyymmdd
    $date = Get-Date -Format "yyyyMMdd"

    # Get the current directory for the output file
    $outputDir = Get-Location

    # Create the output file names
    $outputXmlFile = Join-Path -Path $outputDir -ChildPath "${hostname}_wifi_${date}.xml"
    $outputCsvFile = Join-Path -Path $outputDir -ChildPath "${hostname}_wifi_${date}.csv"

    # Retrieve all saved WiFi profiles
    $wifiProfiles = netsh wlan show profiles | Select-String -Pattern ":(.*)$" | ForEach-Object { $_.Matches.Groups[1].Value.Trim() }

    # Display the list of SSIDs found
    Write-Output "Found the following SSIDs:"
    $wifiProfiles
}
process{
    # Scriptblock to export profile to XML
    [System.Management.Automation.ScriptBlock]$ExportProfileScriptBlock = {
        param (
            $wifiProfile,
            $outputDir
        )

        # Export WiFi profile to XML
        netsh wlan export profile name="$wifiProfile" folder="$outputDir" key=clear | Out-Null

        # Return the profile name for tracking purposes
        $wifiProfile
    }

    # Create RunspacePool
    Write-Verbose "Setting up RunspacePool..."
    $RunspacePool = [System.Management.Automation.Runspaces.RunspaceFactory]::CreateRunspacePool(1, $Threads)
    $RunspacePool.Open()
    [System.Collections.ArrayList]$Jobs = @()

    Write-Verbose "Setting up Jobs..."

    # Start jobs for each WiFi profile
    foreach ($wifiProfile in $wifiProfiles) {
        # Hashtable to pass parameters to the script block
        $ScriptParams = @{
            wifiProfile = $wifiProfile
            outputDir = $outputDir
        }

        # Create a new PowerShell job
        $Job = [System.Management.Automation.PowerShell]::Create().AddScript($ExportProfileScriptBlock).AddParameters($ScriptParams)
        $Job.RunspacePool = $RunspacePool

        # Begin asynchronously invoking the job
        $JobObj = [pscustomobject] @{
            ProfileName = $wifiProfile
            Pipe = $Job
            Result = $Job.BeginInvoke()
        }

        # Add job object to collection
        [void]$Jobs.Add($JobObj)
    }

    Write-Verbose "Waiting for jobs to complete & starting to process results..."

    # Wait for all jobs to complete
    $Jobs | ForEach-Object {
        $_.Pipe.EndInvoke($_.Result)
        $_.Pipe.Dispose()
    }

    Write-Verbose "All jobs completed."


}

End{
    # Combine all XML files into one
    $combinedXml = New-Object System.Xml.XmlDocument
    $root = $combinedXml.CreateElement("WLANProfiles")
    $combinedXml.AppendChild($root)

    # Check for the existence of XML files and wait until they are available
    $retryCount = 0
    do {
        $xmlFiles = Get-ChildItem -Path $outputDir -Filter "*.xml"
        if ($xmlFiles.Count -eq $wifiProfiles.Count) {
            break
        } else {
            Write-Verbose "Waiting for XML files to be fully written..."
            Start-Sleep -Seconds 1
            $retryCount++
        }
    } while ($retryCount -lt 30)  # Retry for up to 30 seconds

    # Import each XML file into the combined XML document
    foreach ($xmlFile in $xmlFiles) {
        $profileXml = [xml](Get-Content $xmlFile.FullName)
        $importedNode = $combinedXml.ImportNode($profileXml.WLANProfile, $true)
        $root.AppendChild($importedNode)
        Remove-Item $xmlFile.FullName
    }

    # Save the combined XML file
    $combinedXml.Save($outputXmlFile)

    # Filter out profiles without passwords and extract SSID and password
    $ssidPasswordList = foreach ($profile in $combinedXml.WLANProfiles.WLANProfile) {
        $password = $profile.MSM.security.sharedKey.keyMaterial
        if (![string]::IsNullOrWhiteSpace($password)) {
            [PSCustomObject]@{
                SSID = $profile.SSIDConfig.SSID.name
                Password = $password
            }
        }
    }

    # Export SSID and password to CSV
    $ssidPasswordList | Export-Csv -Path $outputCsvFile -NoTypeInformation

    # Confirmation message
    Write-Output "WiFi profiles retrieval completed. Check the files $outputXmlFile and $outputCsvFile for WiFi credentials."
}
