#region Functions
## http://poshcode.org/2216

##############################################################################
##
## Send-File
##
## From Windows PowerShell Cookbook (O'Reilly)
## by Lee Holmes (http://www.leeholmes.com/guide)
##
##############################################################################

<#

.SYNOPSIS

Sends a file to a remote session.

.EXAMPLE

PS >$session = New-PsSession leeholmes1c23
PS >Send-File c:\temp\test.exe c:\temp\test.exe $session

#>
function Send-File
{param(
    ## The path on the local computer
    [Parameter(Mandatory = $true)]
    $Source,

    ## The target path on the remote computer
    [Parameter(Mandatory = $true)]
    $Destination,

    ## The session that represents the remote computer
    [Parameter(Mandatory = $true)]
    [System.Management.Automation.Runspaces.PSSession] $Session
)

Set-StrictMode -Version Latest

## Get the source file, and then get its content
$sourcePath = (Resolve-Path $source).Path
$sourceBytes = [IO.File]::ReadAllBytes($sourcePath)
$streamChunks = @()

## Now break it into chunks to stream
Write-Progress -Activity "Sending $Source" -Status "Preparing file"
$streamSize = 1MB
for($position = 0; $position -lt $sourceBytes.Length;
    $position += $streamSize)
{
    $remaining = $sourceBytes.Length - $position
    $remaining = [Math]::Min($remaining, $streamSize)

    $nextChunk = New-Object byte[] $remaining
    [Array]::Copy($sourcebytes, $position, $nextChunk, 0, $remaining)
    $streamChunks += ,$nextChunk
}

$remoteScript = {
    param($destination, $length)

    ## Convert the destination path to a full filesytem path (to support
    ## relative paths)
    $Destination = $executionContext.SessionState.`
        Path.GetUnresolvedProviderPathFromPSPath($Destination)

    ## Create a new array to hold the file content
    $destBytes = New-Object byte[] $length
    $position = 0

    ## Go through the input, and fill in the new array of file content
    foreach($chunk in $input)
    {
        Write-Progress -Activity "Writing $Destination" `
            -Status "Sending file" `
            -PercentComplete ($position / $length * 100)

        [GC]::Collect()
        [Array]::Copy($chunk, 0, $destBytes, $position, $chunk.Length)
        $position += $chunk.Length
    }

    ## Write the content to the new file
    [IO.File]::WriteAllBytes($destination, $destBytes)

    ## Show the result
    Get-Item $destination
    [GC]::Collect()
}

## Stream the chunks into the remote script
$streamChunks | Invoke-Command -Session $session $remoteScript `
    -ArgumentList $destination,$sourceBytes.Length
    }
##end function
#endregion
<# Powershell to remove Visio and install Visio viewer
Steve Ellis 23 Nov 2016
Prerequisites:
- WinRM access to target machines
- these files must exist in c:\temp:
    visioviewer32bit.exe from Microsoft
    config.xml
Assumes that OS is 64-bit but can be changed to 32-bit (see comments)
Assumes Visio 2010 but can be used for 2007 (see comments)
#>

#set to working location of the script and config.xml files
set-location z:\ps1\VisioCleanup
#replace "machinename" with actual target machine name
$session = New-PSSession -ComputerName machinename

#have to use function to send file to target machine as there is no copy function in PS 3 to WinRM session
#assumes existence of c:\temp on target machine
Send-File config.xml c:\temp\config.xml $session
#Visio 2007; comment out above line and uncomment below
#Send-File config2007.xml c:\temp\config2007.xml $session

#copy MS Visio viewer to target
Send-File visioviewer32bit.exe c:\temp\visioviewer32bit.exe $session

Invoke-Command -Session $session -scriptblock {
# stop if Visio is running; if user is running Visio maybe you shouldn't remove it :)
$isVisioRunning = ((Get-Process |where-object {$_.ProcessName -eq "visio"}) | measure).count

if (!$isVisioRunning)
{
# see what Visio apps are running before removal
write-output "Visio apps BEFORE:"
Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -match “Visio”}

#assumes 64-bit OS; if 32-bit OS remove the " (x86)" text below
#Visio 2010
$VisioUninstallString = 'C:\Program Files (x86)\Common Files\Microsoft Shared\Office14\Office Setup Controller\setup.exe'
#Visio 2007 - comment out above string and uncomment below
#$VisioUninstallString = 'C:\Program Files (x86)\Common Files\Microsoft Shared\Office12\Office Setup Controller\setup.exe'

#Visio 2010
$VisioUninstallArgs = '/uninstall visio /config c:\temp\config.xml'
#Visio 2007 - comment out above string and uncomment below
#$VisioUninstallArgs = '/uninstall vispro /config c:\temp\config2007.xml'
start-process $VisioUninstallString -argumentlist $VisioUninstallArgs -wait

#sleep for 15 seconds just in case
start-sleep 15

#install the free viewer
$VisioViewerInstallString = 'C:\temp\visioviewer32bit.exe'
$VisioViewerInstallArgs = '/quiet /norestart /log:c:\temp\visioviewer.txt'
start-process $VisioViewerInstallString -ArgumentList $VisioViewerInstallArgs -wait

#sleep some more to be sure app is installed
start-sleep 15

#confirm apps installed after the process
write-output "Visio apps AFTER:"
Get-WmiObject -Class Win32_Product | sort-object Name | select Name | where { $_.Name -match “Visio”}
}
}
Remove-PSSession $session