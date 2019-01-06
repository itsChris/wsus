###############################################################################################
# Solvia GmbH
# 
# Function: Force Installing Updates
# 
# Date        	ver    Name                 Remarks 
#*************************************************************************
# 24.08.2015	1.0    Christian Muggli     1st implementation
# 12.03.2016    1.1    Christian Muggli     added WSUS Support
# 17.11.2017    1.2    Christian Casutt     Added Logging
###############################################################################################

#Common Variables
[string] $ScriptName            = $MyInvocation.MyCommand.Name
[string] $LogDir                = "$ENV:SystemDrive\Solvia\Logs"
[string] $LogFilePath           = [string]::Format("{0}\{1}_{2}.log", $LogDir, "$(get-date -format `"yyyyMMdd_hhmmsstt`")",$ScriptName.Replace(".ps1",""))
[string] $UpdateSearchFilter    = "IsInstalled=0 and Type='Software' and IsHidden=0"
# Functions

<# 
.Synopsis 
   Write-Log writes a message to a specified log file with the current time stamp. 
.DESCRIPTION 
   The Write-Log function is designed to add logging capability to other scripts. 
   In addition to writing output and/or verbose you can write to a log file for 
   later debugging. 
.NOTES 
   Created by: Jason Wasser @wasserja 
   Modified: 11/24/2015 09:30:19 AM   
 
   Changelog: 
    * Code simplification and clarification - thanks to @juneb_get_help 
    * Added documentation. 
    * Renamed LogPath parameter to Path to keep it standard - thanks to @JeffHicks 
    * Revised the Force switch to work as it should - thanks to @JeffHicks 
 
   To Do: 
    * Add error handling if trying to create a log file in a inaccessible location. 
    * Add ability to write $Message to $Verbose or $Error pipelines to eliminate 
      duplicates. 
.PARAMETER Message 
   Message is the content that you wish to add to the log file.  
.PARAMETER Path 
   The path to the log file to which you would like to write. By default the function will  
   create the path and file if it does not exist.  
.PARAMETER Level 
   Specify the criticality of the log information being written to the log (i.e. Error, Warning, Informational) 
.PARAMETER NoClobber 
   Use NoClobber if you do not wish to overwrite an existing file. 
.EXAMPLE 
   Write-Log -Message 'Log message'  
   Writes the message to c:\Logs\PowerShellLog.log. 
.EXAMPLE 
   Write-Log -Message 'Restarting Server.' -Path c:\Logs\Scriptoutput.log 
   Writes the content to the specified log file and creates the path and file specified.  
.EXAMPLE 
   Write-Log -Message 'Folder does not exist.' -Path c:\Logs\Script.log -Level Error 
   Writes the message to the specified log file as an error message, and writes the message to the error pipeline. 
.LINK 
   https://gallery.technet.microsoft.com/scriptcenter/Write-Log-PowerShell-999c32d0 
#> 
function Write-Log 
{ 
    [CmdletBinding()] 
    Param 
    ( 
        [Parameter(Mandatory=$true, ValueFromPipelineByPropertyName=$true)] 
        [ValidateNotNullOrEmpty()] 
        [Alias("LogContent")] 
        [string]$Message, 
 
        [Parameter(Mandatory=$false)] 
        [Alias('LogPath')] 
        [string]$Path=$LogFilePath, 
         
        [Parameter(Mandatory=$false)] 
        [ValidateSet("Error","Warn","Info")] 
        [string]$Level="Info", 
         
        [Parameter(Mandatory=$false)] 
        [switch]$NoClobber 
    ) 
 
    Begin 
    { 
        # Set VerbosePreference to Continue so that verbose messages are displayed. 
        $VerbosePreference = 'Continue' 
    } 
    Process 
    { 
        try
        {
            # If the file already exists and NoClobber was specified, do not write to the log. 
            if ((Test-Path $Path) -AND $NoClobber) { 
                Write-Error "Log file $Path already exists, and you specified NoClobber. Either delete the file or specify a different name." 
                Return 
                } 
 
            # If attempting to write to a log file in a folder/path that doesn't exist create the file including the path. 
            elseif (!(Test-Path $Path)) { 
                Write-Verbose "Creating $Path." 
                $NewLogFile = New-Item $Path -Force -ItemType File 
                } 
 
            else { 
                # Nothing to see here yet. 
                } 
 
            # Format Date for our Log File 
            $FormattedDate = Get-Date -Format "yyyy-MM-dd HH:mm:ss" 
 
            # Write message to error, warning, or verbose pipeline and specify $LevelText 
            switch ($Level) { 
                'Error' { 
                    Write-Error $Message 
                    $LevelText = 'ERROR:' 
                    } 
                'Warn' { 
                    Write-Warning $Message 
                    $LevelText = 'WARNING:' 
                    } 
                'Info' { 
                    Write-Verbose $Message 
                    $LevelText = 'INFO:' 
                    } 
                } 
         
            # Write log entry to $Path 
            "$FormattedDate $LevelText $Message" | Out-File -FilePath $Path -Append -Width 1024
        }
        catch{
            $ErrorMessage = $_.Exception.Message
        }
    } 
    End 
    { 
    } 
}

Function CheckForAdmin(){
    $currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
    $res = $currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
    return $res
}

Function SendNotification($subject, $body){
    try
    {
        Write-Log -Message "SendNotification entered.." -Path $LogFilePath -Level Info

        $Subject = $Subject
        $Body = $Body
        $SMTPClient = New-Object Net.Mail.SmtpClient($SmtpServer) 
        $SMTPClient.Credentials = New-Object System.Net.NetworkCredential($SmtpUsername, $SmtpPassword); 
        $SMTPClient.Send($EmailFrom, $EmailTo, $Subject, $Body) # SendAsync? prevent slowing down login times? don't know, maybe in v2.0

        Write-Log -Message "Mail sent!" -Path $LogFilePath -Level Info
    }
    catch
    {
        Write-Log -Message ($_.Exception.Message) -Path $LogFilePath -Level Error    
    }
}

Write-Log -Message 'Starting..' 

Write-Log ([string]::Format("Check for admin privileges..."))

if(-not (CheckForAdmin)){
    Write-Log -Message ("Run Script as admin - will quit now!") -Level Warn
    exit
}

$UpdateSession = New-Object -ComObject 'Microsoft.Update.Session'
$UpdateSession.ClientApplicationID = 'Solvia Windows Update Installer'
$UpdateSearcher = $UpdateSession.CreateUpdateSearcher()
 
Write-Log -Message 'Searching for updates...' 
Write-Log -Message ([string]::Format("SearchFilter is: {0}", $UpdateSearchFilter))

$SearchResult = $UpdateSearcher.Search($UpdateSearchFilter)
 
if ($SearchResult.Updates.Count -ne 0) {
    Write-Log -Message  ([string]::Format("There are: {0} applicable updates on the machine", $SearchResult.Updates.Count))
}
else {
    Write-Log -Message 'There are no applicable updates' 
    break
}
Write-Log -Message 'Creating a collection of updates to download:'
$UpdatesToDownload = New-Object -ComObject 'Microsoft.Update.UpdateColl'
foreach ($Update in $SearchResult.Updates) {
    [bool]$addThisUpdate = $false
    if ($Update.InstallationBehavior.CanRequestUserInput) {
        Write-Log -Message "> Skipping: $($Update.Title) because it requires user input"
    }
    else {
        if (!($Update.EulaAccepted)) {
            Write-Log -Message "> Note: $($Update.Title) has a license agreement that must be accepted:"
            $Update.EulaText
            $strInput = Read-Host 'Do you want to accept this license agreement? (Y/N)'
            if ($strInput.ToLower() -eq 'y') {
                $Update.AcceptEula()
                [bool]$addThisUpdate = $true
            }
            else {
                Write-Log -Message "> Skipping: $($Update.Title) because the license agreement was declined"
            }
        }
        else {
            [bool]$addThisUpdate = $true
        }
    }
    if ([bool]$addThisUpdate) {
        Write-Log -Message "Adding: $($Update.Title)"
        $UpdatesToDownload.Add($Update) |Out-Null
    }
}
 
if ($UpdatesToDownload.Count -eq 0) {
    Write-Log -Message 'All applicable updates were skipped.'
    break
}

Write-Log -Message 'Downloading updates...'
$Downloader = $UpdateSession.CreateUpdateDownloader()
$Downloader.Updates = $UpdatesToDownload
$Downloader.Download()
 
$UpdatesToInstall = New-Object -ComObject 'Microsoft.Update.UpdateColl'
 
[bool]$rebootMayBeRequired = $false
Write-Log -Message 'Successfully downloaded updates'
 
foreach ($Update in $SearchResult.Updates) {
    if ($Update.IsDownloaded) {
        Write-Log -Message "> $($Update.Title)"
        $UpdatesToInstall.Add($Update)
 
        if ($Update.InstallationBehavior.RebootBehavior -gt 0) {
            [bool]$rebootMayBeRequired = $true
        }
    }
}
 
if ($UpdatesToInstall.Count -eq 0) {
    Write-Log -Message 'No updates were succsesfully downloaded'
}
 
if ($rebootMayBeRequired) {
    Write-Log -Message 'These updates may require a reboot'
}
 
Write-Log -Message 'Installing updates...'

$Installer = $UpdateSession.CreateUpdateInstaller()
$Installer.Updates = $UpdatesToInstall
$InstallationResult = $Installer.Install()

Write-Log -Message "Installation Result: $($InstallationResult.ResultCode)"
Write-Log -Message "Reboot Required: $($InstallationResult.RebootRequired)"
Write-Log -Message 'Listing of updates installed and individual installation results'

for ($i = 0; $i -lt $UpdatesToInstall.Count; $i++) {
    Write-Log -Message "> $($Update.Title) : $($InstallationResult.GetUpdateResult($i).ResultCode)"
}

Write-Log -Message 'Done -> Restart might be required!'