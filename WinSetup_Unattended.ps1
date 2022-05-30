<# 
MANUAL INSTALLATION
#Use this script to perform first-time setup on a Windows 10 (maybe 7) machine with Powershell (Non-Admin)

Clear-Host
Write-Host -ForegroundColor Yellow -BackgroundColor DarkBlue 
"Setting working Directory to $Home"
Set-Location $HOME

function ElevTest {
    $RunningElevated = [Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544'
    
    if ($RunningElevated -eq $False) {Write-Host -ForegroundColor Red "NOTICE: Running with Standard user rights!"}
    else{ Write-Host -ForegroundColor Red -BackgroundColor Red "NOTICE: Running Elevated!"}
    
}

function WarningMsg {
    $Warning= Write-Host -ForegroundColor Cyan "Error! Try again"
    return($Warning)
}
function Browsers_MakeFolder {
    $Test = Test-Path $Home\Downloads\Browsers\Browsers_OLD
    if ($Test -eq $true) {
        Remove-Item $home\Downloads\Browsers_OLD
        
    }
    New-Item -ItemType Directory $Home\Downloads\Browsers
}
function Browsers_DeleteFolder {
    return{Remove-Item $HOME\Downloads\Browsers -Recurse -Force
    Remove-Item $Home\Downloads\Browsers_OLD -Recurse -Force
    Write-Host -ForegroundColor DarkRed "Deleted Browsers and Browsers_OLD located in $Home\Downloads"}
}
function Browsers_Install{
    while ($true) {
        
    $Prompt = Read-Host -Prompt "Would you like to install the downloaded browsers?"
    $BrowserDownload = "$Home\Downloads\Browsers"
    $BrowserInstalls =  "$BrowserDownload\BraveSetup.exe",                        
                        "$BrowserDownload\ChromeSetup.exe",
                        "$BrowserDownload\FirefoxSetup.exe",
                        "$BrowserDownload\OperaSetup.exe",
                        "$BrowserDownload\NewEdgeSetup.exe"
                        
    if($prompt -like "y*")  {$BrowserInstalls | ForEach-Object {if (Test-Path $_) {Start-Process -Wait $_}
    }
    }
    elseif ($Prompt -like "n*") { Write-Output "Continuing..."}
    }
        break
    else {
        WarningMsg
        continue
    }
}
function Browsers_Download {
    
    Write-Host -ForegroundColor Green '
    Now, we will select and download a new browser'
    $DownloadChrome =   "https://dl.google.com/chrome/install/375.126/chrome_installer.exe"
    $DownloadFirefox =  "https://download.mozilla.org/?product=firefox-latest-ssl&os=win64&lang=en-US"
    $DownloadBrave =    "https://laptop-updates.brave.com/latest/winx64"
    $DownloadNewEdge =  "https://go.microsoft.com/fwlink/?linkid=2109047&Channel=Stable&language=en&consent=1"
    $DownloadOpera =    "https://www.opera.com/computer/thanks?ni=stable&os=windows"
    $DownloadLocation = "$HOME\Downloads\Browsers" 

    TestBrowserFolder

while ($True) {
    Clear-Host
    Write-Host -ForegroundColor Green '
    1. Google Chrome
    2. Mozilla Firefox
    3. Brave
    4. Microsoft Edge
    5. Opera [Not recommended]'
    $BrowserSelection = Read-Host -Prompt '
    Type your number selection or type Done/Skip to bypass'
    
    if      ($BrowserSelection -like '1*')      {Invoke-WebRequest $DownloadChrome  -OutFile    $DownloadLocation\ChromeSetup.exe }
    elseif  ($BrowserSelection -like '2*')      {Invoke-WebRequest $DownloadFirefox -OutFile    $DownloadLocation\FirefoxSetup.exe }
    elseif  ($BrowserSelection -like '3*')      {Invoke-WebRequest $DownloadBrave   -OutFile    $DownloadLocation\BraveSetup.exe }
    elseif  ($BrowserSelection -like '4*')      {Invoke-WebRequest $DownloadNewEdge -OutFile    $DownloadLocation\NewEdgeSetup.exe }
    elseif  ($BrowserSelection -like '5*')      {Invoke-WebRequest $DownloadOpera   -OutFile    $DownloadLocation\OperaSetup.exe
                                                Write-Host 'NOTE: This brower is not recommended due to the insecurity of the owners/developers'}
    elseif ($BrowserSelection -like 'skip')     {Write-Output 'Skipping :(' Start-Sleep -Milliseconds 350
        break}
    elseif ($BrowserSelection -like 'done') {Browsers_Install}
    else {WarningMsg
        Continue
    }    
}


function Browswers_DeleteFolder{
    while ($true) {        
    
        $DeleteFolder = Read-Host -Prompt "Would you like to delete the Browsers folder created in $Home\Downloads ?"
            if ($DeleteFolder -like 'y*') {Write-Output "- - - Deleting folder - - -"
                                            Start-Sleep -Seconds 1.2
                                            Remove-Item $HOME\Downloads\Browsers -Recurse -Force
                                            Remove-Item $Home\Downloads\Browsers_OLD -Recurse -Force
                                            break
                                            }
    elseif ($DeleteFolder -like 'n*') {Write-Output "Continuing..."; Move-Item $Home\Downloads\Browsers_OLD $Home\Downloads
        break
    }
    else {
        WarningMsg
        continue 
    }
}
            
}
if ($OpenDownloads -like 'y*') {
    Browsers_Install
    break
    }

elseif ($OpenDownloads -like 'n*') { Write-Output "Keep in mind your installer files are stored in ($Home\Downloads\Browsers)"
    Write-Output 'Continuing...'
break
}
else {
    WarningMsg
Continue}
}

function TestBrowserFolder {
    $TestPath = Test-Path $Home\Downloads\Browsers
    if ($TestPath -eq $False)       { Browsers_MakeFolder > $null; Write-Host -ForegroundColor Cyan "Browsers folder does not exist - creating new one in $Home\Downloads"
    }
    elseif ($TestPath -eq $True) {Rename-Item $Home\Downloads\Browsers Browsers_OLD -Force

                                    Write-Host -ForegroundColor Red "Browsers folder already exists in $Home\Downloads...Renamed to Browsers_OLD!"
                                    Browsers_MakeFolder
    }   
}
function ModifyDarkMode {
    while ($true) {

        $ChangeDarkMode = Read-Host -Prompt 'Would you like to Enable Dark Mode [Y/N/Skip]?'
            if ($ChangeDarkMode -like 'y*')  { Write-Output "yes" | reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ /v SystemUsesLightTheme /t REG_DWORD /D 0
            Write-Output "yes" | reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ /v AppsUseLightTheme /t REG_DWORD /D 0
    }
            elseif ($ChangeDarkMode -like 'n*')   {Write-Output "yes" | reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ /v SystemUsesLightTheme /t REG_DWORD /D 1 > $null
            Write-Output "yes" | reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ /v AppsUseLightTheme /t REG_DWORD /D 1    
    }
            elseif ($ChangeDarkMode -like 'skip') {Write-Output 'Skipping :('; Start-Sleep -Milliseconds 500
                break
    }
            else {
                WarningMsg
                Continue
    }
    break
}
    Write-Host 'Dark Mode has been updated!'
}

function InstallOffice {
    
$DownloadOffice       = "https://account.microsoft.com/services/microsoft365/install"
$DownloadLibreOffice  = "https://www.libreoffice.org/donate/dl/win-x86_64/7.1.6/en-US/LibreOffice_7.1.6_Win_x64.msi"
$DownloadOnlyOffice   = "https://download.onlyoffice.com/install/desktop/editors/windows/distrib/onlyoffice/DesktopEditors_x64.exe?_ga=2.83887002.2046324471.1631196325-1839179190.1631046844"
$DownloadOpenOffice   = "https://sourceforge.net/projects/openofficeorg.mirror/files/4.1.10/binaries/en-US/Apache_OpenOffice_4.1.10_Win_x86_install_en-US.exe/download"

Write-Host "Would you like to download MS Office?"
$InstallMSOffice = Read-Host -Prompt "Note: This will require an active Microsoft/Office Account with License (Subcription)"

while ($True) {

    if      ($InstallMSOffice -like 'y*')           { Write-Host 'Opening Office Download Page...'
                                                        Start-Process $DownloadOffice
                                                    break}
    elseif  ($InstallMSOffice -like 'n*')           { $FreeOffice = Read-Host -Prompt "Would you like to download a free Office alternative?"
    break
    }
    elseif  ($FreeOffice -like "y*"){
        Write-Output $DownloadOnlyOffice
        Write-Output $DownloadLibreOffice
        Write-Output $DownloadOnlyOffice
        Write-Output $DownloadOpenOffice
    }
    else {
        WarningMsg
        Continue
    }  
 Break
}
}

function TZChange {
    
    while ($true) {
        
    $TimeZone = Read-Host -Prompt "Please enter your time zone [Central Standard Time | Eastern Standard Time | Pacific Standard Time]"
        if ($TimeZone -like 'C*')     {Set-TimeZone -Name "Central Standard Time" 
        break
    }
    elseif ($TimeZone -like 'E*')     {Set-Timezone -Name "Eastern Standard Time" 
        break
    }
    elseif ($TimeZone -like 'P*')     {Set-Timezone -Name "Pacific Standard Time" 
        break
    }
    else                            {Write-Output WarningMsg
        Continue
    }
        
}
}
function RenameComputer {
    $Elevated = [Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544'
        if ($Elevated -eq $false)    {Start-Process ms-settings:about}
        elseif (ElevTest -like $true) {
                $RenamePC = Read-Host -Prompt "Please enter your computer name"
                Rename-Computer -Name $RenamePC
            }
}
function WindowsUpdate {

    while ($true) {
    if ($WindowsUpdate -like 'y*')      {Start-Process ms-settings:windowsupdate-action        
        break
        }
    elseif ($WindowsUpdate -like 'n*')  {Write-Output "Continuing..."
        break
        }
    else {
        WarningMsg
        Continue
        }
    }
}

$CurrentDateTime = Get-Date -Format G
Write-Host -ForegroundColor Yellow -BackgroundColor DarkBlue "

WELCOME to the Windows Setup Script!
------------------------------------
Copyright - September, 2021"
Write-Host -ForegroundColor Green -BackgroundColor DarkBlue "
Use this script for setting up Windows 10 on a new device!"

$CurrentDateTime

                             
ElevTest

Write-Host "------------------------------------------"

TZChange

ModifyDarkMode

Browsers_Download

Browsers_Install

InstallOffice

RenameComputer

WindowsUpdate
 AUTOMATED INSTALLATION
 Get Logged on user: (Get-WmiObject -Class Win32_ComputerSystem).Username - Shows COMPUTER\Logged on User
#>
# Powershell script to install applications and perform general Windows 10 Setup Tasks
# Add a feature to gather the following information:
# CsName
# CsNumberOfProcessors
# CsNumbersOfLogicalProcessors
# CsProcessors
# CsUserName
# TimeZone
# OsTotalVisibleMemorySize
# WindowsInstallDateFromRegistry
# WindowsProductName
Clear-Host
$host.UI.RawUI.WindowTitle="Windows Setup"
#$Date = Get-Date -Format "yyyy-MM-dd HHmm"
$LogFile = ".\Windows.Setup.log"
$LogTime = Get-Date -Format "MM-dd-yyyy HH:mm:ss CST"
$ScriptStart = Get-Date -Format "HH:mm:ss tt"
while($true)
{
    if
        (
            (Test-path "$LogFile") -like $true
        )
        {
            Write-Output "Windows Setup Script`n==================== $LogTime ====================" | Tee-Object  -Append $LogFile
            break
        }
    elseif 
        (
            (
            Test-path "$LogFile"
            ) -like $false
        )
        {
        New-Item -Name $LogFile | Out-Null
        continue
        }
}


$NewComputerName = Read-Host -Prompt "Please enter your new computer name"
"Computer name: $Env:COMPUTERNAME" | Tee-Object -Append $LogFile

if
    (
        (
            [Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544'
        ) -like "True"
    )
        {
        Write-Output "Running as ADMIN" | Tee-Object -Append $LogFile
        }
else
    {
        Write-Output "Running with STANDARD user privliges`n" | Tee-Object -Append $LogFile
    }
Set-Timezone -Name "Central Standard Time"
Write-Output "Set Timezone to Central Standard Time" | Tee-Object -Append $LogFile
Write-Output "Enabling Dark Mode for Apps and System" | Tee-Object -Append $LogFile

# Registry Changes

Write-Output "Enabling dark mode for 'System'....Status: $(reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ /v SystemUsesLightTheme  /t REG_DWORD /D 0 /f)" | Tee-Object -Append $LogFile 
Write-Output "Enabling dark mode for 'Apps'....Status: $(reg add HKCU\Software\Microsoft\Windows\CurrentVersion\Themes\Personalize\ /v AppsUseLightTheme     /t REG_DWORD /D 0 /f)" | Tee-Object -Append $LogFile
Write-Output "Removing 'Cortana' button from Taskbar....Status: $(reg add HKCU\SOFTWARE\MICROSOFT\WINDOWS\CURRENTVERSION\EXPLORER\ADVANCED /v   ShowCortanaButton     /t REG_DWORD /D 0 /f)" | Tee-Object -Append $LogFile
Read-Host -Prompt "The following applications will be installed: $((Import-CSV .\Apps.csv).'Software Name') -- Press ENTER to continue"
Write-Output "Downloading app installers" | Tee-Object -Append $LogFile
Import-Csv .\apps.csv | ForEach-Object {
    Write-Output "Downloading $($_.'Software Name') $($_.Homepage)" | Tee-Object -Append $LogFile
    Invoke-WebRequest -Uri $_.'Direct-Download Link' -OutFile $_.'Installer' | Tee-Object -Append $LogFile
}
Write-Output "Installing Downloaded Browsers!" | Tee-Object $LogFile
Write-Host -ForegroundColor Yellow "NOTE: To install with Admin rights, click 'Yes' on the following UAC Prompts. Otherwise, click 'No' to install with Standard rights"
Import-Csv .\apps.csv | ForEach-Object {
    Write-Output "Downloading $($_.'Software Name') to --> .\$($_.'Installer') / ($($_.'Homepage'))"
    Invoke-WebRequest -Uri $($_.'Direct-Download Link') -OutFile $($_.'Installer')
}
if
    (
        (
            [Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544'
        ) -like "False"
    )
        {
        Write-Host -ForegroundColor Yellow "Enter your computer name on the next window..."
        Start-Sleep -Seconds 1
        Start-Process ms-settings:about -Wait
        Read-Host -Prompt "Press ENTER to continue"
        }
elseif
    (
        (
            [Security.Principal.WindowsIdentity]::GetCurrent().Groups -contains 'S-1-5-32-544'
        ) -like "True"
    )
    {
    Rename-Computer -NewName ($NewComputerName)
    $NewComputerName = $null
    }

$ScriptEnd = Get-Date -Format "HH:mm:ss tt"
$ScriptDuration = New-TimeSpan -Start $ScriptStart -End $ScriptEnd
Write-Output "Started at $($ScriptStart)`nEnded at $($ScriptEnd)`nScript took $($ScriptDuration.Hours) hours, $($ScriptDuration.Minutes) minutes, and $($ScriptDuration.Seconds) seconds" | Tee-Object -Append $LogFile