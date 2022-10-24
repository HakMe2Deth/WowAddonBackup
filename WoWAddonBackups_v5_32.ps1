## -------------------------------------------------------------------------------------------------
## Parameters (passed from shortcut after install)
## -------------------------------------------------------------------------------------------------

$ParamUseOneDrive = $args[0]

if( ([string]::IsNullOrWhitespace($ParamUseOneDrive)) -or ($ParamUseOneDrive -ne "YesOneDrive" ) ){
 	$AlwaysUseOneDrive="NoOneDrive"
}
Else
{
 	$AlwaysUseOneDrive="YesOneDrive"
}

## -------------------------------------------------------------------------------------------------
## Functions
## -------------------------------------------------------------------------------------------------

function Test-RegistryValue {
	param (
		[parameter(Mandatory=$true)]
		[ValidateNotNullOrEmpty()]$Path,
		[parameter(Mandatory=$true)]
		[ValidateNotNullOrEmpty()]$Value
	)
	try {
		Get-ItemProperty -Path $Path | Select-Object -ExpandProperty $Value -ErrorAction Stop | Out-Null
		return $true
	}
	catch {
		return $false

	}
}

## -------------------------------------------------------------------------------------------------
## Script Info
$ScriptVersion='v5.32'
$ScriptDate='October 2022'
## Script Banner

$PSWinWidth = $Host.UI.RawUI.WindowSize.Width
$BannerLine = "-" * ($PSWinWidth - 1)
$BannerSpace = " " * ($PSWinWidth - 1)

$DescLine1 = "World of Warcraft Addons Backup Script " + $ScriptVersion + " " + $ScriptDate
$DescLine2 = "Developed by HAK - EU-Thunderhorn"
$DescLine3 = "Please visit The Older Gamers on Discord"
$DescLine4 = "https://discord.com/servers/tog-the-older-gamers-313467016985968641"
$DescLine5 = "No restrictions or license requirements for the code needed"
$DescLine6 = "Acknowledgments, Issues and comments appreciated"
$DescLine7 = " "

$infoLine1 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLine1.length/2) ) ) + $DescLine1
$infoLine2 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLine2.length/2) ) ) + $DescLine2
$infoLine3 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLine3.length/2) ) ) + $DescLine3
$infoLine4 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLine4.length/2) ) ) + $DescLine4
$infoLine5 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLine5.length/2) ) ) + $DescLine5
$infoLine6 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLine6.length/2) ) ) + $DescLine6
$infoLine7 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLine7.length/2) ) ) + $DescLine7


## Debug options
## -------------------------------------------------
## - to debug - set following to 1
## - to turn on full debug - set following to 2
## - Default debug off = 0
## -------------------------------------------------

cls
## "Start - $AlwaysUseOneDrive"
## start-sleep -s 5

$debug=0

## add the dialogue box module
Add-Type -AssemblyName PresentationFramework



## -------------------------------------------------------------------------------------------------
$WowInstallDir = ""
if ( Test-RegistryValue -Path 'HKLM:\Software\WOW6432Node\Blizzard Entertainment\World of Warcraft' -Value 'InstallPath') {
	$WowInstallDir = (Get-ItemProperty -path 'HKLM:\Software\WOW6432Node\Blizzard Entertainment\World of Warcraft') | Select-Object -ExpandProperty 'InstallPath'
	if (-not ($debug -eq 0)) {
		"Warcraft Install Folder: $WowInstallDir"
	}
}
else
{
	[System.Windows.MessageBox]::Show('No Warcraft Retail folder found. Please resolve.','WOW Install Error!')
	if (-not ($debug -eq 0)) {
		Write-output "no install"
	}
	break
}

$CurrentScriptPath=Get-Location
$CurrentScriptName=$MyInvocation.MyCommand.Name
$CurrentScript="$CurrentScriptPath" + '\' + "$CurrentScriptName"


$WowInterfaceDir=$WowInstallDir+'Interface'
$WowWTFDir=$WowInstallDir+'WTF'


## "Test to see if folder [$WowInterfaceDir] exists"
if (-not (Test-Path -Path $WowInterfaceDir)) {
	if (-not ($debug -eq 0)) {
		"Path doesn't exist! $WowInterfaceDir"
	}
	break
}

## "Test to see if folder [$WowWTFDir] exists"
if (-not (Test-Path -Path $WowWTFDir)) {
	if (-not ($debug -eq 0)) {
		"Path doesn't exist! $WowWTFDir"
	}
	break
}

$FirstScriptInstall=0

$WowAddonsBKPRootDir=$WowInstallDir+'WowAddonsBackup'
$WowAddonsBKPScriptsDir=$WowAddonsBKPRootDir+'\WowBackupScripts'
$WowAddonsBKPDataDir=$WowAddonsBKPRootDir+'\WowBackupData'

## "Test to see if folder [$WowAddonsBKPRootDir] exists"
if (-not (Test-Path -Path $WowAddonsBKPRootDir)) {
	if (-not ($debug -eq 0)) {
		"Path doesn't exist! $WowAddonsBKPRootDir"
	}
	[System.Windows.MessageBox]::Show('Welcome to the WoW Addons Backup Process. This will allow you to backup your Addons and Addon Config data to a compressed file for wach Day of the Week. It will use 7Zip if installed or the TAR backup in Windows if not. It will also detect if you have OneDrive Personal and give the option to copy to the cloud.','WOW Addon Backup Introduction')
	$msgBoxInput =  [System.Windows.MessageBox]::Show('No Addon Backup Folder. Create folders and copy script?','Script Setup','YesNo','Error')
	switch  ($msgBoxInput) {
		'Yes' {
		## No Script - will create
		$FirstScriptInstall=1
		}
		'No' {
		## No Script - exit
		$FirstScriptInstall=2
		}
	}
}

## "First Install? $FirstScriptInstall"

if ($FirstScriptInstall -eq 2) {
	if (-not ($debug -eq 0)) {
		"Path doesn't exist! $WowAddonsBKPRootDir - no create"
	}
	[System.Windows.MessageBox]::Show('New Install Cancelled. Exiting Script','WOW Addon Backup Not Installed')
	break
}


## Check for Onedrive and override Passed Param
$OneDriveCloudDefault="No"
if (test-path 'env:OneDriveConsumer') {
	$OneDriveCloudDefault="Yes"
}
else 
{
	$AlwaysUseOneDrive="NoOneDrive"
}

if (-not ($debug -eq 0)) {
	write-output "Default Upload to OneDrive: $OneDriveCloudDefault - OneDrive In Shortcut : $AlwaysUseOneDrive"
}



if ($FirstScriptInstall -eq 1) {
	$host.UI.RawUI.ForegroundColor = "White"
	$host.UI.RawUI.BackgroundColor = "Blue"
	cls
	$bannerline
	$DescLineInstall1 = "Script Install Stage"
	$infoLineInstall1 = $BannerSpace.Substring(0, ( [math]::Round($PSWinWidth/2) - [math]::Round($DescLineInstall1.length/2) ) ) + $DescLineInstall1
	$infoLineInstall1
	$bannerline

	if (-not ($debug -eq 0)) {
		"Path doesn't exist! $WowAddonsBKPRootDir - creating"
	}
	$null = new-item -Path "$WowAddonsBKPRootDir" -itemtype Directory

	## "Test to see if folder [$WowAddonsBKPScriptsDir] exists"
	if (-not (Test-Path -Path $WowAddonsBKPScriptsDir)) {
		if (-not ($debug -eq 0)) {
			"Path doesn't exist! $WowAddonsBKPScriptsDir - creating"
		}
		$null = new-item -Path "$WowAddonsBKPScriptsDir" -itemtype Directory
	}
	
	## "Test to see if folder [$WowAddonsBKPDataDir] exists"
	if (-not (Test-Path -Path $WowAddonsBKPDataDir)) {
		if (-not ($debug -eq 0)) {
			"Path doesn't exist! $WowAddonsBKPDataDir - creating"
		}
		$null = new-item -Path "$WowAddonsBKPDataDir" -itemtype Directory
	}

	## file and path
	$WowScriptName="$WowAddonsBKPScriptsDir\" + "$CurrentScriptName"
	if (-not ($debug -eq 0)) {
		"Path doesn't exist! $WowAddonsBKPDataDir - creating"
		"Current Script : $CurrentScript"
		"Current Script Path : $CurrentScriptPath"
		"Current Script Name : $CurrentScriptName"
		"New Script : $WowScriptName"
		start-sleep -s 15
	}
	$null = copy-item "$CurrentScript" -Destination "$WowAddonsBKPScriptsDir"
	
	## -------------------------------------------------------------------------------------------------
	## Creating shortcut
	## -------------------------------------------------------------------------------------------------
	## powershell -Command "& '<PATH_TO_PS1_FILE>' '<ARG_1>' '<ARG_2>' ... '<ARG_N>'"
	
	$WowAddonIconLocation = "$WowInstallDir" + "WOW.exe,0"
	$DesktopPath = [Environment]::GetFolderPath("Desktop")
	$SCArguments = ""
	
	if ($OneDriveCloudDefault -eq "No") {
		$SCArguments = '-ExecutionPolicy Bypass -file "' + $WowScriptName + '" NoOneDrive'
	}
	else
	{
		$msgBoxInput = [System.Windows.MessageBox]::Show('OneDrive Detected. Set Backup Shortcut to Default to Create a OneDrive Copy for Each Day of the Week?','OneDrive Copy in Shortcut?','YesNo','Error')
		switch  ($msgBoxInput) {
		'Yes' {
			## Always Use OneDrive
			$SCArguments = '-ExecutionPolicy Bypass -file "' + $WowScriptName + '" YesOneDrive'
		}
		'No' {
			## Prompt for OneDrive
			$SCArguments = '-ExecutionPolicy Bypass -file "' + $WowScriptName + '" NoOneDrive'
		}
	}
	
	
		$WScriptShell = New-Object -ComObject WScript.Shell
		$TargetEXE = "`"C:\Windows\System32\WindowsPowerShell\v1.0\powershell.exe`""
		$TargetFile = "$WowScriptName"
		$ShortCutFile = "$WowAddonsBKPScriptsDir\WowAddonBackup.lnk"
		$ShortCut = $WScriptShell.CreateShortcut($ShortcutFile)
		$ShortCut.TargetPath = $TargetEXE
		$ShortCut.IconLocation = "$WowAddonIconLocation"
		$ShortCut.Arguments = $SCArguments
		$ShortCut.Save()
	
		$msgBoxInput = [System.Windows.MessageBox]::Show('WOW Addons Backup system created. Copy Shortcut to Desktop?','WOW Addons Backup Script','YesNo','Error')
		switch  ($msgBoxInput) {
		'Yes' {
			## Copy Shortcut to Desktop
			$DesktopShortcut = "$DesktopPath" + "\WowAddonBackup.lnk"
			if (Test-Path $DesktopShortcut ) {
				$null = Remove-Item $DesktopShortcut
			}
			$null = copy-item "$ShortcutFile" -Destination "$DesktopPath"
		}
			'No' {
				## Dont Copy Shortcut
			}
		}
	}
		
	## start-sleep -s 15
	## Running new script
	[System.Windows.MessageBox]::Show('New Install Complete. Running New Shortcut','First Run WOW Addon Backup')
	& "$ShortcutFile"
	break
}

## -------------------------------------------------------------------------------------------------



## -------------------------------------------------
## Individual Filename for multiple computers
## -------------------------------------------------

$wowaddonbackfname=($env:COMPUTERNAME)+"_wowaddonbackup"

## -------------------------------------------------------------------------------------------------
## Set default Zip exe - TAR - Windows 10 1803 onwards

##	3 types of compression:
##		0 - Not defined - exit
##		1 - TAR		- built into W10 1803 onwards
##		2 - 7ZIP	- Needs to be installed
## -------------------------------------------------------------------------------------------------

## $plist= @($env:path -split ";") ; foreach ($i in $plist) {  Get-ChildItem $i -Filter "*tar*" | Select-Object name, directory }
## -------------------------------------------------------------------------------------------------


$ziptype=0

## Debug Switch to force use of internal TAR compression. 1=Use TAR
$UseTAR=0

## Set default to TAR W10 1803
## $Zipexe='tar.exe'
$Zipexe=""
$zipextension='gz'
## checking TAR - Windows 10 1803 onwards

$flist=@()
$plist= @($env:path -split ";") 
foreach ($i in $plist) {  
	$floc = Get-ChildItem $i -Filter "tar.exe" -ErrorAction SilentlyContinue | Select fullname -ExpandProperty fullname
	$flist+=$floc
}

$TarZipexe=$flist[0]
if (-not ($TarZipexe -eq "")) {
	$ziptype=1
	$Zipexe=$TarZipexe
	$zipextension='gz'
	if (-not ($debug -eq 0)) {
		"TAR file default 1 - found : $Zipexe"
		if ($debug -eq 2) {
			"List of TAR files found in PATH:"
			$flist
			"Testing TAR EXE"
			start-process -NoNewWindow -wait -filepath "$TarZipexe" -argumentlist "-?"
		}
	}
}


if (-not ($UseTAR -eq 1)) {

	$flist=@()
	$plist= @($env:programfiles -split ";") 
	foreach ($i in $plist) {  
		$floc = Get-ChildItem $i -recurse -Filter "7z.exe" -ErrorAction SilentlyContinue | Select fullname -ExpandProperty fullname
		$flist+=$floc
	}

	$SevenZipexe=$flist[0]
	if (-not ($SevenZipexe -eq "")) {
		$ziptype=2
		$Zipexe=$SevenZipexe
		$zipextension='7z'
		if (-not ($debug -eq 0)) {
			"7Z file default 2 - found : $Zipexe"
			if ($debug -eq 2) {
				"List of 7Z files found in PATH:"
				$flist
				"NOT Testing 7Z EXE as it has multiple screens - pausing 5 seconds"
##			start-sleep -s 5
##			start-process -NoNewWindow -filepath "$SevenZipexe" -argumentlist "-?"
			}
		}
	}

}
## If no ZIP found, error and exit

if ($ziptype -eq 0) {
	cls
	"Not found TAR or 7Zip exe files. Exiting after 15 seconds"
	start-sleep -s 15
	exit
}

## -------------------------------------------------------------------------------------------------
## Ready for zip process
## -------------------------------------------------------------------------------------------------

if (-not ($debug -eq 0)) {
	"Ziptype = $ziptype - EXE  = $zipexe - Extension = $zipextension"
""
""
	"WOW Game Install Folder   = $WowInstallDir"
	"WOW Mods Interface Folder = $WowInterfaceDir"
	"WOW Mods WTF Folder       = $WowWTFDir"
""
""
	"WOW Addons Backup Root   = $WowAddonsBKPRootDir"
	"WOW Addons Backup Script = $WowAddonsBKPScriptsDir"
	"WOW Addons Backup Data   = $WowAddonsBKPDataDir"
""
""
##	start-sleep -s 5
}




## Get Day of Week as a number
$DOW=[Int] (Get-Date).DayOfWeek
$fname="$WowAddonsBKPDataDir\$wowaddonbackfname$DOW.$zipextension"

if (-not ($debug -eq 0)) {
	"Starting Zip Process - Day $DOW of the week"
}

if (Test-Path $fname ) {
	if (-not ($debug -eq 0)) {
		"file $fname exists so removing (debug wait for 30 seconds)"
		start-sleep -s 30
	}
	$null = Remove-Item $fname
}


## -------------------------------------------------------------------------------------------------
## Zip
## -------------------------------------------------------------------------------------------------

## changing screen colours - only if not in debug mode
if ($debug -eq 0) {
	$host.UI.RawUI.ForegroundColor = "White"
	$host.UI.RawUI.BackgroundColor = "DarkGreen"
	cls
}

$Bannerline
$infoLine1
$infoLine2
$infoLine3
$infoLine4
$infoLine5
$infoLine6
$infoLine7
$Bannerline


## -------------------------------------------------
## Set to TAR
## -------------------------------------------------


if ($ziptype -eq 1) {
	$Bannerline
	"Backing up WOW WTF amd Interface files from :"
	$WowInstallDir
	"Saving in :"
	$fname
	"Using built in TAR compression"
	$Bannerline
	
	$passexe= '"' + $zipexe + '"' + " "
	$passargs= " --exclude=**/*.bak -czf " + '"' + $fname + '"' + " " + '"' + $WowWTFDir + '"' + " " + '"' + $WowInterfaceDir + '"'
##	$passexe
##	$passargs
	"Windows TAR Backup in progress. Please be patient"
	$Bannerline
	start-process -NoNewWindow -wait -filepath "$zipexe" -argumentlist $passargs
		if (-not ($debug -eq 0)) {
			"TAR process complete - pausing 30"
			start-sleep -s 30
		}
}


## -------------------------------------------------
## Set to 7Zip
## -------------------------------------------------

if ($ziptype -eq 2) {
	$Bannerline
	"Backing up WOW WTF amd Interface files from :"
	$WowInstallDir
	"Saving in :"
	$fname
	"Using installed 7Zip Compression"
	$Bannerline

	$passexe= '"' + $zipexe + '"' + " "
	$passargs= "a -bso0 -bsp1 " + '"' + $fname + '"' + " " + '"' + $WowWTFDir + '"' + " " + '"' + $WowInterfaceDir + '"' + " -xr!*.bak"

##	$passexe
##	$passargs
 	start-process -NoNewWindow -wait -filepath $passexe -argumentlist $passargs
}

## Final code for OneDrive
##	"Always - $AlwaysUseOneDrive"
##	start-sleep -s 10
if (($OneDriveCloudDefault -eq "No") -or ($AlwaysUseOneDrive -ne "YesOneDrive")) {
	[System.Windows.MessageBox]::Show('Addon Backup Complete','Complete')
	## start-sleep -s 10
}
else 
{
	## "Always - $AlwaysUseOneDrive"
	## start-sleep -s 10
	$DoOneDriveCopy=0
	if ($AlwaysUseOneDrive -eq "YesOneDrive") {
		$DoOneDriveCopy=1
	}
	else
	{
		$DoOneDriveCopy=0
	}
##	"Always - $AlwaysUseOneDrive"
##	"DoOD Copy - $DoOneDriveCopy"
##	start-sleep -s 10
	
	if ($DoOneDriveCopy -eq 1) {
		$OneDriveBackupDocsFolder="$env:OneDriveConsumer\Documents\WOWAddonsBackup"
		if (-not (Test-Path -Path $OneDriveBackupDocsFolder)) {
			if (-not ($debug -eq 0)) {
				"Path doesn't exist! $OneDriveBackupDocsFolder"
				start-sleep -s 10
			}
			$null = new-item -Path "$OneDriveBackupDocsFolder" -itemtype Directory	
		}
		$ODfname="$OneDriveBackupDocsFolder\$wowaddonbackfname$DOW.$zipextension"
##		$OneDriveBackupDocsFolder
##		$ODfname
			
		if (Test-Path $ODfname ) {
			if (-not ($debug -eq 0)) {
				"file $ODfname exists so removing (debug wait for 10 seconds)"
				start-sleep -s 10
			}
			$null = Remove-Item $ODfname
		}
		$null = copy-item -force "$fname" -Destination "$OneDriveBackupDocsFolder"
	}
	
	[System.Windows.MessageBox]::Show('Addon Backup Complete','Complete')
	cls
##	start-sleep -s 10
}

## exit