# Installer-x for quick and automatic utilities installation and Windows Updates

# Dot source functions file (evil)
.\functions.ps1

# Check admin req
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $PSCommandArgs" -Verb RunAs
	Exit
}

# Check OS Version
if ( [System.Environment]::OSVersion.Version.Build -lt 18363 ) {
    "Warning: you need to update to Windows 10 version 1909 at least ..."
    Start-Sleep -s 5
    Exit
}

Write-Output "### Automatic Installation - Personalize it"
printNewLine(3)

# net temporary mapping
$device = "device to find"
$rootPath = "\\ip\path"

try { 
    New-PSDrive -Name $device -PSProvider "FileSystem" -Root $rootPath
} catch {
    "$device already mapped"
}

# winget packages
# FIXME: check if packages are already installed before invoking the web request
$DesktopAppPath = ".\DesktopAppInstaller.appxbundle"
$wingetRelease = Invoke-WebRequest 'https://api.github.com/repos/microsoft/winget-cli/releases/latest' -UseBasicParsing
$wingetVersion = (ConvertFrom-Json $wingetRelease).tag_name
$DesktopAppInstallerUri = "https://github.com/microsoft/winget-cli/releases/download/$wingetVersion/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" 
Invoke-WebRequest -Uri $DesktopAppInstallerUri -UseBasicParsing -OutFile $DesktopAppPath 
addPackages($DesktopAppPath)


$VCLibsAppPath  = ".\VCLibs.appx"
$env:PROCESSOR_ARCHITECTURE
if ([Environment]::Is64BitOperatingSystem) {
    Invoke-WebRequest -Uri "https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx" -OutFile $VCLibsAppPath -UseBasicParsing
}
else {
    Invoke-WebRequest -Uri "https://aka.ms/Microsoft.VCLibs.x86.14.00.Desktop.appx" -OutFile $VCLibsAppPath -UseBasicParsing
}
addPackages($VCLibsAppPath) 


## Utilities Install (WIP)
#uni-variables
$deployType
$suppCountBusiness
$menuresponse
$officeType
$suppcountOffice2
$menuresponseO


########### Program Start

#main menu for installation type
enum InstallationType {
    school = 1
    business = 2
    private = 3
}   

do {
    printNewLine(1)
    Write-Host "Choose installation Type :" +
        "1. School `n" +
        "2. Business `n" +
        "3. Private"
    [InstallationType]$deployType = read-host [Inserisci scelta]    
} until (1..3 -contains $deployType)

if ($deployType == [InstallationType]::Business) {
    chooseWindowsUpdateChannel;
}


# main menu for office type
# FIXME: Use enums
do {
    printNewLine(1)
    Write-Host "do you want to install Office?"
    $menuresponseO = read-host "(Y/N)"
    Switch ($menuresponseO) {
        "Y" {
            $officeType = chooseOfficeReleaseName
            switch ($officeType) {
                1 { $officeType = 0 }
                2 { $officeType = 1 }
                3 { $officeType = 2 }
                4 { $officeType = 3 }
            }
            $suppcountOffice = 1
        }
        
        "N" {$suppCountOffice = 0}
    }  
} until (0..1 -contains $suppCountOffice)


#Utilities winget
printNewLine(1)
$utilities = @("7zip.7zip","Google.Chrome","Oracle.JavaRuntimeEnvironment","Adobe.Acrobat.Reader.64-bit")
$businessUtilities = @("CLechasseur.PathCopyCopy","WinDirStat","Microsoft.dotNetFramework")
$privatesUtilities = @("")
$utilities

# Select utilities 
if ($deployType -eq 2) {
    $utilities += $businessUtilities
}
elseif ($deployType -eq 3 ) {
    $utilities += $privatesUtilities
}

foreach ($utility in $utilities) {
    if (winget list --Id $utility) {
        Write-Host "$utility already installed";
    } else {
        winget install -e --id $utility
    }
}

# office installation paths
# FIXME: Static Path specific
if ($suppCountOffice -eq 1) {
    $mainPath = @("\\ip\path")
    $officePath = @(
        "Office 2016\Office_2016_64Bit_STD_VolumeLicensing\setup.exe",
        "Office 2016\Home & Businnes Retail x86 x64\HomeBusinessRetail 2016 x86 x64\setup.exe",
        "Office 2019\OfficeProPlus2019ESD\retail\ProPlus2019RetailItalian1\Setup.exe")
    $officeToInstall = $officePath[$officeType]
    start-process -FilePath "$mainPath\$officeToInstall"
}

# Get windows updates
# FIXME: AutoReboot doesn't work
# FIXME: Declare version strings elsewhere
Find-PackageProvider -Name "NuGet" -AllVersions
Install-PackageProvider -Name "NuGet" -MinimumVersion 2.8.5.201 -Force;

if (Get-Module -ListAvailable -Name PSWindowsUpdate) {
    Write-Host "PSWindowsUpdate Module already installed"
}
else {
    Install-Module PSWindowsUpdate -Force
}

Import-Module PSWindowsUpdate
Get-WindowsUpdate -Install -AcceptAll -RecurseCycle 2 -AutoReboot

###end of the script, ty.

# PR welcome.