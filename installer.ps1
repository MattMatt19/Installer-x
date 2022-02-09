# Installer-x for quick and automatic utilities installation and Windows Updates


Function printNewLine($nLines) {
    $count = 0
    while ( $count -lt $nLines) {
        write-host "`n"
        $count++
    }
}

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

Write-Output "Automatic Installation - Personalize it"
 
printNewLine(3)

#net temporary mapping
$device = "device to find"
$rootPath = "\\ip\path"

try { 
    New-PSDrive -Name $device -PSProvider "FileSystem" -Root $rootPath
} catch {
    "$device already mapped"
}

#winget packages (still don't know how to check if they are already installed before the invoke-webReq #damnit)
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


# Add Package (or PackageX??) listed in $packages, checking if packages are already installed
Function addPackages($packages) {
    $FileVersion = (Get-ItemProperty -Path $packages ).VersionInfo.ProductVersion
    $HighestInstalledVersion = Get-AppxPackage -Name Microsoft.VCLibs* | Sort-Object -Property Version | Select-Object -ExpandProperty Version -Last 1
    
    if ($HighestInstalledVersion -eq "") {
        Add-AppPackage -path $packages
        Write-Host "$packages installed"
    } else {
        if ($HighestInstalledVersion -lt $FileVersion ) {
            Add-AppxPackage -Path $packages
        } else {
            printNewLine(1)
            Write-Host "$packages is updated"
        }
    }
}


##Utilities Install (code in progress for office options)
#uni-variables
$installationType
$suppCountBusiness
$menuresponse
$officeType
$suppcountOffice2
$menuresponseO

function chooseWindowsUpdateChannel {
    do {
        printNewLine(1)
        Write-Host "Do you want to use Windows Updates semi-annual channel and turn off Preview releases?"
        $menuresponseB = read-host "(Y/N)"
            Switch ($menuresponseB) {
                "Y" {SetWindowsUpdateChannel; $suppCountBusiness = 1 }
                "N" {$suppCountBusiness = 2}
            }
    } 
    until (1..2 -contains $suppCountBusiness) 
}

#sub-menù for Office option
function chooseOfficeReleaseName {
    $officeRelease
    do {
        printNewLine(1)
        Write-Host "Select the Office release to be installed:" +
            "1. Office 2016 VL 64bit `n" +
            "2. Office 2016 Home & Business 32bit `n" +
            "3. Office 2019 std-Professional VL `n" +
            "4. Office 2019 proPlus retail" 
        $officeRelease = Read-Host [inserisci scelta]
    } until (1..4 -contains $officeRelease)
    
    return $officeRelease
}

########### Program Start

#main menu for installation type
do {
    printNewLine(1)
    Write-Host "Choose installation Type :" +
        "1. School `n" +
        "2. Business `n" +
        "3. Private"
    $installationType = read-host [Inserisci scelta]    
} until (1..3 -contains $installationType) 

if ($installationType == 2) {
    chooseWindowsUpdateChannel;
}


# main menu for office type
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
}
until (0..1 -contains $suppCountOffice)

## semi-annual channel and disable preview build WU (Registry path mod)
function SetWindowsUpdateChannel {
    set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name BranchReadinessLevel -Value 32
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate";

    if ( !(Test-Path $registryPath) ) { 
        New-Item -Path $registryPath -Force
    }

    New-ItemProperty -Path $registryPath -Name "ManagePreviewBuilds" -Value 1 -PropertyType DWORD -Force;
    New-ItemProperty -Path $registryPath -Name "ManagePreviewBuildsPolicyValue" -Value 0 -PropertyType DWORD -Force;
}

#Utilities winget
printNewLine(1)
$basicUtilities = @("7zip.7zip","Google.Chrome","Oracle.JavaRuntimeEnvironment","Adobe.Acrobat.Reader.64-bit")
$businessUtilities = @("CLechasseur.PathCopyCopy","WinDirStat","Microsoft.dotNetFramework")
$privatesUtilities = @("")

$utilities

# Select utilities 
$utilities = $basicUtilities
if ($installationType -eq 2) {
    $utilities = $basicUtilities + $businessUtilities
}
elseif ($installationType -eq 3 ) {
    $utilities = $basicUtilities + $privatesUtilities
}

foreach ($utility in $utilities) {
    if (winget list --Id $utility) {
        Write-Host "$utility already installed";
    } 
    else {
        winget install -e --id $utility
    }
}

# office installation paths
if ($suppCountOffice -eq 1) {
    $mainPath = @("\\ip\path")
    $officePath = @("Office 2016\Office_2016_64Bit_STD_VolumeLicensing\setup.exe","Office 2016\Home & Businnes Retail x86 x64\HomeBusinessRetail 2016 x86 x64\setup.exe","Office 2019\OfficeProPlus2019ESD\retail\ProPlus2019RetailItalian1\Setup.exe")
    $officeToInstall = $officePath[$officeType]
    start-process -FilePath "$mainPath\$officeToInstall"
}

#windows update (autoreboot may not work, #damnit)
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

#end of the script, ty.

#to suggest any change to the script just contact MattMatt19. :)