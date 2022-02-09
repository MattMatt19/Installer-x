# Installer-x for quick and automatic utilities installation and Windows Updates


Function newLine($nLines) {
    $count = 0
    while ( $count -lt $nLines) {
        write-host "`n"
        $count++
    }
}

#admin req
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $PSCommandArgs" -Verb RunAs
	Exit
}

#system check
if ( [System.Environment]::OSVersion.Version.Build -lt 18363 ) {
    "Warning: you need to update to Windows 10 version 1909 at least ..."
    Start-Sleep -s 5
    Exit
}

Write-Output "Automatic Installation - Personalize it"
 
newLine(3)

#net temporary mapping
$device = "device to find"
$rootPath = "\\ip\path"

try { 
    New-PSDrive -Name $device -PSProvider "FileSystem" -Root $rootPath
} catch {
    "$device already mapped"
    newLine(1)
}

#winget packages (still don't know how to check if they are already installed before the invoke-webReq #damnit)
$DesktopAppPath = ".\DesktopAppInstaller.appxbundle"
$VCLibsAppPath  = ".\VCLibs.appx"

$wingetRelease =Invoke-WebRequest 'https://api.github.com/repos/microsoft/winget-cli/releases/latest' -UseBasicParsing
$wingetVersion = (ConvertFrom-Json $wingetRelease).tag_name
$DesktopAppInstallerUri = "https://github.com/microsoft/winget-cli/releases/download/$wingetVersion/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" 
Invoke-WebRequest -Uri $DesktopAppInstallerUri -UseBasicParsing -OutFile $DesktopAppPath 

addPackages($DesktopAppPath)

$env:PROCESSOR_ARCHITECTURE

#variables for checking if package are already downloaded
$path64 = "Microsoft.VCLibs.x64.14.00.Desktop.appx"
$path86 = "Microsoft.VCLibs.x86.14.00.Desktop.appx"

if ([Environment]::Is64BitOperatingSystem) {
    Invoke-WebRequest -Uri "https://aka.ms/$path64" -OutFile $VCLibsAppPath -UseBasicParsing
}
else {
    Invoke-WebRequest -Uri "https://aka.ms/$path86" -OutFile $VCLibsAppPath -UseBasicParsing
}


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
            newLine(1)
            Write-Host "$packages is updated"
        }
    }
}

addPackages($VCLibsAppPath) 


##Utilities Install (code in progress for office options)
#uni-variables
$installationType
$suppCountBusiness
$menuresponse
$officeType
$suppcountOffice2
$menuresponseO

function setWindowsUpdateChannel {
    do {
        newLine(1)
        Write-Host "do you want to set semi-annual channel and turn off preview releases?"
        $menuresponseB = read-host "(Y/N)"
            Switch ($menuresponseB) {
                "Y" {WindowsUpdateKRMod; $suppCountBusiness = 1 }
                "N" {$suppCountBusiness = 2}
            }
    } 
    until (1..2 -contains $suppCountBusiness) 
}

#sub-menù for Office option
function subOfficeChoice  {
    $suppOfficeType
    do {
        newLine(1)
        Write-Host "which office do you want to install?"
        Write-Host "1. Office 2016 VL 64bit `n2. Office 2016 Home & Business 32bit `n3. Office 2019 std-Professional VL `n4. Office 2019 proPlus retail " 
        $suppOfficeType =Read-Host [inserisci scelta]
   }
    until (1..4 -contains $suppOfficeType)
    return $suppOfficeType
}

#main menu for installation type
do {
    newLine(1)
    Write-Host "choise the type of installation by the customer:"
    Write-Host "1. School `n2. Business `n3. Private"
    $menuresponse = read-host [Inserisci scelta]
    switch ($menuresponse) {
        1 { $installationType = 1 }
        2 { setWindows
        UpdateChannel; $installationType = 2}
        3 { $installationType = 3 }       
    }
}
until (1..3 -contains $menuresponse) 

# main menu for office type
do {
    newLine(1)
    Write-Host "do you want to install Office?"
    $menuresponseO = read-host "(Y/N)"
    Switch ($menuresponseO) {
        "Y" {
            $officeType = subOfficeChoice
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

##semi-annual channel and disable preview build WU (Registry path mod)
function WindowsUpdateKRMod {
    set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name BranchReadinessLevel -Value 32
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate";

    if ( !(Test-Path $registryPath) ) { 
        New-Item -Path $registryPath -Force
    }

    New-ItemProperty -Path $registryPath -Name "ManagePreviewBuilds" -Value 1 -PropertyType DWORD -Force;
    New-ItemProperty -Path $registryPath -Name "ManagePreviewBuildsPolicyValue" -Value 0 -PropertyType DWORD -Force;
}

#Utilities winget
newLine(1)
$listOfUtilities = @("7zip.7zip","Google.Chrome","Oracle.JavaRuntimeEnvironment","Adobe.Acrobat.Reader.64-bit")
$businessUtilities = @("CLechasseur.PathCopyCopy","WinDirStat","Microsoft.dotNetFramework")
$privatesUtilities = @("")
$toInstall

if ($installationType -eq 2) {
    $toInstall = $listOfUtilities + $businessUtilities
}
elseif ($installationType -eq 3 ) {
    $toInstall = $listOfUtilities + $privatesUtilities
}
else {
    $toInstall = $listOfUtilities
}

foreach ($utility in $toInstall) {
    if (winget list --Id $utility) {
        Write-Host "$utility already installed";
    } 
    else {
        winget install -e --id $utility
    }
}

#office installation paths
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