##Installer-x for quick and automatic utilities installation and Windows Updates

#admin req
If (!([Security.Principal.WindowsPrincipal][Security.Principal.WindowsIdentity]::GetCurrent()).IsInRole([Security.Principal.WindowsBuiltInRole]"Administrator")) {
	Start-Process powershell.exe "-NoProfile -ExecutionPolicy Bypass -File `"$PSCommandPath`" $PSCommandArgs" -Verb RunAs
	Exit
}

#system check
if([System.Environment]::OSVersion.Version.Build -lt 18363) {"ATTENZIONE: Per usare lo script e' necessario aggiornare almeno a Windows 10 versione 1909 ..."; Start-Sleep -s 5; Exit}

Write-Output "Installazione automatizzata sistema M2"

#net temporary mapping
try { 
    New-PSDrive -Name "ReadyNAS 104 [NASGETGEAR]" -PSProvider "FileSystem" -Root "\\172.25.10.6\iso\Iso Microsoft"
} catch {
    "rete già mappata"
}

#winget packages (still don't know how to check if they are already installed #damnit)
$output=".\DesktopAppInstaller.appxbundle"
$output2=".\VCLibs.appx"
$json=Invoke-WebRequest 'https://api.github.com/repos/microsoft/winget-cli/releases/latest' -UseBasicParsing
$psobj = ConvertFrom-Json $json
$version=$psobj.tag_name
Invoke-WebRequest -Uri "https://github.com/microsoft/winget-cli/releases/download/$version/Microsoft.DesktopAppInstaller_8wekyb3d8bbwe.msixbundle" -OutFile $output -UseBasicParsing
$env:PROCESSOR_ARCHITECTURE

#variables for checking if package are already downloaded
$path64 = "Microsoft.VCLibs.x64.14.00.Desktop.appx"
$path86 = "Microsoft.VCLibs.x86.14.00.Desktop.appx"

if ([Environment]::Is64BitOperatingSystem) {
    Invoke-WebRequest -Uri "https://aka.ms/$path64" -OutFile $output2 -UseBasicParsing
}
else {
    Invoke-WebRequest -Uri "https://aka.ms/$path86" -OutFile $output2 -UseBasicParsing
}

#check if packages are already installed
Function getPackagesVersion ($outputx) {
    $FileVersion = (Get-ItemProperty -Path $outputx ).VersionInfo.ProductVersion
    $HighestInstalledVersion = Get-AppxPackage -Name Microsoft.VCLibs* |
        Sort-Object -Property Version |
        Select-Object -ExpandProperty Version -Last 1
    
    if ($HighestInstalledVersion -eq "") {
        Add-AppPackage -path $outputx
        Write-Host "$outputx installed"
    } else {
        if ($HighestInstalledVersion -lt $FileVersion ) {
            Add-AppxPackage -Path $outputx
        } else {
            Write-Host "$outputx is updated `n"
        }
    }
}

getPackagesVersion ($output2)
getPackagesVersion ($output)

##Utilities Install (code in progress for office options)
#uni-variables
$installationType
$suppCountBusiness
$menuresponse
$officeType
$suppcountOffice
$menuresponseO

#RK choise
function businessOption {
    do {

    Write-Host "Vuoi Impostare i criteri di gruppo per Windows Update a canale semestrale e disattivare l'installazione degli aggiornamenti in preview? (N.B. selezionando <Y> verranno cambiati i valori delle chiavi di registro necessarie)"
    $menuresponseB = read-host "(Y/N)"
        Switch ($menuresponseB) {
            "Y" {WindowsUpdateKRMod; $suppCountBusiness = 1 }
            "N" {$suppCountBusiness = 2}
        }
    } 
        until (1..2 -contains $suppCountBusiness) 
}

#sub-men� for Office option
function subOfficeChoise {
    do {
        Write-Host "quale Office vuoi installare?"
       Write-Host "1. Office 2016 VL 64bit `n2. Office 2016 Home & Business 32bit `n3. Office 2019 std-Professional VL `n4. Office 2019 proPlus retail " 
        $menuresponseSO =Read-Host [inserisci scelta]
       switch ($menuresponseSO) {
           1 { $officeType = 1 }
           2 { $officeType = 2 }
           3 { $officeType = 3 }
           4 { $officeType = 4 }
       }
   }
   until (0..4 -contains $officeType)
}

#main men� for installation type
do {
    Write-Host "scegliere il tipo di installazione in base al cliente:"
    Write-Host "1. Scuola `n2. Azienda `n3. Privato"
    $menuresponse = read-host [Inserisci scelta]
    switch ($menuresponse) {
        1 { $installationType = 1 }
        2 { businessOption; $installationType = 2}
        3 { $installationType = 3 }       
    }
}
    until (1..3 -contains $menuresponse) 

# main men� for office type
do {
    Write-Host "vuoi installare anche un pacchetto office?"
    $menuresponseO = read-host "(Y/N)"
    Switch ($menuresponseO) {
        "Y" { subOfficeChoise; $suppcountOffice = 1}
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

#Utilitieswinget
$listOfUtilities = @("7zip.7zip","Google.Chrome","Oracle.JavaRuntimeEnvironment","Adobe.Acrobat.Reader.64-bit")
$otionalUtilities = @("CLechasseur.PathCopyCopy","WinDirStat","Microsoft.dotNetFramework")
$toInstall

if ($installationType -eq 2)
{
    $toInstall = $listOfUtilities + $otionalUtilities
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
$mainPath =  @("\\172.25.10.6\iso\Iso Microsoft")
$officePath = @("Office 2016\Office_2016_64Bit_STD_VolumeLicensing\setup.exe2","Office 2016\Home & Businnes Retail x86 x64\HomeBusinessRetail 2016 x86 x64\setup.exe","Office 2019\OfficeProPlus2019ESD\retail\ProPlus2019RetailItalian1\Setup.exe")
$type = 0
$officeToInstall
Foreach ($type in $officePath) {
    if ($officePath[$type] -eq $officeType) {
        $officeToInstall = $type
    }
    else {
        $type++
    }
}
start-process -FilePath "$mainpath\$officeToInstall" -Wait 

#windows update (autoreboot may not work, damn it)
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

#to suggest any change to the script just contact Matteo Cannoletta. :)