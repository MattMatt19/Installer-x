Function printNewLine($nLines) {
    $count = 0
    while ( $count -lt $nLines) {
        write-host "`n"
        $count++
    }
}

function chooseWindowsUpdateChannel {
    do {
        printNewLine(1)
        Write-Host "Do you want to use Windows Updates semi-annual channel and turn off Preview releases?"
        $menuresponseB = read-host "(Y/N)"
            Switch ($menuresponseB) {
                "Y" {SetWindowsUpdateChannelKRMod; $suppCountBusiness = 1 }
                "N" {$suppCountBusiness = 2}
            }
    } 
    until (1..2 -contains $suppCountBusiness) 
}

## semi-annual channel and disable preview build WU (Registry path mod)
function SetWindowsUpdateChannelKRMod {
    set-ItemProperty -Path HKLM:\SOFTWARE\Microsoft\WindowsUpdate\UX\Settings -Name BranchReadinessLevel -Value 32
    $registryPath = "HKLM:\SOFTWARE\Policies\Microsoft\Windows\WindowsUpdate";

    if ( !(Test-Path $registryPath) ) { 
        New-Item -Path $registryPath -Force
    }

    New-ItemProperty -Path $registryPath -Name "ManagePreviewBuilds" -Value 1 -PropertyType DWORD -Force;
    New-ItemProperty -Path $registryPath -Name "ManagePreviewBuildsPolicyValue" -Value 0 -PropertyType DWORD -Force;
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

