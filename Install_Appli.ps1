[console]::CursorVisible = $false
$Host.UI.RawUI.WindowTitle = "  Interiale - Poste Travail"
$Host.UI.RawUI.BackgroundColor = 0

$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Clear-Host

if (!$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Write-Host "`n Lancer le script en mode admin`n" -f red
    " Appuyez sur Entrée pour continuer ou CTRL+C pour quitter"
    $Host.UI.ReadLine()
    Clear-Host
    return
}

$prog = @(
    #Installation winget
    @{name = "Dell Command";alt="Dell"; wgt = "Dell.CommandUpdate" }
    @{name = "7 Zip";alt="7 Zip"; wgt = "7zip.7zip"; wpck = "7-zip" }
    @{name = "Google Chrome";alt="Chrome"; wgt = "Google.Chrome"; wpck = "Google Chrome" }
    @{name = "Adobe Acrobat Reader";alt="Acrobat Reader"; wgt = "Adobe.Acrobat.Reader.64-bit"; wpck = "Adobe Acrobat" }
    @{name = "Microsoft Edge";alt="Edge"; wgt = "Microsoft.Edge"; wpck = "Microsoft Edge" }
    @{name = "Microsoft Office";alt="Office"; wgt = "Microsoft.Office"; wpck = "Office" }
    @{name = "Microsoft Teams";alt="Teams"; wgt = "Microsoft.Teams"; wpck = "" }
    #Installation locale
    @{name = "Dell Encryption";alt="Dell Encryption"; fname = "DDPE_v11.1-force.bat"; wgt = ""; wpck = "Dell Encryption" }
    @{name = "Cyberwatch Agent";alt="Cyberwatch Agent"; fname = "cyberwatch_agent-x64-4.11-SRVCYB01PRD.msi"; wgt = ""; wpck = "CyberwatchAgent" }
    @{name = "Trellix";alt="McAfee"; fname = "McAfee_agent.5.7.6.251.exe"; wgt = ""; wpck = "Trellix" }
    @{name = "Multiplateforme";alt="Multiplateforme"; fname = "multiplatform_201812041125_x64.exe"; wgt = "" }
    @{name = "Forcepoint";alt="Forcepoint"; fname = "instal_force_proxy_forcepoint.bat"; wgt = ""; wpck = "FORCEPOINT" }
    @{name = "Global Protect";alt="Global Protect"; fname = "GlobalProtect64-6.0.0.msi"; wgt = ""; wpck = "GlobalProtect" }
)

Function CConsole {
    for ($x = 0; $x -le 80; $x++) {
        for ($y = 0; $y -le 25; $y++) {
            [Console]::SetCursorPosition($x, $y)
            Write-Host " " -BackgroundColor Black
        }
        Start-Sleep -Milliseconds 5
    }
    [Console]::SetCursorPosition(4, 4)
}

CConsole
# Installation Winget
$hasPackageManager = Get-AppPackage -name 'Microsoft.DesktopAppInstaller'
if (!$hasPackageManager -or [version]$hasPackageManager.Version -lt [version]"1.10.0.0") {
    "Installing winget Dependencies"
    Add-AppxPackage -Path 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'
    $releases_url = 'https://api.github.com/repos/microsoft/winget-cli/releases/latest'
    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $releases = Invoke-RestMethod -uri $releases_url
    $latestRelease = $releases.assets | Where-Object { $_.browser_download_url.EndsWith('msixbundle') } | Select-Object -First 1

    Write-Host "`n`n      Installation winget apartir $($latestRelease.browser_download_url)" -f yellow
    Add-AppxPackage -Path $latestRelease.browser_download_url
    return
}
else {
    Write-Host "`n`n  Winget: installé" -f green
    Start-Sleep -Seconds 2
}
CConsole

# Copie du dossier pstools
Write-Host "`n   Copie du dossier pstools vers le C:" -f white

if (!(Test-Path -path "C:\pstools")) {New-Item "C:\pstools" -Type Directory}
Copy-Item -Path 'C:\RESTANT_A_FAIRE\pstools\*' -Destination 'C:\pstools\' -Recurse | Out-Null
Start-sleep -Seconds 3
CConsole
if(Test-Path -Path C:\pstools\mdt_flag.txt -PathType Leaf) {
    Write-Host "`n   Copie du dossier pstools vers le C:" -f green
    Start-Sleep -Seconds 3
} else {
    Write-Host "`n   Copie du dossier pstools vers le C:" -f red
    Write-Host "`n Verifier si le dossier existe ou il est sur le bon endroit`n" -f red
    " Appuyez sur Entrée pour continuer ou CTRL+C pour quitter"
    $Host.UI.ReadLine()
    return
}
CConsole

# Installation d'application
Write-Host "    Installation" -f cyan

foreach ($p in $prog) {
    Write-Host "        " $p.name -f yellow
    if ($p.wgt -ne "") {
        # installation winget
        winget install --exact $p.wgt --source "winget" --silent
        if ($p.name -like "*Microsoft Office*") {
            Wait-Process -name "*Microsoft Office*"
        }
        elseif ($p.fname -like "*.msi*") {
            # installation fichiers msi
            $msi = "C:\pstools\" + $p.fname
            $arg = "/i $msi"
            Start-Process "msiexec.exe" -ArgumentList $arg -Wait -nonewwindow
        }
        else {
            # installation fichiers bat et exe
            Start-Process -FilePath $p.fname -Wait -WorkingDirectory "C:\pstools\" -Verb RunAs -WindowStyle Normal
        }
    }
}
CConsole


# Verificer la clé OneDrive
Set-itemproperty -Path 'HKLM:\HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\OneDrive' -name DisableFileSyncNGSC 0
$prop = Get-ItemProperty -Path 'HKLM:\HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\OneDrive' -name DisableFileSyncNGSC | Select-Object DisableFileSyncNGSC
if($prop -ne 0) {
    Write-Host "        OneDrive:       1"
} elses {
    Write-Host "        OneDrive:       0"
}
Start-Sleep -Seconds 2
# Layout
if (!(Test-Path -path "C:\temp")) {New-Item "C:\temp" -Type Directory}
Copy-Item "C:\RESTANT_A_FAIRE\pstools\layout2023.xml" -Destination "C:\temp" -Recurse | Out-Null
Import-StartLayout -LayoutPath C:\Temp\layout2023.xml -MountPath C:\

# MAJ Windows
Write-Host "        MAJ Windows"
Install-Module -Name PSWindowsUpdate -Force
Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot

# Verifier l'installation
Write-Host " `n moment...`n" -f Yellow
$pack = get-package | Select-Object name -unique -Wait
CConsole
Foreach($prg in $prog) {
    if($prg.wpck -ne "") {
         Write-host " " $prg.name\$prg.alt ":   " -f yellow -NoNewline
        for($i = 0; $i -le $pack.Count - 1; $i++) {
             if(($pack[$i].name -like "*"+$prg.wpck+"*") -or ($pack[$i].name -like "*"+$prg.alt+"*")) {
                Write-Host "Installe " -f green
                break
            } elseif ($i -eq $pack.Count - 1) {
                write-host "Pas installe" -f red
            }
        }
    }
    Start-Sleep -Milliseconds 50
}
CConsole

# Supprimer dossier RESTANT_A_FAIRE
Remove-Item -Path "C:\RESTANT_A_FAIRE" -Force -Recurse
Write-Host "        Installation terminé" -f cyan
Write-Host "`n`n        Entree pour dire: Au revoir"
$Host.UI.ReadLine()
Exit-PSSession
exit
