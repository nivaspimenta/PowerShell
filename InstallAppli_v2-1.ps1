#region Install & Verification d'Applis
#region Tete
$host.ui.RawUI.WindowTitle = "Preparation de Poste"
#endregion

#region Variables
# Logo
$logo = @(
@{a = "                                                                         "}
@{a = "                  # #    # ##### ###### #####  #   ##   #      ###### "}
@{a = "                  # ##   #   #   #      #    # #  #  #  #      #      "}
@{a = "                  # # #  #   #   #####  #    # # #    # #      #####  "}
@{a = "                  # #  # #   #   #      #####  # ###### #      #      "}
@{a = "                  # #   ##   #   #      #   #  # #    # #      #      "}
@{a = "                  # #    #   #   ###### #    # # #    # ###### ###### "}
@{a = "                                                                  v2.1"}
@{a = ""}
@{a = "                                PREPARATION DE POSTE"}
 )

#Tous les programmes qu'on doit installer
$prog = @(
    #Installation winget
    @{name="Dell Command";wgt="Dell.CommandUpdate"}
    @{name="7 Zip";wgt="7zip.7zip";wpck="7-zip"}
    @{name="Google Chrome";wgt="Google.Chrome";wpck="Google Chrome"}
    @{name="Adobe Acrobat Reader";wgt="Adobe.Acrobat.Reader.64-bit";wpck="Adobe Acrobat"}
    @{name="Citrix Workspace";wgt="Citrix.Workspace";wpck="Citrix Workspace"}
    @{name="Microsoft Edge";wgt="Microsoft.Edge";wpck="Microsoft Edge"}
    @{name="Microsoft Office";wgt="Microsoft.Office";wpck="Office"}
    @{name="Microsoft Teams";wgt="Microsoft.Teams";wpck=""}
    #Installation locale
    @{name="Dell Encryption";fname="DDPE_v11.1-force.bat";wgt="";wpck="Dell Encryption"}
    @{name="Cyberwatch Agent";fname="cyberwatch_agent-x64-4.11-SRVCYB01PRD.msi";wgt="";wpck="CyberwatchAgent"}
    @{name="McAfee";fname="McAfee_agent.5.7.6.251.exe";wgt="";wpck="McAfee Agent"}
    @{name="Multiplateforme";fname="multiplatform_201812041125_x64.exe";wgt=""}
    #@{name="UEM Agent";fname="UEMAgent.bat";wgt="";wpck="Matrix42 UEM Agent"}
    @{name="Forcepoint";fname="instal_proxy_forcepoint _gpo-pc.bat";wgt="";wpck="FORCEPOINT"}
    #@{name="Unistall Visual c++ 2010-2014";wgt="";fname="uninstall visual c++.bat"}
    @{name="Global Protect";fname="GlobalProtect64-6.0.0.msi";wgt="";wpck="GlobalProtect"}
)

#Output d'installation
$output = @()

#installation
[int] $_vins

if (!(Test-Path -path "C:\RESTANT_A_FAIRE\inst.int" -PathType Leaf)) {
    $_vins = 0
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
} else {
    $_vins = Get-Content C:\RESTANT_A_FAIRE\inst.int
}


#endregion

#region Functions
Function Logo {
    clear-host
Foreach($lg in $logo) {
    Write-Host $lg.a -f yellow
}
}

Function Installation {
#Verification admin
$currentPrincipal = New-Object Security.Principal.WindowsPrincipal([Security.Principal.WindowsIdentity]::GetCurrent())
$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)
Clear-Host
Logo
if($_vins -eq 0) {

if(!$currentPrincipal.IsInRole([Security.Principal.WindowsBuiltInRole]::Administrator)) {
    Logo
    Write-Host "`n Lancer le script en mode admin`n" -f red
    Read-Host -Prompt " Appuyez sur Entrée pour continuer ou CTRL+C pour quitter"
    return
}
$_vins += 1
$_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
}

if($_vins -eq 1) {
#Installation Winget
Logo
$hasPackageManager = Get-AppPackage -name 'Microsoft.DesktopAppInstaller'
if (!$hasPackageManager -or [version]$hasPackageManager.Version -lt [version]"1.10.0.0") {
    "Installing winget Dependencies"
    Add-AppxPackage -Path 'https://aka.ms/Microsoft.VCLibs.x64.14.00.Desktop.appx'

    $releases_url = 'https://api.github.com/repos/microsoft/winget-cli/releases/latest'

    [Net.ServicePointManager]::SecurityProtocol = [Net.SecurityProtocolType]::Tls12
    $releases = Invoke-RestMethod -uri $releases_url
    $latestRelease = $releases.assets | Where { $_.browser_download_url.EndsWith('msixbundle') } | Select -First 1

    Write-Host "`n`n      Installation winget apartir $($latestRelease.browser_download_url)" -f yellow
    Add-AppxPackage -Path $latestRelease.browser_download_url
    return
}
else {
    Write-Host "`n`n  Winget: installé" -f green
}
$_vins += 1
$_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
}
 #region Installation et MAJs

 if($_vins -eq 2) {
    Logo
    Write-Host "`n   Copie du dossier pstools vers le C:" -f white
    if (!(Test-Path -path "C:\pstools")) {New-Item "C:\pstools" -Type Directory}
    Copy-Item -Path 'C:\RESTANT_A_FAIRE\pstools\*' -Destination 'C:\pstools\' -Recurse | Out-Null
    Start-sleep -Seconds 3

#Flag
    if(Test-Path -Path C:\pstools\mdt_flag.txt -PathType Leaf) {
        Logo
        Write-Host "`n   Copie du dossier pstools vers le C:" -f green
    } else {
        Logo
        Write-Host "`n   Copie du dossier pstools vers le C:" -f red
        Start-Sleep -Seconds 5
        return
    }
    $_vins += 1
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
}
Start-Sleep -Seconds 3

if($_vins -eq 3) {
#region Verification domain
    Logo
    Write-Host "`n Poste:" -f cyan
    $np = Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select name
    Write-Host "`n Nom du Poste: " -NoNewline; Write-Host $np.name -f yellow
    $domain = Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select Domain
    if($domain.Domain -like '*interiale.intra*') {
        Write-Host " Domain: " -NoNewline; Write-Host $domain.Domain -f green
        $output += "Domain: " + $domain.Domain
    } else {
        Write-Host " Domain: " -NoNewline; Write-Host $domain.Domain -f red        $output += "Domain: " + $domain.Domain        return    }
    $_vins += 1
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
}
    Start-Sleep -Seconds 5
    #endregion

#region Installation des aplications
    ##################################################################################################################################
    if($_vins -eq 4) {
    for($i = 0; $i -le $prog.Count - 1; $i++) {
        try {
        Logo
            Write-Host "`n------- Installation des applications" -f white
            Write-Host " `n" $prog[$i].name -f yellow

            if($prog[$i].wgt -ne "") {
                #installation winget
                winget install --exact $prog[$i].wgt --source "winget" --silent
                if($prog[$i].name -like "*Microsoft Office*") {
                    Wait-Process -name "*Microsoft Office*"
                }
            } elseif ($prog[$i].fname -like "*.msi*") {
                #installation fichiers msi
                $msi = "C:\pstools\" + $prog[$i].fname
                $arg = "/i $msi"
                Start-Process "msiexec.exe" -ArgumentList $arg -Wait -nonewwindow
            } else {
                #installation fichiers bat et exe
                Start-Process -FilePath $prog[$i].fname -Wait -WorkingDirectory "C:\pstools\" -Verb RunAs -WindowStyle Normal
            }
            $output += $prog[$i].name + "- Installe"
        } catch {
            $output += $prg.name + "- Pas possible d'installer"
            Write-Host "   Pas Installe" -f red
        }
        Start-Sleep -Seconds 3
    }
    $_vins += 1
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
    }
    ##################################################################################################################################
   
    #endregion
   

    #MAJ Microsoft Store
    if($_vins -eq 5) {
    Write-Host "`n------- Applications - MAJ Microsoft Store" -f white
    winget upgrade --all
    $_vins += 1
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
    Start-Sleep -Seconds 3
    }
    
    if($_vins -eq 6) {
    #Appliquer le Layout
    Write-Host "------- Layout" -f white
    if (!(Test-Path -path "C:\temp")) {New-Item "C:\temp" -Type Directory}
    Copy-Item "C:\RESTANT_A_FAIRE\pstools\layout_proplusv2.xml" -Destination "C:\temp" -Recurse | Out-Null
    Import-StartLayout -LayoutPath C:\Temp\layout_proplusv2.xml -MountPath C:\
    $_vins += 1
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
    Start-Sleep -Seconds 3
    }
        
    
    
    if($_vins -eq 6) {
    #Desactiver\Activer OneDrive
    Write-Host "------- OneDrive" -f white
    Set-itemproperty -Path 'HKLM:\HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\OneDrive' -name DisableFileSyncNGSC 0
    Start-Sleep -Seconds 3
    $_vins += 1
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
    }

    #endregion

#region Installation MAJ Windows
    if($_vins -eq 7) {
    Write-Host "------- MAJ Windows Update`n" -f white
    Find-Module -Name PSWindowsUpdate | Install-Module -Confirm
    Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot
    Start-Sleep -Seconds 5
    $_vins += 1
    $_vins | Out-file C:\RESTANT_A_FAIRE\inst.int
    }
    #endregion

}

Function Verification {
    Logo
    #Chercher les programmes pour
    Write-Host " `n Chargement des programmes...`n" -f Yellow
    $pack = get-package | select name -unique -Wait
    Logo
    Write-Host "`n------- Verification des applications`n" -f white
    Foreach($prg in $prog) {
    if($prg.wpck -ne "") {
            Write-host " " $prg.name -f yellow
            for($i = 0; $i -le $pack.Count - 1; $i++) {
                if($pack[$i].name -like "*"+$prg.wpck+"*") {
                    Write-Host " > Installé " -f green
                    $output += $prg.name + " - Installé"
                    break
                } elseif ($i -eq $pack.Count - 1) {
                    write-host " > Pas installé" -f red
                    $output += $prg.name + " - Pas Installé"
                }
            }
        }
         Start-Sleep -Seconds 5
    }

}

Function FichierLog {
    Logo
    
    if($output -ne "") {
        $output | Out-File "C:\RESTANT_A_FAIRE\Log.txt"
    }
     
    Write-Host "`n ****************************************************************" -f red
    Write-Host "            Fichier Log creé - C:\RESTANT_A_FAIRE\Log.txt" -f yellow
    Write-Host " ****************************************************************`n`n" -f red
    Write-Host "        Redemarrage en 20 seconds" -f yellow
    $host.ui.RawUI.BackgroundColor = 5
    Start-sleep -Seconds 20
    Restart-Computer
}
#endregion

#region Appel des funtions
Installation
Verification
FichierLog
#endregion

#endregion

<#

NOTE:



#>