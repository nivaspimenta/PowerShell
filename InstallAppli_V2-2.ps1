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

#region Variables
$logo = @(
    @{t = "" }
    @{t = "     #### ##    ## ######## ######## ########  ####    ###    ##       ######## "; cor = "white" }
    @{t = "      ##  ###   ##    ##    ##       ##     ##  ##    ## ##   ##       ##       "; cor = "white" }
    @{t = "      ##  ####  ##    ##    ##       ##     ##  ##   ##   ##  ##       ##       "; cor = "white" }
    @{t = "      ##  ## ## ##    ##    ######   ########   ##  ##     ## ##       ######   "; cor = "white" }
    @{t = "      ##  ##  ####    ##    ##       ##   ##    ##  ######### ##       ##       "; cor = "white" }
    @{t = "      ##  ##   ###    ##    ##       ##    ##   ##  ##     ## ##       ##       "; cor = "Magenta" }
    @{t = "     #### ##    ##    ##    ######## ##     ## #### ##     ## ######## ######## "; cor = "Magenta" }
    @{t = "                                                     ----------------------------" }
    @{t = "                                                      Preparation de poste - v2.2"; cor = "yellow" }
    @{t = "                                                     ----------------------------" }
    @{t = "" }
)

$bye = @(
    @{t = ".########..##....##.########....########..##....##.########"; cor = "Magenta" }
    @{t = ".##.....##..##..##..##..........##.....##..##..##..##......"; cor = "Magenta" }
    @{t = ".##.....##...####...##..........##.....##...####...##......"; cor = "white" } 
    @{t = ".########.....##....######......########.....##....######.."; cor = "white" } 
    @{t = ".##.....##....##....##..........##.....##....##....##......"; cor = "white" } 
    @{t = ".##.....##....##....##..........##.....##....##....##......"; cor = "white" } 
    @{t = ".########.....##....########....########.....##....########"; cor = "white" }
)

# formulaire menu
$form = @(
    @{t = "Domain:        "; vlr = ""; act = "edt"; dta = "interiale.intra" }
    @{t = "Nom            "; vlr = ""; act = "edt"; dta = "" }
    @{t = "Log            "; vlr = ""; act = "boo"; dta = "1" }
    @{t = "Installation   "; vlr = "ins"; act = "sub"; dta = "" }
    @{t = "Verifier       "; vlr = "ver"; act="sub";dta="" }
    @{t = "Puff           "; vlr = "mga"; act ="sub";dta=""}
    @{t = "Sortir         "; vlr = "srt"; act = "sub"; dta = "" }
)

$fver = @(
    @{t= "Oui";vlr="ver";act="sub";dta=""}
    @{t= "Non";vlr="menu";act="sub";dta=""}
)

# programmes
$prog = @(
    #Installation winget
    @{name = "Dell Command"; wgt = "Dell.CommandUpdate" }
    @{name = "7 Zip"; wgt = "7zip.7zip"; wpck = "7-zip" }
    @{name = "Google Chrome"; wgt = "Google.Chrome"; wpck = "Google Chrome" }
    @{name = "Adobe Acrobat Reader"; wgt = "Adobe.Acrobat.Reader.64-bit"; wpck = "Adobe Acrobat" }
    @{name = "Microsoft Edge"; wgt = "Microsoft.Edge"; wpck = "Microsoft Edge" }
    @{name = "Microsoft Office"; wgt = "Microsoft.Office"; wpck = "Office" }
    @{name = "Microsoft Teams"; wgt = "Microsoft.Teams"; wpck = "" }
    #Installation locale
    @{name = "Dell Encryption"; fname = "DDPE_v11.1-force.bat"; wgt = ""; wpck = "Dell Encryption" }
    @{name = "Cyberwatch Agent"; fname = "cyberwatch_agent-x64-4.11-SRVCYB01PRD.msi"; wgt = ""; wpck = "CyberwatchAgent" }
    @{name = "Trellix"; fname = "McAfee_agent.5.7.6.251.exe"; wgt = ""; wpck = "Trellix" }
    @{name = "Multiplateforme"; fname = "multiplatform_201812041125_x64.exe"; wgt = "" }
    @{name = "Forcepoint"; fname = "instal_force_proxy_forcepoint.bat"; wgt = ""; wpck = "FORCEPOINT" }
    @{name = "Global Protect"; fname = "GlobalProtect64-6.0.0.msi"; wgt = ""; wpck = "GlobalProtect" }
)

# output installation
$output = @()

# print logo
$P_logo = {
    Clear-Host
    foreach ($l in $logo) {
        [System.ConsoleColor]$cor = "white"
        if (!([string]::IsNullOrEmpty($l.cor))) {
            $cor = $l.cor
        }
        Write-Host $l.t -f $cor
    }
}

# print message bye
$P_bye = {
    Clear-Host
    foreach ($l in $bye) {
        [System.ConsoleColor]$cor = "white"
        if (!([string]::IsNullOrEmpty($l.cor))) {
            $cor = $l.cor
        }
        Write-Host $l.t -f $cor
    }
}

Function CursorHover {
    param ([int]$p, [int]$off, [hashtable[]]$CHINT)
    [Console]::SetCursorPosition(1, $p + $off)
    Write-Host "    " -NoNewline
    Write-Host "" $CHINT[$p].t -BackgroundColor White -f black
    [Console]::SetCursorPosition(0, $p + $off)
}

Function CursorNormal {
    param ([int]$p, [int]$off, [hashtable[]]$CNINT)
    [Console]::SetCursorPosition(1, $p + $off)
    Write-Host "    " $CNINT[$p].t
}

Function ShowMenu {
    param([hashtable[]] $Interface, [int]$off)
    $pos = 0
    $P_logo.Invoke()
    
    [Console]::SetCursorPosition(0, $off)
    foreach ($l in $Interface) {
        switch ($l.act) {
            "boo" { 
                if (($l.dta -eq "0") -or ([string]::IsNullOrEmpty($l.dta))) {
                    Write-Host "     " $l.t "         [ ]"
                }
                else {
                    Write-Host "     " $l.t "         [X]"
                }
            }
            Default { Write-Host "     " $l.t "        " $l.dta }
        }
        
    } 
    CursorHover -p $pos -off $off -CHINT $Interface
    $menu = $true
    while ($menu) {
        if ([Console]::KeyAvailable) {
            $x = $Host.UI.RawUI.ReadKey()
            [Console]::SetCursorPosition(1, $pos + $off)
            CursorNormal -p $pos -off $off -CNINT $Interface
            switch ($x.VirtualKeyCode) { 
                38 {
                    #down key
                    if ($pos -gt 0) {
                        $pos -= 1
                    }
                }
                40 {
                    #up key
                    if ($pos -lt $Interface.Length - 1) {
                        $pos += 1
                    }
                }
                13 {
                    #enter key
                    switch ($Interface[$pos].act) {
                        "sub" {
                            $menu = $false;
                            switch ($Interface[$pos].vlr) {
                                "ins" { $ins.Invoke() }
                                "srt" { $srt.Invoke() }
                                "ver" { $ver.Invoke() }
                                "mga" { $mga.Invoke() }
                            }
                        }
                        "edt" {
                            [Console]::SetCursorPosition(50, $pos + $off)
                            [console]::CursorVisible = $true
                            Write-Host " > " -f green -NoNewline
                            $Interface[$pos].dta = $Host.UI.ReadLine()
                            [console]::CursorVisible = $false
                            Clear-Host
                            ShowMenu -Interface $Interface -off $off
                        }
                        "boo" { 
                            [Console]::SetCursorPosition(0, $pos + $off)
                            if ($Interface[$pos].dta -eq "0") {
                                $Interface[$pos].dta = "1"
                                Write-Host "     " $Interface[$pos].t "         [X]"
                            }
                            else {
                                $Interface[$pos].dta = "0"
                                Write-Host "     " $Interface[$pos].t "         [ ]"
                            }
                        }
                    }
                }
            }
        }
        CursorHover -p $pos -off $off -CHINT $Interface
        Start-Sleep -Milliseconds 90 
    }

}

Function C_Console {
    param ([int]$off)
    for ($x = 0; $x -le 80; $x++) {
        for ($y = $off; $y -le 25; $y++) {
            [Console]::SetCursorPosition($x, $y)
            Write-Host " " -BackgroundColor Black
        }
        Start-Sleep -Milliseconds 5
    }
    [Console]::SetCursorPosition(0, $off)
}

$ins = {
    C_Console 13
    #region Installation Winget
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
    }
    #endregion

    #region 

    #region Copie dossier pstools
    C_Console  -off 13
    Write-Host "`n   Copie du dossier pstools vers le C:" -f white
    if (!(Test-Path -path "C:\pstools")) {New-Item "C:\pstools" -Type Directory}
    Copy-Item -Path 'C:\RESTANT_A_FAIRE\pstools\*' -Destination 'C:\pstools\' -Recurse | Out-Null
    Start-sleep -Seconds 3
    C_Console 13
    if(Test-Path -Path C:\pstools\mdt_flag.txt -PathType Leaf) {
        Write-Host "`n   Copie du dossier pstools vers le C:" -f green
    } else {
        Write-Host "`n   Copie du dossier pstools vers le C:" -f red
        Write-Host "`n Verifier si le dossier existe ou il est sur le bon endroit`n" -f red
        " Appuyez sur Entrée pour continuer ou CTRL+C pour quitter"
        $Host.UI.ReadLine()
        return
    }
    #endregion

    #region Installation Applications
    $output += "    Installation d'Application`n"
    C_Console -off 13
    Write-Host "    Installation" -f cyan
    foreach($p in $prog) {
        C_Console -off 14
        try {
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
                C_Console -off 14

            }
        }catch {
            $output += $p.name + ": NOT OK`n"
        }finally {
            $output += $p.name + ": OK`n"
        }
    }
    #endregion

    <#
    #region MAJ Microsoft Store

    C_Console -off 13
    Write-Host "    MAJ Windows Store" -f cyan
    winget upgrade --all
    #endregion
    #>

    #region One Drive
    C_Console -off 13
    Set-itemproperty -Path 'HKLM:\HKEY_LOCAL_MACHINE\Software\Policies\Microsoft\Windows\OneDrive' -name DisableFileSyncNGSC 0
    #endregion

    #region Layout

    if (!(Test-Path -path "C:\temp")) {New-Item "C:\temp" -Type Directory}
    Copy-Item "C:\RESTANT_A_FAIRE\pstools\layout2023.xml" -Destination "C:\temp" -Recurse | Out-Null
    Import-StartLayout -LayoutPath C:\Temp\layout2023.xml -MountPath C:\

    #endregion

    #region MAJ Windows
    try {
        Clear-Host
        $P_logo.Invoke()
        [Console]::SetCursorPosition(0, 13);
        Write-Host "    MAJ Windows" -f cyan
        Install-Module -Name PSWindowsUpdate -Force
        Get-WindowsUpdate -AcceptAll -Install -IgnoreReboot
    } catch {
        $output += "MAJ Windows:    NOT OK"
    } finally {
        $output += "MAJ Windows:    OK"
    }
    #endregion

    <#
    #region Ne pas toucher
    try {
        if($form[0].dta -notmatch "interiale.intra") {
            Add-Computer -LocalCredential interiale.intra\adm.XX -DomainName interiale.intra -OUPath "OU=testOU" -Credential interiale.intra\adm.XX
        }
    } catch {
        $output += "Domain:     NOT OK"
    } finally {
        $output += "Domain:     OK"
    }

    

    #endregion
    #>
    

    Clear-Host
    if ($form[2].dta -eq "1") {
        $output | Out-File "Log.txt"
        Write-Host "    Fichier Log - c:\Log.txt`n" -f Red
    }
    Write-Host "`n`n        Voulez-vous verifier l'installation ?"
    ShowMenu -Interface $fver -off 17
}

$srt = {
    $P_bye.Invoke()
    Start-Sleep -Seconds 2
    Exit-PSSession
}

$ver = {
    Clear-Host
    $P_logo.Invoke()
    Write-Host " `n Chargement des programmes...`n" -f Yellow
    $pack = get-package | Select-Object name -unique -Wait
    
    C_Console -off 13
    $output += "    Verifier Installation`n"
    Foreach($prg in $prog) {
        if($prg.wpck -ne "") {
            Write-host " " $prg.name ":   " -f yellow -NoNewline
            for($i = 0; $i -le $pack.Count - 1; $i++) {
                if($pack[$i].name -like "*"+$prg.wpck+"*") {
                    Write-Host "Installe " -f green
                    $output += $prg.name + ":   Installe`n"
                    break
                } elseif ($i -eq $pack.Count - 1) {
                    write-host "Pas installe" -f red
                    $output += $prg.name + ":   Pas Installe`n"
                }
            }
        }
        Start-Sleep -Milliseconds 50
    }

    Start-Sleep -Seconds 4
    C_Console -off 13

    if ($form[2].dta -eq "1") {
        
        $output | Out-File "Log.txt"
        Write-Host "    Fichier Log - c:\Log.txt`n" -f Red
    }

    Write-Host "`n       Entree pour continuer" -f cyan
    $Host.UI.ReadLine()
    Clear-Host
    ShowMenu -Interface $form -off 13

}

$mga = {
    Clear-Host
    Remove-Item -Path "C:\RESTANT_A_FAIRE" -Force -Recurse
    Exit-PSSession
}

#endregion
$domain = Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select-Object Domain
$np = Get-WmiObject -Namespace root\cimv2 -Class Win32_ComputerSystem | Select-Object name
$form[0].dta = $domain.Domain
$form[1].dta = $np.name
ShowMenu -Interface $form -off 13