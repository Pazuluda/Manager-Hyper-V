# ==============================
#        OUTIL HYPER-V 
# ==============================

function Show-Tableau {
    param (
        [string[]]$Lignes,
        [string]$CouleurBordure = "Cyan",
        [string]$CouleurTexte = "White"
    )
    $largeurConsole = [console]::WindowWidth
    $maxLen = ($Lignes | Measure-Object -Property Length -Maximum).Maximum
    $largeurTableau = [Math]::Min($maxLen + 6, $largeurConsole - 2)
    $maxContent = $largeurTableau - 4
    $finalLignes = @()
    foreach ($ligne in $Lignes) {
        while ($ligne.Length -gt $maxContent) {
            $finalLignes += $ligne.Substring(0, $maxContent)
            $ligne = $ligne.Substring($maxContent)
        }
        $finalLignes += $ligne
    }
    $videHaut = [Math]::Max(0, [int](($host.UI.RawUI.WindowSize.Height - $finalLignes.Count - 4) / 2))
    for ($i = 0; $i -lt $videHaut; $i++) { Write-Host "" }
    $bordure = "+" + ("-" * ($largeurTableau - 2)) + "+"
    Write-Host (" " * [Math]::Max(0,($largeurConsole - $largeurTableau)/2)) -NoNewline
    Write-Host $bordure -ForegroundColor $CouleurBordure
    foreach ($ligne in $finalLignes) {
        $txt = " $ligne"
        $pad = $largeurTableau - 2 - ($txt).Length
        if ($pad -lt 0) { $pad = 0 }
        $gauche = [Math]::Max(0,($largeurConsole - $largeurTableau)/2)
        Write-Host (" " * $gauche + "|" + $txt + (" " * $pad) + "|") -ForegroundColor $CouleurTexte
    }
    Write-Host (" " * [Math]::Max(0,($largeurConsole - $largeurTableau)/2)) -NoNewline
    Write-Host $bordure -ForegroundColor $CouleurBordure
}

function Saisir-Nombre($MESSAGE) {
    DO {
        $VALEUR = READ-HOST $MESSAGE
        IF ($VALEUR -match '^\d+$') { RETURN [int]$VALEUR }
        Clear-Host
        Show-Tableau @("SAISISSEZ UNIQUEMENT UN NOMBRE (EX : 4)") -CouleurBordure "Red" -CouleurTexte "White"
    } WHILE ($TRUE)
}

function Barre-ProgressSimulee($Message) {
    $steps = 25
    $startTime = Get-Date
    for ($j = 1; $j -le $steps; $j++) {
        $percent = [math]::Round($j * (100 / $steps))
        $elapsed = (Get-Date) - $startTime
        $timer = $elapsed.ToString("hh\:mm\:ss")
        Write-Host -NoNewline "`r$Message : $percent %  ($timer)     "
        Start-Sleep -Milliseconds 100
    }
    Write-Host ""
}

function Export-VM-AvcProgress($Name, $ExportPath) {
    Clear-Host
    Show-Tableau @("EXPORTATION DE $Name VERS $ExportPath...") -CouleurBordure "Yellow"
    EXPORT-VM -NAME $Name -PATH $ExportPath | Out-Null
    Barre-ProgressSimulee "EXPORT EN COURS"
    Clear-Host
    Show-Tableau @("EXPORT TERMINE !") -CouleurBordure "Green"
}

function Detect-TypeOS($VMName) {
    $result = @{}
    if ($VMName -match "windows|win|core|serveur|server") {
        $result.cpu = 2
        $result.msg = "TYPE DETECTE : WINDOWS/SERVER/CORE - PROPOSITION : 2 COEURS"
    } elseif ($VMName -match "linux|debian|ubuntu|centos|freebsd|firewall|pf|pfsense|openwrt|alpine|vyos|opnsense") {
        $result.cpu = 1
        $result.msg = "TYPE DETECTE : LINUX/BSD/FIREWALL/AUTRES - PROPOSITION : 1 COEUR"
    } else {
        $result.cpu = 1
        $result.msg = "OS NON DETECTE, PAR DEFAUT : 1 COEUR"
    }
    return $result
}

function Selection-Env {
    DO {
        Clear-Host
        Show-Tableau @(
            "",
            "SELECTION DE L'ENVIRONNEMENT",
            "",
            "1. WINDOWS SERVEUR CORE",
            "2. WINDOWS SERVEUR CLASSIQUE",
            "3. WINDOWS NORMAL (DESKTOP/PRO)",
            "",
            "R. RAFRAICHIR / RECENTRER LE MENU",
            ""
        ) -CouleurBordure "Magenta" -CouleurTexte "Yellow"
        $TYPE_WINDOWS = READ-HOST "CHOISISSEZ VOTRE ENVIRONNEMENT (1-3, R)"
        IF ($TYPE_WINDOWS.ToUpper() -eq "R") { CONTINUE }
    } WHILE ($TYPE_WINDOWS -notin "1","2","3")
    RETURN $TYPE_WINDOWS
}

function Copy-VM-AvcProgress($SRC, $DEST_PARENT) {
    $VM_FOLDER_NAME = Split-Path $SRC -Leaf
    $DEST_FINAL = Join-Path $DEST_PARENT $VM_FOLDER_NAME
    if (Test-Path $DEST_FINAL) {
        Clear-Host
        Show-Tableau @("LE DOSSIER $DEST_FINAL EXISTE DEJA. SUPPRIME-LE AVANT OU CHOISIS UN AUTRE NOM DE DOSSIER.") -CouleurBordure "Red"
        return $null
    }
    Clear-Host
    Show-Tableau @("COPIE DE $SRC VERS $DEST_FINAL...") -CouleurBordure "Yellow"
    robocopy $SRC $DEST_FINAL /E /COPYALL /R:1 /W:1 | Out-Null
    Barre-ProgressSimulee "COPIE EN COURS"
    Clear-Host
    Show-Tableau @("COPIE TERMINEE !") -CouleurBordure "Green"
    return $DEST_FINAL
}

$CHEMIN_ISO_DEFAUT = "D:\IMAGES_ISO\INSTALL_WINDOWS10.ISO"

function Afficher-MenuDepart {
    Clear-Host
    Show-Tableau @(
        "",
        "OUTIL SCRIPT GESTION HYPER-V EN LIGNE DE COMMANDE",
        "",
        "1. MODE MENU AVANCE DOUBLE NIVEAU",
        "",
        "R. RAFRAICHIR / RECENTRER LE MENU",
        "Q. QUITTER"
    ) -CouleurBordure "Blue" -CouleurTexte "White"
}

function MENU_AVANCE {
    $TYPE_WINDOWS = Selection-Env
    DO {
        Clear-Host
        Show-Tableau @(
            "",
            "MODE MENU AVANCE",
            "",
            "1. GESTION DES MACHINES VIRTUELLES",
            "2. GESTION DES COMMUTATEURS VIRTUELS",
            "3. GESTION DES SNAPSHOTS",
            "4. GESTION DU STOCKAGE",
            "",
            "R. RAFRAICHIR / RECENTRER LE MENU",
            "B. RETOUR MENU PRINCIPAL"
        ) -CouleurBordure "DarkCyan" -CouleurTexte "White"
        $CHOIX_AVANCE = READ-HOST "VOTRE CHOIX (1-4, R ou B)"
        SWITCH ($CHOIX_AVANCE.ToUpper()) {
            1 {
                DO {
                    Clear-Host
                    Show-Tableau @(
                        "",
                        "GESTION DES MACHINES VIRTUELLES",
                        "",
                        "1. LISTER LES VMS",
                        "2. DEMARRER UNE VM",
                        "3. ARRETER UNE VM",
                        "4. CREER UNE VM AVANCEE",
                        "5. SUPPRIMER UNE VM",
                        "6. EXPORTER UNE VM",
                        "7. DEPLOYER UNE VM DE STOCK",
                        "",
                        "R. RAFRAICHIR / RECENTRER LE MENU",
                        "B. RETOUR MENU AVANCE"
                    ) -CouleurBordure "DarkYellow" -CouleurTexte "Gray"
                    $CHOIX_VM = READ-HOST "VOTRE CHOIX (1-7, R ou B)"
                    SWITCH ($CHOIX_VM.ToUpper()) {
                        1 {
                            Clear-Host
                            $vms = GET-VM | SELECT-OBJECT NAME, STATE, CPUUSAGE, MEMORYASSIGNED, UPTIME
                            $lines = @("LISTE DES MACHINES VIRTUELLES :","")
                            foreach ($vm in $vms) {
                                $lines += ("{0,-25} | Etat: {1,-10} | CPU: {2,3}% | RAM: {3,6}MB | Uptime: {4}" -f $vm.Name, $vm.State, $vm.CPUUsage, $vm.MemoryAssigned, $vm.Uptime)
                            }
                            Show-Tableau $lines -CouleurBordure "Cyan" -CouleurTexte "White"
                        }
                        2 {
                            $NAME = READ-HOST "NOM DE LA VM A DEMARRER"
                            START-VM -NAME $NAME
                            Clear-Host
                            Show-Tableau @("VM $NAME DEMARREE.") -CouleurBordure "Green"
                        }
                        3 {
                            $NAME = READ-HOST "NOM DE LA VM A ARRETER"
                            STOP-VM -NAME $NAME
                            Clear-Host
                            Show-Tableau @("VM $NAME ARRETEE.") -CouleurBordure "Green"
                        }
                        4 {
                            DO {
                                Clear-Host
                                Show-Tableau @(
                                    "",
                                    "CREATION VM AVANCEE",
                                    "",
                                    "1. CREER UNE SEULE VM AVEC ISO",
                                    "2. CREER PLUSIEURS VMS D'UN COUP AVEC ISO",
                                    "3. CREER UNE VM DE TEST (SANS ISO)",
                                    "4. CREER UNE VM VIDE (SANS ISO NI VHDX)",
                                    "",
                                    "R. RAFRAICHIR / RECENTRER LE MENU",
                                    "B. RETOUR"
                                ) -CouleurBordure "DarkGreen" -CouleurTexte "White"
                                $choix_create = READ-HOST "VOTRE CHOIX (1-4, R ou B)"
                                SWITCH ($choix_create.ToUpper()) {
                                    1 {
                                        $NAME = READ-HOST "NOM DE LA NOUVELLE VM"
                                        $RAMGO = Saisir-Nombre "MEMOIRE RAM EN GO (EX : 4)"
                                        $MEMORY = [long]($RAMGO * 1GB)
                                        $CPU = Saisir-Nombre "NOMBRE DE COEURS DE PROCESSEUR (EX : 2)"
                                        $VHDX = READ-HOST "CHEMIN DU DISQUE VHDX (EX : D:\HYPERV\$NAME.VHDX)"
                                        $SWITCH = READ-HOST "NOM DU SWITCH VIRTUEL"
                                        IF (-NOT (TEST-PATH $VHDX)) {
                                            $TAILLE_VHDX = Saisir-Nombre "TAILLE DU DISQUE VHDX EN GO (EX : 60)"
                                            $TAILLE_VHDX_BYTES = [long]($TAILLE_VHDX * 1GB)
                                            NEW-VHD -PATH $VHDX -SIZEBYTES $TAILLE_VHDX_BYTES -DYNAMIC
                                            Clear-Host
                                            Show-Tableau @("DISQUE $VHDX CREE.") -CouleurBordure "Green"
                                        }
                                        Show-Tableau @("CHEMIN ISO UTILISE : $CHEMIN_ISO_DEFAUT") -CouleurBordure "Blue"
                                        TRY {
                                            $vm = NEW-VM -NAME $NAME -MEMORYSTARTUPBYTES $MEMORY -SWITCHNAME $SWITCH -VHDPATH $VHDX -ERRORACTION STOP
                                            IF ($vm) {
                                                SET-VM -NAME $NAME -PROCESSORCOUNT $CPU
                                                SET-VMDVDDRIVE -VMNAME $NAME -PATH $CHEMIN_ISO_DEFAUT
                                                START-VM -NAME $NAME
                                                Clear-Host
                                                Show-Tableau @("VM $NAME CREEE ET DEMARREE.") -CouleurBordure "Green"
                                            }
                                        } CATCH {
                                            Clear-Host
                                            Show-Tableau @("ERREUR LORS DE LA CREATION DE LA VM : $($_.Exception.Message)") -CouleurBordure "Red"
                                        }
                                    }
                                    2 {
                                        $NB = Saisir-Nombre "NOMBRE DE VMS A CREER (EX : 2)"
                                        $PREFIX = READ-HOST "PREFIXE POUR LES NOMS DE VM (EX : TESTVM)"
                                        $RAMGO = Saisir-Nombre "MEMOIRE RAM EN GO (EX : 4)"
                                        $MEMORY = [long]($RAMGO * 1GB)
                                        $CPU = Saisir-Nombre "NOMBRE DE COEURS DE PROCESSEUR (EX : 2)"
                                        $VHDX_PATH = READ-HOST "CHEMIN DOSSIER VHDX (EX : D:\HYPERV)"
                                        $SWITCH = READ-HOST "NOM DU SWITCH VIRTUEL"
                                        $TAILLE_VHDX = Saisir-Nombre "TAILLE DU DISQUE VHDX EN GO (EX : 60)"
                                        $TAILLE_VHDX_BYTES = [long]($TAILLE_VHDX * 1GB)
                                        Show-Tableau @("CHEMIN ISO UTILISE : $CHEMIN_ISO_DEFAUT") -CouleurBordure "Blue"
                                        FOR ($i = 1; $i -le $NB; $i++) {
                                            $VMNAME = "$PREFIX$i"
                                            $VHDX = "$VHDX_PATH\$VMNAME.VHDX"
                                            IF (-NOT (TEST-PATH $VHDX)) {
                                                NEW-VHD -PATH $VHDX -SIZEBYTES $TAILLE_VHDX_BYTES -DYNAMIC
                                                Clear-Host
                                                Show-Tableau @("DISQUE $VHDX CREE.") -CouleurBordure "Green"
                                            }
                                            TRY {
                                                $vm = NEW-VM -NAME $VMNAME -MEMORYSTARTUPBYTES $MEMORY -SWITCHNAME $SWITCH -VHDPATH $VHDX -ERRORACTION STOP
                                                IF ($vm) {
                                                    SET-VM -NAME $VMNAME -PROCESSORCOUNT $CPU
                                                    SET-VMDVDDRIVE -VMNAME $VMNAME -PATH $CHEMIN_ISO_DEFAUT
                                                    START-VM -NAME $VMNAME
                                                    Clear-Host
                                                    Show-Tableau @("VM $VMNAME CREEE ET DEMARREE.") -CouleurBordure "Green"
                                                }
                                            } CATCH {
                                                Clear-Host
                                                Show-Tableau @("ERREUR LORS DE LA CREATION DE LA VM $VMNAME : $($_.Exception.Message)") -CouleurBordure "Red"
                                            }
                                        }
                                    }
                                    3 {
                                        $NAME = READ-HOST "NOM DE LA VM DE TEST"
                                        $RAMGO = Saisir-Nombre "MEMOIRE RAM EN GO (EX : 4)"
                                        $MEMORY = [long]($RAMGO * 1GB)
                                        $CPU = Saisir-Nombre "NOMBRE DE COEURS DE PROCESSEUR (EX : 2)"
                                        $VHDX = READ-HOST "CHEMIN DU DISQUE VHDX (EX : D:\HYPERV\$NAME.VHDX)"
                                        $SWITCH = READ-HOST "NOM DU SWITCH VIRTUEL"
                                        IF (-NOT (TEST-PATH $VHDX)) {
                                            $TAILLE_VHDX = Saisir-Nombre "TAILLE DU DISQUE VHDX EN GO (EX : 60)"
                                            $TAILLE_VHDX_BYTES = [long]($TAILLE_VHDX * 1GB)
                                            NEW-VHD -PATH $VHDX -SIZEBYTES $TAILLE_VHDX_BYTES -DYNAMIC
                                            Clear-Host
                                            Show-Tableau @("DISQUE $VHDX CREE.") -CouleurBordure "Green"
                                        }
                                        TRY {
                                            $vm = NEW-VM -NAME $NAME -MEMORYSTARTUPBYTES $MEMORY -SWITCHNAME $SWITCH -VHDPATH $VHDX -ERRORACTION STOP
                                            IF ($vm) {
                                                SET-VM -NAME $NAME -PROCESSORCOUNT $CPU
                                                START-VM -NAME $NAME
                                                Clear-Host
                                                Show-Tableau @("VM $NAME CREEE ET DEMARREE SANS ISO.") -CouleurBordure "Green"
                                            }
                                        } CATCH {
                                            Clear-Host
                                            Show-Tableau @("ERREUR LORS DE LA CREATION DE LA VM : $($_.Exception.Message)") -CouleurBordure "Red"
                                        }
                                    }
                                    4 {
                                        $NAME = READ-HOST "NOM DE LA VM VIDE"
                                        $RAMGO = Saisir-Nombre "MEMOIRE RAM EN GO (EX : 4)"
                                        $MEMORY = [long]($RAMGO * 1GB)
                                        $CPU = Saisir-Nombre "NOMBRE DE COEURS DE PROCESSEUR (EX : 2)"
                                        $SWITCH = READ-HOST "NOM DU SWITCH VIRTUEL"
                                        TRY {
                                            $vm = NEW-VM -NAME $NAME -MEMORYSTARTUPBYTES $MEMORY -SWITCHNAME $SWITCH -ERRORACTION STOP
                                            IF ($vm) {
                                                SET-VM -NAME $NAME -PROCESSORCOUNT $CPU
                                                START-VM -NAME $NAME
                                                Clear-Host
                                                Show-Tableau @("VM $NAME CREEE ET DEMARREE SANS ISO NI VHDX.") -CouleurBordure "Green"
                                            }
                                        } CATCH {
                                            Clear-Host
                                            Show-Tableau @("ERREUR LORS DE LA CREATION DE LA VM : $($_.Exception.Message)") -CouleurBordure "Red"
                                        }
                                    }
                                    "B" { Clear-Host; Show-Tableau @("RETOUR") -CouleurBordure "Magenta" }
                                    "R" { CONTINUE }
                                    DEFAULT { Clear-Host; Show-Tableau @("CHOIX INVALIDE") -CouleurBordure "Red" }
                                }
                                IF ($choix_create.ToUpper() -NE "B" -and $choix_create.ToUpper() -NE "R") { PAUSE }
                            } WHILE ($choix_create.ToUpper() -NE "B")
                        }
                        5 {
                            $NAME = READ-HOST "NOM DE LA VM A SUPPRIMER"
                            STOP-VM -NAME $NAME -FORCE
                            REMOVE-VM -NAME $NAME -FORCE
                            Clear-Host
                            Show-Tableau @("VM $NAME SUPPRIMEE.") -CouleurBordure "Green"
                        }
                        6 {
                            $NAME = READ-HOST "NOM DE LA VM A EXPORTER"
                            $EXPORT_PATH = READ-HOST "CHEMIN DE DESTINATION POUR L'EXPORT (EX : D:\STOCK_VMS\$NAME)"
                            Export-VM-AvcProgress $NAME $EXPORT_PATH
                        }
                        7 {
                            Clear-Host
                            $STOCK_PATH = READ-HOST "CHEMIN DU DOSSIER DE STOCKAGE DES VMS PRETES (EX : D:\STOCK_VMS)"
                            $VM_LIST = GET-CHILDITEM -PATH $STOCK_PATH -DIRECTORY | SELECT-OBJECT -EXPAND NAME
                            if ($VM_LIST.Count -eq 0) {
                                Show-Tableau @("AUCUNE VM EN STOCK TROUVEE.") -CouleurBordure "Red"
                            } else {
                                $lines = @("VMS DISPONIBLES EN STOCK :")
                                for ($i = 0; $i -lt $VM_LIST.Count; $i++) {
                                    $lines += ("{0,2}. {1}" -f ($i+1), $VM_LIST[$i])
                                }
                                Show-Tableau $lines -CouleurBordure "Yellow" -CouleurTexte "White"
                                DO {
                                    $CHOIX_STOCK = Saisir-Nombre "NUMERO DE LA VM A DEPLOYER"
                                } WHILE ($CHOIX_STOCK -lt 1 -OR $CHOIX_STOCK -gt $VM_LIST.Count)
                                $VM_TO_COPY = $VM_LIST[$CHOIX_STOCK-1]
                                $SRC_PATH = "$STOCK_PATH\$VM_TO_COPY"
                                $DEST_PARENT = READ-HOST "DOSSIER DE DESTINATION (EX : D:\HYPERV)"
                                $DEST_FINAL = Copy-VM-AvcProgress $SRC_PATH $DEST_PARENT
                                IF (-NOT $DEST_FINAL) { BREAK }
                                $CONFIG_PATH = Get-ChildItem -Path $DEST_FINAL -Recurse -Include *.vmcx,*.xml -ErrorAction SilentlyContinue | Select-Object -First 1
                                IF ($CONFIG_PATH) {
                                    Show-Tableau @("FICHIER DE CONFIGURATION TROUVE : $($CONFIG_PATH.FullName)") -CouleurBordure "Blue"
                                    Barre-ProgressSimulee "IMPORTATION EN COURS"
                                    $DEST_VM = READ-HOST "DOSSIER DE CONFIGURATION DE LA VM"
                                    $DEST_CP = READ-HOST "DOSSIER DES POINTS DE CONTROLE"
                                    $DEST_SP = READ-HOST "DOSSIER DE PAGINATION INTELLIGENTE"
                                    $DEST_VHD = READ-HOST "DOSSIER DES VHDX"
                                    TRY {
                                        $IMPORTED_VM = Import-VM -Path $CONFIG_PATH.FullName `
                                            -Copy `
                                            -GenerateNewId `
                                            -VirtualMachinePath $DEST_VM `
                                            -SnapshotFilePath $DEST_CP `
                                            -SmartPagingFilePath $DEST_SP `
                                            -VhdDestinationPath $DEST_VHD -ErrorAction Stop
                                        Clear-Host
                                        Show-Tableau @("IMPORT TERMINE !") -CouleurBordure "Green"
                                        $VMS = @()
                                        IF ($IMPORTED_VM -is [System.Collections.IEnumerable]) {
                                            $VMS = $IMPORTED_VM
                                        } ELSE {
                                            $VMS = @($IMPORTED_VM)
                                        }
                                        foreach ($VM in $VMS) {
                                            $VMNAME = $VM.Name
                                            $MIN_CPU = 1
                                            $MAX_CPU = [Math]::Min(12, (Get-VMHost).LogicalProcessorCount)
                                            $osInfo = Detect-TypeOS $VMNAME
                                            $CPU_AUTO = $osInfo.cpu
                                            Show-Tableau @($osInfo.msg) -CouleurBordure "Cyan"
                                            DO {
                                                $CPU_INPUT = Read-Host "NOMBRE DE COEURS DE PROCESSEUR POUR $VMNAME (ENTRE $MIN_CPU ET $MAX_CPU) [`$CPU_AUTO` par defaut]"
                                                if ([string]::IsNullOrWhiteSpace($CPU_INPUT)) {
                                                    $CPU = $CPU_AUTO
                                                } elseif ($CPU_INPUT -match '^\d+$' -and [int]$CPU_INPUT -ge $MIN_CPU -and [int]$CPU_INPUT -le $MAX_CPU) {
                                                    $CPU = [int]$CPU_INPUT
                                                } else {
                                                    Show-Tableau @("NOMBRE DE COEURS INVALIDE. ENTREE IGNORÉE.") -CouleurBordure "Red"
                                                    $CPU = $null
                                                }
                                            } while ($CPU -eq $null)
                                            Set-VMProcessor -VMName $VMNAME -Count $CPU
                                            $NICs = Get-VMNetworkAdapter -VMName $VMNAME
                                            foreach ($NIC in $NICs) {
                                                $DEFAULT_COM = $NIC.SwitchName
                                                $COM = Read-Host "NOM DU COMMUTATEUR (VIRTUAL SWITCH) POUR LA CARTE RESEAU '$($NIC.Name)' (par defaut : $DEFAULT_COM)"
                                                if ([string]::IsNullOrWhiteSpace($COM)) { $COM = $DEFAULT_COM }
                                                if (-not (Get-VMSwitch | Where-Object {$_.Name -eq $COM})) {
                                                    Show-Tableau @("COMMUTATEUR '$COM' INEXISTANT.") -CouleurBordure "Red"
                                                    $TYPE = Read-Host "TYPE POUR $COM (EXTERNAL/INTERNAL/PRIVATE)"
                                                    if ($TYPE -eq "EXTERNAL") {
                                                        $adapters = GET-NETADAPTER | WHERE-OBJECT {$_.STATUS -EQ "UP"}
                                                        $aList = @("ADAPTATEURS DISPONIBLES :")
                                                        foreach ($a in $adapters) { $aList += ("- {0}" -f $a.Name) }
                                                        Show-Tableau $aList -CouleurBordure "Cyan" -CouleurTexte "White"
                                                        $adapter = Read-Host "NOM DE L'ADAPTATEUR POUR L'EXTERNAL"
                                                        NEW-VMSWITCH -NAME $COM -NETADAPTERNAME $adapter -ALLOWMANAGEMENTOS $TRUE
                                                    } elseif ($TYPE -eq "INTERNAL" -or $TYPE -eq "PRIVATE") {
                                                        NEW-VMSWITCH -NAME $COM -SWITCHTYPE $TYPE
                                                    } else {
                                                        Show-Tableau @("TYPE NON VALIDE, COMMUTATEUR NON CREE.") -CouleurBordure "Red"
                                                    }
                                                }
                                                Connect-VMNetworkAdapter -VMName $VMNAME -SwitchName $COM
                                            }
                                            $RAMGO = Saisir-Nombre "MEMOIRE RAM EN GO POUR $VMNAME (EX : 4)"
                                            $MEMORY = [long]($RAMGO * 1GB)
                                            Set-VM -Name $VMNAME -MemoryStartupBytes $MEMORY
                                            Show-Tableau @("VM $VMNAME CONFIGUREE AVEC $CPU COEURS, $RAMGO GO RAM, ET COMMUTATEUR RESEAU ($COM) OK.") -CouleurBordure "Green"
                                        }
                                    } CATCH {
                                        Clear-Host
                                        Show-Tableau @("ERREUR LORS DE L'IMPORT DE LA VM : $($_.Exception.Message)") -CouleurBordure "Red"
                                    }
                                } ELSE {
                                    Clear-Host
                                    Show-Tableau @("ERREUR : AUCUN FICHIER DE CONFIGURATION (.VMCX OU .XML) TROUVE DANS $DEST_FINAL") -CouleurBordure "Red"
                                }
                            }
                            PAUSE
                        }
                        "B" { Clear-Host; Show-Tableau @("RETOUR MENU AVANCE") -CouleurBordure "Magenta" }
                        "R" { CONTINUE }
                        DEFAULT { Clear-Host; Show-Tableau @("CHOIX INVALIDE") -CouleurBordure "Red" }
                    }
                    IF ($CHOIX_VM.ToUpper() -NE "B" -and $CHOIX_VM.ToUpper() -NE "R") { PAUSE }
                } WHILE ($CHOIX_VM.ToUpper() -NE "B")
            }
            2 {
                DO {
                    Clear-Host
                    Show-Tableau @(
                        "",
                        "GESTION COMMUTATEUR HYPER-V",
                        "",
                        "1. LISTER LES COMMUTATEURS",
                        "2. CREER UN COMMUTATEUR",
                        "3. SUPPRIMER UN COMMUTATEUR",
                        "",
                        "R. RAFRAICHIR / RECENTRER LE MENU",
                        "B. RETOUR MENU AVANCE"
                    ) -CouleurBordure "DarkMagenta" -CouleurTexte "Gray"
                    $CHOIXC = READ-HOST "VOTRE CHOIX (1-3, R ou B)"
                    SWITCH ($CHOIXC.ToUpper()) {
                        1 {
                            Clear-Host
                            $switchs = GET-VMSWITCH | SELECT-OBJECT NAME, SWITCHTYPE
                            $lines = @("LISTE DES COMMUTATEURS HYPER-V :","")
                            foreach ($s in $switchs) {
                                $lines += ("{0,-30} | Type: {1,-10}" -f $s.Name, $s.SwitchType)
                            }
                            Show-Tableau $lines -CouleurBordure "Cyan" -CouleurTexte "White"
                        }
                        2 {
                            $NOMC = READ-HOST "NOM DU COMMUTATEUR (EX : COM_LAN_PRENOM)"
                            $TYPE = READ-HOST "TYPE (EXTERNAL, INTERNAL, PRIVATE)"
                            IF ($TYPE -EQ "EXTERNAL") {
                                $ADAPTERS = GET-NETADAPTER | WHERE-OBJECT {$_.STATUS -EQ "UP"}
                                $aList = @("ADAPTATEURS DISPONIBLES :")
                                foreach ($a in $ADAPTERS) {
                                    $aList += ("- {0}" -f $a.Name)
                                }
                                Show-Tableau $aList -CouleurBordure "Cyan" -CouleurTexte "White"
                                $ADAPTER = READ-HOST "NOM DE L'ADAPTATEUR RESEAU POUR L'EXTERNAL (EX : ETHERNET)"
                                NEW-VMSWITCH -NAME $NOMC -NETADAPTERNAME $ADAPTER -ALLOWMANAGEMENTOS $TRUE
                            } ELSEIF ($TYPE -EQ "INTERNAL" -OR $TYPE -EQ "PRIVATE") {
                                NEW-VMSWITCH -NAME $NOMC -SWITCHTYPE $TYPE
                            } ELSE {
                                Clear-Host
                                Show-Tableau @("TYPE NON VALIDE") -CouleurBordure "Red"
                            }
                        }
                        3 {
                            $NOMC = READ-HOST "NOM DU COMMUTATEUR A SUPPRIMER"
                            REMOVE-VMSWITCH -NAME $NOMC -FORCE
                            Clear-Host
                            Show-Tableau @("COMMUTATEUR $NOMC SUPPRIME") -CouleurBordure "Green"
                        }
                        "B" { Clear-Host; Show-Tableau @("RETOUR MENU AVANCE") -CouleurBordure "Magenta" }
                        "R" { CONTINUE }
                        DEFAULT { Clear-Host; Show-Tableau @("CHOIX INVALIDE") -CouleurBordure "Red" }
                    }
                    IF ($CHOIXC.ToUpper() -NE "B" -and $CHOIXC.ToUpper() -NE "R") { PAUSE }
                } WHILE ($CHOIXC.ToUpper() -NE "B")
            }
            3 {
                DO {
                    Clear-Host
                    Show-Tableau @(
                        "",
                        "GESTION DES SNAPSHOTS (CHECKPOINTS)",
                        "",
                        "1. LISTER LES CHECKPOINTS",
                        "2. CREER UN CHECKPOINT",
                        "3. RESTAURER UN CHECKPOINT",
                        "4. SUPPRIMER UN CHECKPOINT",
                        "",
                        "R. RAFRAICHIR / RECENTRER LE MENU",
                        "B. RETOUR MENU AVANCE"
                    ) -CouleurBordure "DarkBlue" -CouleurTexte "White"
                    $CHOIX_SNAP = READ-HOST "VOTRE CHOIX (1-4, R ou B)"
                    SWITCH ($CHOIX_SNAP.ToUpper()) {
                        1 {
                            $VMNAME = READ-HOST "NOM DE LA VM"
                            $cps = GET-VMCHECKPOINT -VMNAME $VMNAME | SELECT-OBJECT NAME, CREATETIME
                            $lines = @("LISTE DES CHECKPOINTS DE $VMNAME :","")
                            foreach ($cp in $cps) {
                                $lines += ("{0,-25} | Créé le: {1}" -f $cp.Name, $cp.CreationTime)
                            }
                            Clear-Host
                            Show-Tableau $lines -CouleurBordure "Cyan" -CouleurTexte "White"
                        }
                        2 {
                            $VMNAME = READ-HOST "NOM DE LA VM"
                            $CPNAME = READ-HOST "NOM DU CHECKPOINT"
                            CHECKPOINT-VM -VMNAME $VMNAME -SNAPSHOTNAME $CPNAME
                            Clear-Host
                            Show-Tableau @("CHECKPOINT $CPNAME CREE SUR $VMNAME") -CouleurBordure "Green"
                        }
                        3 {
                            $VMNAME = READ-HOST "NOM DE LA VM"
                            $CPNAME = READ-HOST "NOM DU CHECKPOINT"
                            RESTORE-VMCHECKPOINT -VMNAME $VMNAME -NAME $CPNAME
                            Clear-Host
                            Show-Tableau @("CHECKPOINT $CPNAME RESTAURE SUR $VMNAME") -CouleurBordure "Green"
                        }
                        4 {
                            $VMNAME = READ-HOST "NOM DE LA VM"
                            $CPNAME = READ-HOST "NOM DU CHECKPOINT"
                            REMOVE-VMCHECKPOINT -VMNAME $VMNAME -NAME $CPNAME
                            Clear-Host
                            Show-Tableau @("CHECKPOINT $CPNAME SUPPRIME DE $VMNAME") -CouleurBordure "Green"
                        }
                        "B" { Clear-Host; Show-Tableau @("RETOUR MENU AVANCE") -CouleurBordure "Magenta" }
                        "R" { CONTINUE }
                        DEFAULT { Clear-Host; Show-Tableau @("CHOIX INVALIDE") -CouleurBordure "Red" }
                    }
                    IF ($CHOIX_SNAP.ToUpper() -NE "B" -and $CHOIX_SNAP.ToUpper() -NE "R") { PAUSE }
                } WHILE ($CHOIX_SNAP.ToUpper() -NE "B")
            }
            4 {
                DO {
                    Clear-Host
                    Show-Tableau @(
                        "",
                        "GESTION DU STOCKAGE",
                        "",
                        "1. LISTER LES DISQUES DURS VIRTUELS",
                        "2. CREER UN VHDX",
                        "3. SUPPRIMER UN VHDX",
                        "",
                        "R. RAFRAICHIR / RECENTRER LE MENU",
                        "B. RETOUR MENU AVANCE"
                    ) -CouleurBordure "DarkRed" -CouleurTexte "Yellow"
                    $CHOIX_DISK = READ-HOST "VOTRE CHOIX (1-3, R ou B)"
                    SWITCH ($CHOIX_DISK.ToUpper()) {
                        1 {
                            $vhdx = GET-VHD | SELECT-OBJECT PATH, VIRTUALSIZEMB, FILESIZE, ATTACHED
                            $lines = @("LISTE DES DISQUES DURS VIRTUELS :","")
                            foreach ($v in $vhdx) {
                                $lines += ("{0,-50} | Taille: {1,8}MB | Taille fichier: {2,8} | Attaché: {3}" -f $v.Path, $v.VirtualSizeMB, $v.FileSize, $v.Attached)
                            }
                            Clear-Host
                            Show-Tableau $lines -CouleurBordure "Cyan" -CouleurTexte "White"
                        }
                        2 {
                            $PATH = READ-HOST "CHEMIN DU NOUVEAU VHDX"
                            $SIZE = Saisir-Nombre "TAILLE (MB, EX : 40960)"
                            $SIZE_BYTES = [long]($SIZE * 1MB)
                            NEW-VHD -PATH $PATH -SIZEBYTES $SIZE_BYTES -DYNAMIC
                            Clear-Host
                            Show-Tableau @("DISQUE $PATH CREE.") -CouleurBordure "Green"
                        }
                        3 {
                            $PATH = READ-HOST "CHEMIN DU VHDX A SUPPRIMER"
                            REMOVE-ITEM $PATH
                            Clear-Host
                            Show-Tableau @("DISQUE $PATH SUPPRIME.") -CouleurBordure "Green"
                        }
                        "B" { Clear-Host; Show-Tableau @("RETOUR MENU AVANCE") -CouleurBordure "Magenta" }
                        "R" { CONTINUE }
                        DEFAULT { Clear-Host; Show-Tableau @("CHOIX INVALIDE") -CouleurBordure "Red" }
                    }
                    IF ($CHOIX_DISK.ToUpper() -NE "B" -and $CHOIX_DISK.ToUpper() -NE "R") { PAUSE }
                } WHILE ($CHOIX_DISK.ToUpper() -NE "B")
            }
            "B" { Clear-Host; Show-Tableau @("RETOUR MENU PRINCIPAL") -CouleurBordure "Magenta" }
            "R" { CONTINUE }
            DEFAULT { Clear-Host; Show-Tableau @("CHOIX INVALIDE") -CouleurBordure "Red" }
        }
        IF ($CHOIX_AVANCE.ToUpper() -NE "B" -and $CHOIX_AVANCE.ToUpper() -NE "R") { PAUSE }
    } WHILE ($CHOIX_AVANCE.ToUpper() -NE "B")
}

# ========== BOUCLE PRINCIPALE ==========
DO {
    Afficher-MenuDepart
    $CHOIX_DEPART = READ-HOST "VOTRE CHOIX (1, R ou Q)"
    SWITCH ($CHOIX_DEPART.ToUpper()) {
        1 { MENU_AVANCE }
        "Q" { Clear-Host; Show-Tableau @("FIN DU SCRIPT.") -CouleurBordure "Green" }
        "R" { CONTINUE }
        DEFAULT { Clear-Host; Show-Tableau @("CHOIX INVALIDE") -CouleurBordure "Red" }
    }
} WHILE ($CHOIX_DEPART.ToUpper() -NE "Q")
