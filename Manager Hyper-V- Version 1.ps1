# SCRIPT DE GESTION HYPER-V

# --- FONCTIONS UTILITAIRES ---

function AFFICHER-MESSAGE($MESSAGE, $COULEUR = "WHITE") {
    WRITE-HOST $MESSAGE -FOREGROUNDCOLOR $COULEUR
}

function PAUSE {
    AFFICHER-MESSAGE "Cliquer sur Entree pour continuer..." "GRAY"
    $null = READ-HOST
    Clear-Host
}

function SAISIR-NOMBRE($MESSAGE, [int]$MIN = 0, [int]$MAX = [INT32]::MaxValue) {
    DO {
        $VALEUR = READ-HOST $MESSAGE
        IF ($VALEUR -match '^\d+$' -and [int]$VALEUR -ge $MIN -and [int]$VALEUR -le $MAX) { RETURN [int]$VALEUR }
        AFFICHER-MESSAGE "SAISISSEZ UNIQUEMENT UN NOMBRE VALIDE (ENTRE $MIN ET $MAX)" "RED"
    } WHILE ($TRUE)
}

function SAISIR-CHOIX($PROMPT, [string[]]$CHOIX_VALIDE) {
    DO {
        $CHOIX = (READ-HOST $PROMPT).TOUPPER()
        IF ($CHOIX_VALIDE -contains $CHOIX) { RETURN $CHOIX }
        AFFICHER-MESSAGE "CHOIX INVALIDE. VEUILLEZ SELECTIONNER PARMI: $($CHOIX_VALIDE -join ', ')" "RED"
    } WHILE ($TRUE)
}

function DETECT-TYPEOS($VMName) {
    $result = @{}
    if ($VMName -match "WINDOWS|WIN|CORE|SERVEUR|SERVER") {
        $result.cpu = 2
        $result.msg = "TYPE DETECTE : WINDOWS/SERVER/CORE - PROPOSITION : 2 COEURS"
    } elseif ($VMName -match "LINUX|DEBIAN|UBUNTU|CENTOS|FREEBSD|FIREWALL|PF|PFSENSE|OPENWRT|ALPINE|VYOS|OPNSENSE") {
        $result.cpu = 1
        $result.msg = "TYPE DETECTE : LINUX/BSD/FIREWALL/AUTRES - PROPOSITION : 1 COEUR"
    } else {
        $result.cpu = 1
        $result.msg = "OS NON DETECTE, PAR DEFAUT : 1 COEUR"
    }
    return $result
}

function AFFICHER-TABLEAU($ENTETES, $DONNEES) {
    $longueurs = @{}
    
    # Calcul des longueurs maximales pour chaque colonne
    foreach ($entete in $ENTETES) {
        $longueurs[$entete] = $entete.Length
    }
    foreach ($donnee in $DONNEES) {
        foreach ($entete in $ENTETES) {
            $longueur_valeur = "$($donnee.($entete))".Length
            if ($longueur_valeur -gt $longueurs[$entete]) {
                $longueurs[$entete] = $longueur_valeur
            }
        }
    }
    
    # Bordure superieure
    $bordure = "+-" + (($ENTETES | ForEach-Object { "-" * $longueurs[$_] }) -join "-+-") + "-+"
    AFFICHER-MESSAGE $bordure "CYAN"
    
    # En-tetes
    $ligne_entete = "| " + (($ENTETES | ForEach-Object { "{0,-$($longueurs[$_])}" -f $_ }) -join " | ") + " |"
    AFFICHER-MESSAGE $ligne_entete "CYAN"
    
    # Bordure de separation
    $bordure_sep = "+-" + (($ENTETES | ForEach-Object { "-" * $longueurs[$_] }) -join "-+-") + "-+"
    AFFICHER-MESSAGE $bordure_sep "CYAN"
    
    # Donnees
    foreach ($donnee in $DONNEES) {
        $ligne_donnee = "| " + (($ENTETES | ForEach-Object { "{0,-$($longueurs[$_])}" -f "$($donnee.($_))" }) -join " | ") + " |"
        AFFICHER-MESSAGE $ligne_donnee "WHITE"
    }
    
    # Bordure inferieure
    AFFICHER-MESSAGE $bordure "CYAN"
}


# --- GESTION DES MACHINES VIRTUELLES ---
function LISTER-VM {
    Clear-Host
    AFFICHER-MESSAGE "LISTE DES MACHINES VIRTUELLES :" "CYAN"
    TRY {
        $vms = GET-VM | SELECT-OBJECT NAME, STATE, CPUUSAGE, MEMORYASSIGNED, UPTIME, VERSION, DYNAMICMEMORYENABLED

        IF (-NOT $vms) {
            AFFICHER-MESSAGE "AUCUNE VM TROUVEE." "RED"
        } ELSE {
            $donnees_vm = @()
            foreach ($vm in $vms) {
                $ram = "{0:N0} MB" -f ([Math]::ROUND($vm.MEMORYASSIGNED / 1MB))
                $uptime = if ($vm.STATE -eq "RUNNING") {
                    $days = [Math]::FLOOR($vm.UPTIME.TOTALDAYS)
                    $hours = $vm.UPTIME.HOURS
                    $minutes = $vm.UPTIME.MINUTES
                    "$($days)J $($hours)H $($minutes)M"
                } else { "N/A" }
                
                $donnees_vm += [PSCustomObject]@{
                    NOM                = $vm.NAME
                    ETAT               = $vm.STATE
                    CPU                = "$($vm.CPUUSAGE)%"
                    RAM                = $ram
                    UPTIME             = $uptime
                    VERSION            = $vm.VERSION
                    "MEMOIRE DYNAMIQUE" = if ($vm.DYNAMICMEMORYENABLED) { "Oui" } else { "Non" }
                }
            }
            $entetes_vm = "NOM", "ETAT", "CPU", "RAM", "UPTIME", "VERSION", "MEMOIRE DYNAMIQUE"
            AFFICHER-TABLEAU $entetes_vm $donnees_vm
        }
    } CATCH {
        AFFICHER-MESSAGE "ERREUR : IMPOSSIBLE DE LISTER LES VMS. VERIFIEZ QUE LES MODULES HYPER-V SONT INSTALLEES." "RED"
    }
    PAUSE
}

function DEMARRER-VM {
    Clear-Host
    AFFICHER-MESSAGE "VMS DISPONIBLES A DEMARRER :" "CYAN"
    $vms = GET-VM | WHERE-OBJECT { $_.STATE -NE "RUNNING" }
    
    IF ($vms.COUNT -EQ 0) {
        AFFICHER-MESSAGE "AUCUNE VM N'EST ARRETEE." "RED"
        PAUSE
        RETURN
    }

    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm
    
    $CHOIX_VM = SAISIR-NOMBRE "NUMERO DE LA VM A DEMARRER" 1 $vms.COUNT
    $NAME = $vms[$CHOIX_VM-1].NAME
    TRY {
        START-VM -NAME $NAME -ERRORACTION STOP
        AFFICHER-MESSAGE "VM $NAME DEMARREE." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS DU DEMARRAGE DE LA VM : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function ARRETER-VM {
    Clear-Host
    AFFICHER-MESSAGE "VMS DISPONIBLES A ARRETER :" "CYAN"
    $vms = GET-VM | WHERE-OBJECT { $_.STATE -EQ "RUNNING" }

    IF ($vms.COUNT -EQ 0) {
        AFFICHER-MESSAGE "AUCUNE VM N'EST EN COURS D'EXECUTION." "RED"
        PAUSE
        RETURN
    }

    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO DE LA VM A ARRETER" 1 $vms.COUNT
    $NAME = $vms[$CHOIX_VM-1].NAME
    $confirm = READ-HOST "VOULEZ-VOUS VRAIMENT ARRETER LA VM $NAME ? (O/N)"
    IF ($confirm.TOUPPER() -EQ "O") {
        TRY {
            STOP-VM -NAME $NAME -FORCE -ERRORACTION STOP
            AFFICHER-MESSAGE "VM $NAME ARRETEE." "GREEN"
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS DE L'ARRET DE LA VM : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    } ELSE {
        AFFICHER-MESSAGE "OPERATION ANNULEE." "YELLOW"
    }
    PAUSE
}

function REDEMARRER-VM {
    Clear-Host
    AFFICHER-MESSAGE "VMS DISPONIBLES A REDEMARRER :" "CYAN"
    $vms = GET-VM | WHERE-OBJECT { $_.STATE -EQ "RUNNING" }

    IF ($vms.COUNT -EQ 0) {
        AFFICHER-MESSAGE "AUCUNE VM N'EST EN COURS D'EXECUTION." "RED"
        PAUSE
        RETURN
    }

    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO DE LA VM A REDEMARRER" 1 $vms.COUNT
    $NAME = $vms[$CHOIX_VM-1].NAME
    $confirm = READ-HOST "VOULEZ-VOUS VRAIMENT REDEMARRER LA VM $NAME ? (O/N)"
    IF ($confirm.TOUPPER() -EQ "O") {
        TRY {
            RESTART-VM -NAME $NAME -FORCE -ERRORACTION STOP
            AFFICHER-MESSAGE "VM $NAME REDEMARREE." "GREEN"
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS DU REDEMARRAGE DE LA VM : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    } ELSE {
        AFFICHER-MESSAGE "OPERATION ANNULEE." "YELLOW"
    }
    PAUSE
}

function SUPPRIMER-VM {
    Clear-Host
    AFFICHER-MESSAGE "VMS DISPONIBLES A SUPPRIMER :" "CYAN"
    $vms = GET-VM | WHERE-OBJECT { $_.STATE -EQ "OFF" }

    IF ($vms.COUNT -EQ 0) {
        AFFICHER-MESSAGE "AUCUNE VM N'EST ARRETEE POUR ETRE SUPPRIMEE." "RED"
        PAUSE
        RETURN
    }
    
    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO DE LA VM A SUPPRIMER" 1 $vms.COUNT
    $NAME = $vms[$CHOIX_VM-1].NAME
    $confirm = READ-HOST "ETES-VOUS CERTAIN DE VOULOIR SUPPRIMER DEFINITIVEMENT LA VM '$NAME' ET SES DISQUES ? TAPEZ LE NOM DE LA VM POUR CONFIRMER"
    IF ($confirm -EQ $NAME) {
        TRY {
            $vm = GET-VM -NAME $NAME -ERRORACTION STOP
            $vhdPath = $vm.HARDDRIVES.PATH
            STOP-VM -NAME $NAME -FORCE
            REMOVE-VM -NAME $NAME -FORCE -ERRORACTION STOP
            IF (TEST-PATH $vhdPath) {
                REMOVE-ITEM -PATH $vhdPath -FORCE
            }
            AFFICHER-MESSAGE "VM $NAME ET SON DISQUE SUPPRIMES." "GREEN"
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS DE LA SUPPRESSION DE LA VM : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    } ELSE {
        AFFICHER-MESSAGE "OPERATION ANNULEE." "YELLOW"
    }
    PAUSE
}

function EXPORTER-VM {
    Clear-Host
    AFFICHER-MESSAGE "VMS DISPONIBLES A EXPORTER :" "CYAN"
    $vms = GET-VM
    
    if ($vms.COUNT -EQ 0) {
        AFFICHER-MESSAGE "AUCUNE VM TROUVEE A EXPORTER." "RED"
        PAUSE
        RETURN
    }
    
    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm
    $CHOIX_MULTI = SAISIR-CHOIX "VOULEZ-VOUS EXPORTER UNE OU PLUSIEURS VMS ? (U/P)" @("U", "P")
    $VM_SELECTIONNEES = @()
    if ($CHOIX_MULTI -eq "U") {
        $CHOIX_VM = SAISIR-NOMBRE "NUMERO DE LA VM A EXPORTER" 1 $vms.COUNT
        $VM_SELECTIONNEES += $vms[$CHOIX_VM-1].NAME
    } else {
        $CHOIX_NUMEROS = (READ-HOST "NUMEROS DES VMS A EXPORTER (EX: 1,3,5)") -split ','
        foreach ($numero in $CHOIX_NUMEROS) {
            $num = [int]$numero
            if ($num -ge 1 -and $num -le $vms.Count) {
                $VM_SELECTIONNEES += $vms[$num-1].NAME
            }
        }
    }
    if ($VM_SELECTIONNEES.Count -eq 0) {
        AFFICHER-MESSAGE "AUCUNE VM SELECTIONNEE. OPERATION ANNULEE." "YELLOW"
        PAUSE
        RETURN
    }

    $EXPORT_PATH = READ-HOST "CHEMIN de destination pour l'export (EX : C:\EXPORT_VMS)"
    IF (-NOT (TEST-PATH $EXPORT_PATH)) {
        NEW-ITEM -PATH $EXPORT_PATH -ITEMTYPE DIRECTORY | OUT-NULL
    }

    $CHOIX_COMPRESSION = SAISIR-CHOIX "COMPRESSER l'exportation ? (O/N)" @("O", "N")
    if ($CHOIX_COMPRESSION -eq "O") {
        AFFICHER-MESSAGE "CHOIX du format de compression :`n1. ZIP`n2. ZPAQ (REQUIERT PEAZIP)"
        # J'ai mis un message qui ne donne pas de choix pour le .zip pour ne pas faire d'erreur
        AFFICHER-MESSAGE "Le format ZIP est indisponible. Veuillez choisir 2 pour ZPAQ." "RED"
        $CHOIX_FORMAT = SAISIR-NOMBRE "VOTRE CHOIX (2)" 2 2
        
        $ARCHIVE_FORMAT = "ZPAQ"
    }
    
    foreach ($vmName in $VM_SELECTIONNEES) {
        $EXPORT_VM_PATH = JOIN-PATH -PATH $EXPORT_PATH -CHILDpath $vmName
        AFFICHER-MESSAGE "EXPORTATION de $vmName vers $EXPORT_VM_PATH..." "YELLOW"
        try {
            if (TEST-PATH $EXPORT_VM_PATH) {
                AFFICHER-MESSAGE "LE REPERTOIRE D'EXPORTATION EXISTE DEJA. SUPPRESSION..." "YELLOW"
                REMOVE-ITEM -PATH $EXPORT_VM_PATH -RECURSE -FORCE
            }
            
            EXPORT-VM -NAME $vmName -PATH $EXPORT_VM_PATH -ERRORACTION STOP

            AFFICHER-MESSAGE "EXPORT de $vmName TERMINE !" "GREEN"

            if ($CHOIX_COMPRESSION -eq "O") {
                try {
                    if ($ARCHIVE_FORMAT -eq "ZPAQ") {
                        $peazipPath = "C:\Program Files\PeaZip\peazip.exe"
                        if (TEST-PATH $peazipPath) {
                            $cheminArchive = JOIN-PATH -PATH $EXPORT_PATH -CHILDpath "$vmName.zpaq"
                            & $peazipPath -add "$EXPORT_VM_PATH" "$cheminArchive"
                            PAUSE
                        } else {
                            AFFICHER-MESSAGE "ERREUR : PEAZIP NON TROUVE. Installation requise pour ZPAQ." "RED"
                        }
                    }
                } CATCH {
                    AFFICHER-MESSAGE "ERREUR LORS de la compression de $vmName : $($_.EXCEPTION.MESSAGE)" "RED"
                }
            }
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS de l'export de la VM $vmName : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    }
    PAUSE
}


function CREER-VM {
    Clear-Host
    AFFICHER-MESSAGE "CREATION de MACHINE VIRTUELLE" "CYAN"
    AFFICHER-MESSAGE "--------------------------" "CYAN"
    AFFICHER-MESSAGE "ENTREZ LES INFORMATIONS POUR LA VM" "CYAN"

    $NAME = READ-HOST "NOM de la nouvelle VM"
    $MEMOIRE_GO = SAISIR-NOMBRE "MEMOIRE RAM en GO (4-128)" 4 128
    $MEMOIRE_BYTES = [long]($MEMOIRE_GO * 1GB)
    $MIN_CPU = 1
    $MAX_CPU = [Math]::Min(12, (GET-VMHOST).LOGICALPROCESSORCOUNT)
    $osInfo = DETECT-TYPEOS $NAME
    AFFICHER-MESSAGE $osInfo.MSG "CYAN"
    $CPU = SAISIR-NOMBRE "NOMBRE de coeurs de processeur (entre $MIN_CPU et $MAX_CPU)" $MIN_CPU $MAX_CPU
    $VHDX_PATH = READ-HOST "CHEMIN COMPLET du disque VHDX (EX : D:\HYPERV\$NAME.VHDX)"
    $TAILLE_VHDX_GO = SAISIR-NOMBRE "TAILLE du disque VHDX en GO (20-2048)" 20 2048
    $TAILLE_VHDX_BYTES = [long]($TAILLE_VHDX_GO * 1GB)
    
    AFFICHER-MESSAGE "LISTE des switchs virtuels disponibles" "CYAN"
    $switchs = GET-VMSWITCH | SELECT-OBJECT NAME, SWITCHTYPE
    IF ($switchs.Count -EQ 0) {
        AFFICHER-MESSAGE "AUCUN SWITCH VIRTUEL TROUVE. VEUILLEZ EN CREER UN AVANT de continuer." "RED"
        PAUSE
        RETURN
    }
    $donnees_switch = @()
    for ($i=0; $i -lt $switchs.Count; $i++) {
        $donnees_switch += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $switchs[$i].NAME
            "TYPE" = $switchs[$i].SWITCHTYPE
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM", "TYPE") $donnees_switch
    $CHOIX_SWITCH = SAISIR-NOMBRE "NUMERO du switch virtuel" 1 $switchs.Count
    $SWITCH_NOM = $switchs[$CHOIX_SWITCH-1].NAME


    AFFICHER-MESSAGE "VOULEZ-VOUS ATTRIBUER une ISO pour l'installation ?`n1. OUI`n2. NON"
    $CHOIX_ISO = SAISIR-NOMBRE "VOTRE CHOIX (1-2)" 1 2
    $ISO_PATH = ""
    IF ($CHOIX_ISO -EQ 1) {
        $ISO_PATH = READ-HOST "CHEMIN COMPLET de l'image ISO (EX : C:\ISO\WIN.ISO)"
        IF (-NOT (TEST-PATH $ISO_PATH)) {
            AFFICHER-MESSAGE "CHEMIN ISO INVALIDE. LA VM SERA CREE SANS ISO." "YELLOW"
            $ISO_PATH = ""
        }
    }

    $confirm = READ-HOST "CONFIRMER la creation de la VM '$NAME' ? (O/N)"
    IF ($confirm.TOUPPER() -NE "O") {
        AFFICHER-MESSAGE "CREATION ANNULEE." "RED"
        PAUSE
        RETURN
    }

    Clear-Host
    AFFICHER-MESSAGE "CREATION du disque VHDX..." "BLUE"
    TRY {
        NEW-VHD -PATH $VHDX_PATH -SIZEBYTES $TAILLE_VHDX_BYTES -DYNAMIC -ERRORACTION STOP
        AFFICHER-MESSAGE "DISQUE VHDX CREE AVEC SUCCES." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS de la creation du disque VHDX : $($_.EXCEPTION.MESSAGE)" "RED"
        PAUSE
        RETURN
    }
    PAUSE

    Clear-Host
    AFFICHER-MESSAGE "CREATION de la VM..." "BLUE"
    TRY {
        $vm = NEW-VM -NAME $NAME -MEMORYSTARTUPBYTES $MEMOIRE_BYTES -VHDPath $VHDX_PATH -SWITCHNAME $SWITCH_NOM -ERRORACTION STOP
        SET-VM -NAME $NAME -PROCESSORCOUNT $CPU
        IF (-NOT [string]::ISNULLORWHITESPACE($ISO_PATH)) {
            SET-VMDVDDRIVE -VMNAME $NAME -PATH $ISO_PATH
        }
        START-VM -NAME $NAME
        AFFICHER-MESSAGE "VM '$NAME' CREEE ET DEMARREE AVEC SUCCES !" "GREEN"
    } CATCH {
    AFFICHER-MESSAGE "ERREUR LORS de la creation de la VM : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function CONFIGURER-VM {
    Clear-Host
    AFFICHER-MESSAGE "VMS disponibles a configurer :" "CYAN"
    $vms = GET-VM
    IF ($vms.COUNT -EQ 0) {
        AFFICHER-MESSAGE "AUCUNE VM TROUVEE a configurer." "RED"
        PAUSE
        RETURN
    }

    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM a configurer" 1 $vms.COUNT
    $VMNAME = $vms[$CHOIX_VM-1].NAME

    DO {
        Clear-Host
        AFFICHER-MESSAGE "CONFIGURATION de la VM : $VMNAME" "DARKMAGENTA"
        AFFICHER-MESSAGE "1. ACTIVER/DESACTIVER la memoire dynamique" "WHITE"
        AFFICHER-MESSAGE "2. MODIFIER le nombre de coeurs" "WHITE"
        AFFICHER-MESSAGE "3. MODIFIER la memoire allouee" "WHITE"
        AFFICHER-MESSAGE "R. RAFRAICHIR / RECENTER le menu" "WHITE"
        AFFICHER-MESSAGE "B. RETOUR menu gestion VM" "WHITE"
        $CHOIX = READ-HOST "VOTRE CHOIX (1-3, R ou B)"

        SWITCH ($CHOIX.TOUPPER()) {
            1 {
                TRY {
                    $VM = GET-VM -NAME $VMNAME -ERRORACTION STOP
                    $currentStatus = $VM.DYNAMICMEMORYENABLED
                    $newStatus = if ($currentStatus) { $false } else { $true }
                    SET-VMMEMORY -VM $VM -DYNAMICMEMORYENABLED $newStatus -ERRORACTION STOP
                    AFFICHER-MESSAGE "MEMOIRE dynamique de '$VMNAME' changee de $currentStatus a $newStatus." "GREEN"
                } CATCH {
                    AFFICHER-MESSAGE "ERREUR LORS de la modification de la memoire dynamique : $($_.EXCEPTION.MESSAGE)" "RED"
                }
                PAUSE
            }
            2 {
                TRY {
                    $VM = GET-VM -NAME $VMNAME -ERRORACTION STOP
                    $MIN_CPU = 1
                    $MAX_CPU = [Math]::Min(12, (GET-VMHOST).LOGICALPROCESSORCOUNT)
                    $currentCPU = $VM.PROCESSORCOUNT
                    $NEW_CPU = SAISIR-NOMBRE "ENTREZ le nouveau nombre de coeurs (ACTUEL: $currentCPU)" $MIN_CPU $MAX_CPU
                    SET-VM -NAME $VMNAME -PROCESSORCOUNT $NEW_CPU -ERRORACTION STOP
                    AFFICHER-MESSAGE "NOMBRE de coeurs de '$VMNAME' change a $NEW_CPU." "GREEN"
                } CATCH {
                    AFFICHER-MESSAGE "ERREUR LORS de la modification des coeurs : $($_.EXCEPTION.MESSAGE)" "RED"
                }
                PAUSE
            }
            3 {
                TRY {
                    $VM = GET-VM -NAME $VMNAME -ERRORACTION STOP
                    $currentMemGB = $VM.MEMORYSTARTUP / 1GB
                    $NEW_MEM_GO = SAISIR-NOMBRE "ENTREZ la nouvelle memoire en GO (ACTUEL: $currentMemGB)" 4 128
                    $NEW_MEM_BYTES = [long]($NEW_MEM_GO * 1GB)
                    
                    if ($VM.VERSION -like "10.0") { # VM Generation 1
                        SET-VMMEMORY -VM $VM -STARTUPBYTES $NEW_MEM_BYTES -ERRORACTION STOP
                    } else { # VM Generation 2
                        SET-VMMEMORY -VM $VM -MEMORYSTARTUPBYTES $NEW_MEM_BYTES -ERRORACTION STOP
                    }
                    AFFICHER-MESSAGE "MEMOIRE de '$VMNAME' changee a $NEW_MEM_GO GO." "GREEN"
                } CATCH {
                    AFFICHER-MESSAGE "ERREUR LORS de la modification de la memoire : $($_.EXCEPTION.MESSAGE)" "RED"
                }
                PAUSE
            }
            "B" { RETURN }
            "R" { CONTINUE }
            DEFAULT { AFFICHER-MESSAGE "CHOIX INVALIDE" "RED"; PAUSE }
        }
    } WHILE ($TRUE)
}

# --- GESTION DES COMMUTATEURS ---
function LISTER-COMMUTATEUR {
    Clear-Host
    AFFICHER-MESSAGE "LISTE des commutateurs Hyper-V :" "CYAN"
    $switchs = GET-VMSWITCH | SELECT-OBJECT NAME, SWITCHTYPE
    IF (-NOT $switchs) {
        AFFICHER-MESSAGE "AUCUN COMMUTATEUR TROUVE." "RED"
    } ELSE {
        $donnees_switch = @()
        foreach ($switch in $switchs) {
            $donnees_switch += [PSCustomObject]@{
                NOM = $switch.NAME
                TYPE = $switch.SWITCHTYPE
            }
        }
        $entetes_switch = "NOM", "TYPE"
        AFFICHER-TABLEAU $entetes_switch $donnees_switch
    }
    PAUSE
}

function CREER-COMMUTATEUR {
    Clear-Host
    $NOMC = READ-HOST "NOM du commutateur (EX : COM_LAN_PRENOM)"
    $TYPE = SAISIR-CHOIX "TYPE (EXTERNAL, INTERNAL, PRIVATE)" @("EXTERNAL", "INTERNAL", "PRIVATE")

    TRY {
        IF ($TYPE -EQ "EXTERNAL") {
            $ADAPTERS = GET-NETADAPTER | WHERE-OBJECT {$_.STATUS -EQ "UP"} | SELECT-OBJECT -EXPAND NAME
            AFFICHER-MESSAGE "ADAPTATEURS disponibles :" "CYAN"
            $donnees_adapt = @()
            for ($i=0; $i -lt $ADAPTERS.Count; $i++) {
                $donnees_adapt += [PSCustomObject]@{
                    "NUMERO" = $i+1
                    "NOM" = $ADAPTERS[$i]
                }
            }
            AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_adapt
            $CHOIX_ADAPTER = SAISIR-NOMBRE "NUMERO de l'adaptateur reseau pour l'external" 1 $ADAPTERS.COUNT
            $ADAPTER = $ADAPTERS[$CHOIX_ADAPTER-1]
            NEW-VMSWITCH -NAME $NOMC -NETADAPTERNAME $ADAPTER -ALLOWMANAGEMENTOS $TRUE -ERRORACTION STOP
        } ELSE {
            NEW-VMSWITCH -NAME $NOMC -SWITCHTYPE $TYPE -ERRORACTION STOP
        }
        AFFICHER-MESSAGE "COMMUTATEUR $NOMC CREE." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS de la creation du commutateur : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function SUPPRIMER-COMMUTATEUR {
    Clear-Host
    $switchs = GET-VMSWITCH
    IF ($switchs.Count -EQ 0) {
        AFFICHER-MESSAGE "AUCUN COMMUTATEUR TROUVE." "RED"
        PAUSE
        RETURN
    }
    AFFICHER-MESSAGE "COMMUTATEURS disponibles a supprimer :" "CYAN"
    $donnees_switch = @()
    for ($i=0; $i -lt $switchs.Count; $i++) {
        $donnees_switch += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $switchs[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_switch
    
    $CHOIX_COMM = SAISIR-NOMBRE "NUMERO du commutateur a supprimer" 1 $switchs.Count
    $NOMC = $switchs[$CHOIX_COMM-1].Name
    $confirm = READ-HOST "ETES-VOUS CERTAIN de vouloir supprimer le commutateur '$NOMC' ? TAPEZ le nom pour confirmer"
    IF ($confirm -EQ $NOMC) {
        TRY {
            REMOVE-VMSWITCH -NAME $NOMC -FORCE -ERRORACTION STOP
            AFFICHER-MESSAGE "COMMUTATEUR $NOMC SUPPRIME." "GREEN"
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS de la suppression du commutateur : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    } ELSE {
        AFFICHER-MESSAGE "OPERATION ANNULEE." "YELLOW"
    }
    PAUSE
}

# --- GESTION DES CHECKPOINTS ---
function LISTER-CHECKPOINTS {
    Clear-Host
    AFFICHER-MESSAGE "VMS disponibles :" "CYAN"
    $vms = GET-VM
    if ($vms.Count -eq 0) { AFFICHER-MESSAGE "AUCUNE VM TROUVEE." "RED"; PAUSE; RETURN }
    
    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM pour laquelle lister les checkpoints" 1 $vms.COUNT
    $VMNAME = $vms[$CHOIX_VM-1].NAME

    TRY {
        $cps = GET-VMCHECKPOINT -VMNAME $VMNAME -ERRORACTION STOP | SELECT-OBJECT NAME, CREATETIME
        AFFICHER-MESSAGE "LISTE des checkpoints de $VMNAME :" "CYAN"
        IF (-NOT $cps) {
            AFFICHER-MESSAGE "AUCUN CHECKPOINT TROUVE pour cette VM." "RED"
        } ELSE {
            $donnees_cp = @()
            foreach ($cp in $cps) {
                $donnees_cp += [PSCustomObject]@{
                    NOM = $cp.NAME
                    DATE = $cp.CREATETIME.ToString("yyyy-MM-dd HH:mm:ss")
                }
            }
            $entetes_cp = "NOM", "DATE"
            AFFICHER-TABLEAU $entetes_cp $donnees_cp
        }
    } CATCH {
        AFFICHER-MESSAGE "ERREUR : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function CREER-CHECKPOINT {
    Clear-Host
    AFFICHER-MESSAGE "VMS disponibles :" "CYAN"
    $vms = GET-VM
    if ($vms.Count -eq 0) { AFFICHER-MESSAGE "AUCUNE VM TROUVEE." "RED"; PAUSE; RETURN }
    
    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM pour laquelle creer un checkpoint" 1 $vms.COUNT
    $VMNAME = $vms[$CHOIX_VM-1].NAME

    $CPNAME = READ-HOST "NOM du nouveau checkpoint"
    TRY {
        CHECKPOINT-VM -VMNAME $VMNAME -SNAPSHOTNAME $CPNAME -ERRORACTION STOP
        AFFICHER-MESSAGE "CHECKPOINT '$CPNAME' CREE pour la VM '$VMNAME'." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS de la creation du checkpoint : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function RESTAURER-CHECKPOINT {
    Clear-Host
    AFFICHER-MESSAGE "VMS disponibles :" "CYAN"
    $vms = GET-VM
    if ($vms.Count -eq 0) { AFFICHER-MESSAGE "AUCUNE VM TROUVEE." "RED"; PAUSE; RETURN }
    
    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm
    
    $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM a restaurer" 1 $vms.COUNT
    $VMNAME = $vms[$CHOIX_VM-1].NAME

    $cps = GET-VMCHECKPOINT -VMNAME $VMNAME
    if ($cps.Count -eq 0) {
        AFFICHER-MESSAGE "AUCUN CHECKPOINT TROUVE pour la VM '$VMNAME'." "RED"
        PAUSE
        RETURN
    }
    
    AFFICHER-MESSAGE "CHECKPOINTS disponibles pour la VM '$VMNAME' :" "CYAN"
    $donnees_cp = @()
    for ($i = 0; $i -lt $cps.Count; $i++) {
        $donnees_cp += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $cps[$i].Name
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_cp
    
    $CHOIX_CP = SAISIR-NOMBRE "NUMERO du checkpoint a restaurer" 1 $cps.Count
    $CPNAME = $cps[$CHOIX_CP-1].Name

    $confirm = READ-HOST "ETES-VOUS CERTAIN de vouloir restaurer le checkpoint '$CPNAME' pour la VM '$VMNAME' ? CELA PERDRA TOUTES LES DONNEES DEPUIS le dernier point. TAPEZ le nom du checkpoint pour confirmer"
    IF ($confirm -EQ $CPNAME) {
        TRY {
            RESTORE-VMCHECKPOINT -VMNAME $VMNAME -NAME $CPNAME -FORCE -ERRORACTION STOP
            AFFICHER-MESSAGE "VM '$VMNAME' RESTAUREE au checkpoint '$CPNAME'." "GREEN"
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS de la restauration du checkpoint : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    } ELSE {
        AFFICHER-MESSAGE "OPERATION ANNULEE." "YELLOW"
    }
    PAUSE
}

function SUPPRIMER-CHECKPOINT {
    Clear-Host
    AFFICHER-MESSAGE "VMS disponibles :" "CYAN"
    $vms = GET-VM
    if ($vms.Count -eq 0) { AFFICHER-MESSAGE "AUCUNE VM TROUVEE." "RED"; PAUSE; RETURN }
    
    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM" 1 $vms.COUNT
    $VMNAME = $vms[$CHOIX_VM-1].NAME

    $cps = GET-VMCHECKPOINT -VMNAME $VMNAME
    if ($cps.Count -eq 0) {
        AFFICHER-MESSAGE "AUCUN CHECKPOINT TROUVE pour la VM '$VMNAME'." "RED"
        PAUSE
        RETURN
    }
    
    AFFICHER-MESSAGE "CHECKPOINTS disponibles pour la VM '$VMNAME' :" "CYAN"
    $donnees_cp = @()
    for ($i = 0; $i -lt $cps.Count; $i++) {
        $donnees_cp += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $cps[$i].Name
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_cp
    
    $CHOIX_CP = SAISIR-NOMBRE "NUMERO du checkpoint a supprimer" 1 $cps.Count
    $CPNAME = $cps[$CHOIX_CP-1].Name

    $confirm = READ-HOST "ETES-VOUS CERTAIN de vouloir supprimer le checkpoint '$CPNAME' de la VM '$VMNAME' ? TAPEZ le nom du checkpoint pour confirmer"
    IF ($confirm -EQ $CPNAME) {
        TRY {
            REMOVE-VMCHECKPOINT -VMNAME $VMNAME -NAME $CPNAME -FORCE -ERRORACTION STOP
            AFFICHER-MESSAGE "CHECKPOINT '$CPNAME' SUPPRIME." "GREEN"
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS de la suppression du checkpoint : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    } ELSE {
        AFFICHER-MESSAGE "OPERATION ANNULEE." "YELLOW"
    }
    PAUSE
}

# --- GESTION DU STOCKAGE ---
function GERER-STOCKAGE {
    DO {
        Clear-Host
        AFFICHER-MESSAGE "GESTION DU STOCKAGE" "DARKMAGENTA"
        AFFICHER-MESSAGE "1. LISTER les disques VHDX" "WHITE"
        AFFICHER-MESSAGE "2. CREER un disque VHDX" "WHITE"
        AFFICHER-MESSAGE "3. SUPPRIMER un disque VHDX" "WHITE"
        AFFICHER-MESSAGE "4. REDIMENSIONNER un disque VHDX" "WHITE"
        AFFICHER-MESSAGE "5. ATTACHER un disque a une VM" "WHITE"
        AFFICHER-MESSAGE "6. DETACHER un disque d'une VM" "WHITE"
        AFFICHER-MESSAGE "R. RAFRAICHIR / RECENTER le menu" "WHITE"
        AFFICHER-MESSAGE "B. RETOUR menu avance" "WHITE"
        $CHOIX_STOCK = READ-HOST "VOTRE CHOIX (1-6, R ou B)"

        SWITCH ($CHOIX_STOCK.TOUPPER()) {
            1 { LISTER-VHDX }
            2 { CREER-VHDX }
            3 { SUPPRIMER-VHDX }
            4 { REDIMENSIONNER-VHDX }
            5 { ATTACHER-VHDX }
            6 { DETACHER-VHDX }
            "B" { Clear-Host; AFFICHER-MESSAGE "RETOUR menu avance" "MAGENTA"; RETURN }
            "R" { CONTINUE }
            DEFAULT { AFFICHER-MESSAGE "CHOIX INVALIDE" "RED"; PAUSE }
        }
    } WHILE ($TRUE)
}

function LISTER-VHDX {
    Clear-Host
    AFFICHER-MESSAGE "LISTE des disques virtuels (.VHDX) sur ce serveur" "CYAN"
    $vhds = GET-VHD | SELECT-OBJECT PATH, VHDTYPE, FILESIZE, SIZE
    IF (-NOT $vhds) {
        AFFICHER-MESSAGE "AUCUN DISQUE VIRTUEL TROUVE." "RED"
    } ELSE {
        $donnees_vhdx = @()
        foreach ($vhd in $vhds) {
            $donnees_vhdx += [PSCustomObject]@{
                CHEMIN = $vhd.PATH
                TYPE = $vhd.VHDTYPE
                "TAILLE REELLE" = "{0:N2} GB" -f ($vhd.FILESIZE / 1GB)
                "TAILLE MAX" = "{0:N2} GB" -f ($vhd.SIZE / 1GB)
            }
        }
        $entetes_vhdx = "CHEMIN", "TYPE", "TAILLE REELLE", "TAILLE MAX"
        AFFICHER-TABLEAU $entetes_vhdx $donnees_vhdx
    }
    PAUSE
}

function CREER-VHDX {
    Clear-Host
    $PATH = READ-HOST "CHEMIN COMPLET du nouveau disque VHDX (EX: C:\DISQUES\NOUVEAU.VHDX)"
    $TAILLE = SAISIR-NOMBRE "TAILLE du disque en GO (EX: 50)" 1 2048
    $TYPE = SAISIR-CHOIX "TYPE de disque (DYNAMIC, FIXED)" @("DYNAMIC", "FIXED")
    TRY {
        $SIZE_BYTES = [long]($TAILLE * 1GB)
        NEW-VHD -PATH $PATH -SIZEBYTES $SIZE_BYTES -VHDTYPE $TYPE -ERRORACTION STOP
        AFFICHER-MESSAGE "DISQUE VHDX CREE : $PATH" "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS de la creation du VHDX : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function SUPPRIMER-VHDX {
    Clear-Host
    $vhds = GET-VHD
    if ($vhds.Count -eq 0) {
        AFFICHER-MESSAGE "AUCUN DISQUE VIRTUEL TROUVE a supprimer." "RED"
        PAUSE
        RETURN
    }
    AFFICHER-MESSAGE "DISQUES virtuels disponibles a supprimer :" "CYAN"
    $donnees_vhdx = @()
    for ($i = 0; $i -lt $vhds.Count; $i++) {
        $donnees_vhdx += [PSCustomObject]@{
            "NUMERO" = $i+1
            "CHEMIN" = $vhds[$i].PATH
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "CHEMIN") $donnees_vhdx
    
    $CHOIX_VHD = SAISIR-NOMBRE "NUMERO du disque a supprimer" 1 $vhds.Count
    $PATH = $vhds[$CHOIX_VHD-1].Path
    $confirm = READ-HOST "ETES-VOUS CERTAIN de vouloir supprimer definitivement le disque '$PATH' ? TAPEZ le nom pour confirmer"
    IF ($confirm -EQ (SPLIT-PATH $PATH -LEAF)) {
        TRY {
            REMOVE-ITEM -PATH $PATH -FORCE -ERRORACTION STOP
            AFFICHER-MESSAGE "DISQUE VHDX SUPPRIME." "GREEN"
        } CATCH {
            AFFICHER-MESSAGE "ERREUR LORS de la suppression du VHDX : $($_.EXCEPTION.MESSAGE)" "RED"
        }
    } ELSE {
        AFFICHER-MESSAGE "OPERATION ANNULEE." "YELLOW"
    }
    PAUSE
}

function REDIMENSIONNER-VHDX {
    Clear-Host
    $vhds = GET-VHD
    if ($vhds.Count -eq 0) {
        AFFICHER-MESSAGE "AUCUN DISQUE VIRTUEL TROUVE a redimensionner." "RED"
        PAUSE
        RETURN
    }
    AFFICHER-MESSAGE "DISQUES virtuels disponibles a redimensionner :" "CYAN"
    $donnees_vhdx = @()
    for ($i = 0; $i -lt $vhds.Count; $i++) {
        $donnees_vhdx += [PSCustomObject]@{
            "NUMERO" = $i+1
            "CHEMIN" = $vhds[$i].PATH
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "CHEMIN") $donnees_vhdx

    $CHOIX_VHD = SAISIR-NOMBRE "NUMERO du disque a redimensionner" 1 $vhds.Count
    $PATH = $vhds[$CHOIX_VHD-1].Path
    $TAILLE_MAX = SAISIR-NOMBRE "NOUVELLE TAILLE du disque en GO" 1 2048
    TRY {
        RESIZE-VHD -PATH $PATH -SIZEBYTES ([long]($TAILLE_MAX * 1GB)) -ERRORACTION STOP
        AFFICHER-MESSAGE "DISQUE VHDX REDIMENSIONNE a $TAILLE_MAX GO." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS du redimensionnement du VHDX : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function ATTACHER-VHDX {
    Clear-Host
    $VMS = GET-VM
    $VHDX_FILES = GET-CHILDITEM -PATH "C:\" -FILTER "*.VHDX" -RECURSE -FORCE -ERRORACTION SILENTLYCONTINUE | SELECT-OBJECT -EXPAND FULLNAME
    
    if ($VMS.COUNT -eq 0 -or $VHDX_FILES.COUNT -eq 0) {
        AFFICHER-MESSAGE "AUCUNE VM ou disque VHDX disponible." "RED"
        PAUSE
        RETURN
    }

    AFFICHER-MESSAGE "VMS disponibles :" "CYAN"
    $donnees_vm = @()
    for ($i = 0; $i -lt $VMS.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $VMS[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM a modifier" 1 $VMS.Count
    $VMNAME = $VMS[$CHOIX_VM-1].NAME

    AFFICHER-MESSAGE "DISQUES VHDX disponibles :" "CYAN"
    $donnees_vhdx = @()
    for ($i = 0; $i -lt $VHDX_FILES.Count; $i++) {
        $donnees_vhdx += [PSCustomObject]@{
            "NUMERO" = $i+1
            "CHEMIN" = $VHDX_FILES[$i]
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "CHEMIN") $donnees_vhdx

    $CHOIX_VHD = SAISIR-NOMBRE "NUMERO du disque a attacher" 1 $VHDX_FILES.Count
    $VHD_PATH = $VHDX_FILES[$CHOIX_VHD-1]

    TRY {
        ADD-VMHARDDISKDRIVE -VMNAME $VMNAME -PATH $VHD_PATH -ERRORACTION STOP
        AFFICHER-MESSAGE "DISQUE '$VHD_PATH' ATTACHE a la VM '$VMNAME'." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS de l'attache du disque : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function DETACHER-VHDX {
    Clear-Host
    $VMS = GET-VM
    if ($VMS.COUNT -eq 0) {
        AFFICHER-MESSAGE "AUCUNE VM TROUVEE." "RED"
        PAUSE
        RETURN
    }

    AFFICHER-MESSAGE "VMS disponibles :" "CYAN"
    $donnees_vm = @()
    for ($i = 0; $i -lt $VMS.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $VMS[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM a modifier" 1 $VMS.Count
    $VMNAME = $VMS[$CHOIX_VM-1].NAME

    $VHD_DRIVES = GET-VMHARDDISKDRIVE -VMNAME $VMNAME
    if ($VHD_DRIVES.COUNT -eq 0) {
        AFFICHER-MESSAGE "AUCUN DISQUE VHDX ATTACHE a cette VM." "YELLOW"
        PAUSE
        RETURN
    }

    AFFICHER-MESSAGE "DISQUES attaches :" "CYAN"
    $donnees_vhd = @()
    for ($i = 0; $i -lt $VHD_DRIVES.Count; $i++) {
        $donnees_vhd += [PSCustomObject]@{
            "NUMERO" = $i+1
            "CHEMIN" = $VHD_DRIVES[$i].PATH
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "CHEMIN") $donnees_vhd

    $CHOIX_VHD = SAISIR-NOMBRE "NUMERO du disque a detacher" 1 $VHD_DRIVES.Count
    $VHD_TO_REMOVE = $VHD_DRIVES[$CHOIX_VHD-1].PATH

    TRY {
        REMOVE-VMHARDDISKDRIVE -VMNAME $VMNAME -PATH $VHD_TO_REMOVE -ERRORACTION STOP
        AFFICHER-MESSAGE "DISQUE '$VHD_TO_REMOVE' DETACHE de la VM '$VMNAME'." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS du detachement du disque : $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

# --- FONCTIONS SPECIALES ---
function GESTION-PASSTHROUGH {
    DO {
        Clear-Host
        AFFICHER-MESSAGE "GESTION des disques Passthrough" "DARKMAGENTA"
        AFFICHER-MESSAGE "1. LISTER les disques physiques" "WHITE"
        AFFICHER-MESSAGE "2. ATTACHER un disque physique a une VM" "WHITE"
        AFFICHER-MESSAGE "B. RETOUR menu gestion stockage" "WHITE"
        $CHOIX_PASS = READ-HOST "VOTRE CHOIX (1-2, B)"

        SWITCH ($CHOIX_PASS.TOUPPER()) {
            1 {
                Clear-Host
                AFFICHER-MESSAGE "LISTE des disques physiques" "CYAN"
                $disques_phys = GET-DISK | SELECT-OBJECT NUMBER, FRIENDLYNAME, PATH, OPERATIONALSTATUS
                $donnees_disques = @()
                foreach ($disk in $disques_phys) {
                    $donnees_disques += [PSCustomObject]@{
                        NUMERO = $disk.NUMBER
                        NOM = $disk.FRIENDLYNAME
                        CHEMIN = $disk.PATH
                        ETAT = $disk.OPERATIONALSTATUS
                    }
                }
                $entetes_disques = "NUMERO", "NOM", "CHEMIN", "ETAT"
                AFFICHER-TABLEAU $entetes_disques $donnees_disques
                PAUSE
            }
            2 {
                Clear-Host
                $VMS = GET-VM
                $disks = GET-DISK
                
                if ($VMS.COUNT -eq 0 -or $disks.COUNT -eq 0) {
                    AFFICHER-MESSAGE "AUCUNE VM ou disque physique disponible." "RED"
                    PAUSE
                    BREAK
                }

                AFFICHER-MESSAGE "VMS disponibles :" "CYAN"
                $donnees_vm = @()
                for ($i = 0; $i -lt $VMS.Count; $i++) {
                    $donnees_vm += [PSCustomObject]@{
                        "NUMERO" = $i+1
                        "NOM" = $VMS[$i].NAME
                    }
                }
                AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm
                
                $CHOIX_VM = SAISIR-NOMBRE "NUMERO de la VM" 1 $VMS.Count
                $VMNAME = $VMS[$CHOIX_VM-1].NAME
                
                AFFICHER-MESSAGE "DISQUES physiques disponibles par numero :" "CYAN"
                $donnees_disques = @()
                for ($i = 0; $i -lt $disks.Count; $i++) {
                    $donnees_disques += [PSCustomObject]@{
                        "NUMERO" = $disks[$i].NUMBER
                        "NOM" = $disks[$i].FriendlyName
                    }
                }
                AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_disques

                $CHOIX_DISK = SAISIR-NOMBRE "NUMERO du disque physique a attacher" 0 ($disks.Count-1)
                
                TRY {
                    ADD-VMSCSICONTROLLER -VMNAME $VMNAME | OUT-NULL
                    ADD-VMHARDDISKDRIVE -VMNAME $VMNAME -PASSTHROUGHDISKNUMBER $CHOIX_DISK -ERRORACTION STOP
                    AFFICHER-MESSAGE "DISQUE physique numero $CHOIX_DISK ATTACHE a la VM '$VMNAME'." "GREEN"
                } CATCH {
                    AFFICHER-MESSAGE "ERREUR LORS de l'attache du disque Passthrough : $($_.EXCEPTION.MESSAGE)" "RED"
                }
                PAUSE
            }
            "B" { RETURN }
            DEFAULT { AFFICHER-MESSAGE "CHOIX INVALIDE" "RED"; PAUSE }
        }
    } WHILE ($TRUE)
}

function GERER-RESEAU-VM {
    Clear-Host
    AFFICHER-MESSAGE "VMS DISPONIBLES A MODIFIER :" "CYAN"
    $vms = GET-VM | SELECT-OBJECT NAME
    
    IF ($vms.COUNT -EQ 0) {
        AFFICHER-MESSAGE "AUCUNE VM TROUVEE." "RED"
        PAUSE
        RETURN
    }
    
    $donnees_vm = @()
    for ($i = 0; $i -lt $vms.Count; $i++) {
        $donnees_vm += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vms[$i].NAME
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM") $donnees_vm

    $CHOIX_VM = SAISIR-NOMBRE "NUMERO DE LA VM A MODIFIER" 1 $vms.COUNT
    $VMName = $vms[$CHOIX_VM-1].NAME
    
    AFFICHER-MESSAGE "CARTES RESEAUX DE '$VMName' :" "YELLOW"
    
    $vmNetAdapters = GET-VMNetworkAdapter -VMName $VMName
    
    IF ($vmNetAdapters.COUNT -EQ 0) {
        AFFICHER-MESSAGE "CETTE VM N'A AUCUNE CARTE RESEAU." "RED"
        PAUSE
        RETURN
    }

    $donnees_cartes = @()
    for ($i = 0; $i -lt $vmNetAdapters.Count; $i++) {
        $donnees_cartes += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vmNetAdapters[$i].NAME
            "SWITCH" = $vmNetAdapters[$i].SWITCHNAME
            "MAC" = $vmNetAdapters[$i].MACADDRESS
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM", "SWITCH", "MAC") $donnees_cartes

    $CHOIX_CARTE = SAISIR-NOMBRE "NUMERO DE LA CARTE RESEAU A MODIFIER" 1 $vmNetAdapters.COUNT
    $vmNetAdapter = $vmNetAdapters[$CHOIX_CARTE-1]

    AFFICHER-MESSAGE "SWITCHS VIRTUELS DISPONIBLES :" "CYAN"
    $vmswitches = GET-VMSwitch | SELECT-OBJECT NAME, SWITCHTYPE
    
    $donnees_switch = @()
    for ($i = 0; $i -lt $vmswitches.Count; $i++) {
        $donnees_switch += [PSCustomObject]@{
            "NUMERO" = $i+1
            "NOM" = $vmswitches[$i].NAME
            "TYPE" = $vmswitches[$i].SWITCHTYPE
        }
    }
    AFFICHER-TABLEAU @("NUMERO", "NOM", "TYPE") $donnees_switch

    $CHOIX_SWITCH = SAISIR-NOMBRE "NUMERO DU NOUVEAU SWITCH" 1 $vmswitches.Count
    $NEW_SWITCH_NAME = $vmswitches[$CHOIX_SWITCH-1].NAME

    AFFICHER-MESSAGE "TENTATIVE de modification de la carte '$($vmNetAdapter.Name)' vers le switch '$NEW_SWITCH_NAME'..." "YELLOW"
    
    TRY {
        # Solution: Utiliser la commande Connect-VMNetworkAdapter, qui est plus fiable.
        $vmNetAdapter | CONNECT-VMNetworkAdapter -SwitchName $NEW_SWITCH_NAME -ERRORACTION STOP
        AFFICHER-MESSAGE "LA CARTE RESEAU '$($vmNetAdapter.Name)' DE LA VM '$VMName' A ETE CONNECTEE AU SWITCH '$NEW_SWITCH_NAME'." "GREEN"
    } CATCH {
        AFFICHER-MESSAGE "ERREUR LORS DU CHANGEMENT DE SWITCH: $($_.EXCEPTION.MESSAGE)" "RED"
    }
    PAUSE
}

function VERIFIER-PRE-REQUIS {
    Clear-Host
    AFFICHER-MESSAGE "VERIFICATION DES PRE-REQUIS..." "YELLOW"
    TRY {
        GET-VM | OUT-NULL
        $VM_SERVICE_STATUS = GET-SERVICE -NAME "VMMS" | SELECT-OBJECT STATUS
        if ($VM_SERVICE_STATUS.STATUS -eq "RUNNING") {
            AFFICHER-MESSAGE "SERVICE HYPER-V EN COURS D'EXECUTION. TOUS LES MODULES POWERSHELL SONT DISPONIBLES." "GREEN"
        } else {
            AFFICHER-MESSAGE "SERVICE HYPER-V NON DEMARRE. TENTATIVE DE DEMARRAGE..." "RED"
            START-SERVICE -NAME "VMMS" -ERRORACTION STOP
            AFFICHER-MESSAGE "SERVICE HYPER-V DEMARRE AVEC SUCCES." "GREEN"
        }
    } CATCH {
        AFFICHER-MESSAGE "ERREUR FATALE. LES MODULES POWERSHELL HYPER-V NE SONT PAS INSTALLES OU LE SERVICE N'A PAS PU ETRE DEMARRE." "RED"
        AFFICHER-MESSAGE "$($_.EXCEPTION.MESSAGE)" "RED"
        PAUSE
        EXIT
    }
    PAUSE
}

# --- MENUS ---
function MENU-PRINCIPAL {
    DO {
        Clear-Host
        AFFICHER-MESSAGE "MENU PRINCIPAL" "BLUE"
        AFFICHER-MESSAGE "1. GESTION DES MACHINES VIRTUELLES" "WHITE"
        AFFICHER-MESSAGE "2. GESTION DES COMMUTATEURS VIRTUELS" "WHITE"
        AFFICHER-MESSAGE "3. GESTION DES CHECKPOINTS" "WHITE"
        AFFICHER-MESSAGE "4. GESTION DU STOCKAGE (en cours, a ne pas utiiliser)" "WHITE" 
        AFFICHER-MESSAGE "R. RAFRAICHIR / RECENTER LE MENU" "WHITE"
        AFFICHER-MESSAGE "Q. QUITTER" "WHITE"
        $CHOIX_PRINCIPAL = READ-HOST "VOTRE CHOIX (1-4, R ou Q)"
        SWITCH ($CHOIX_PRINCIPAL.TOUPPER()) {
            1 { GESTION-VM }
            2 { GESTION-COMMUTATEUR }
            3 { GESTION-CHECKPOINT }
            4 { GERER-STOCKAGE }
            "R" { CONTINUE }
            "Q" { EXIT }
            DEFAULT { AFFICHER-MESSAGE "CHOIX INVALIDE" "RED"; PAUSE }
        }
    } WHILE ($TRUE)
}

function GESTION-VM {
    DO {
        Clear-Host
        AFFICHER-MESSAGE "GESTION DES MACHINES VIRTUELLES" "DARKYELLOW"
        AFFICHER-MESSAGE "1. LISTER LES VMS" "GRAY"
        AFFICHER-MESSAGE "2. DEMARRER UNE VM" "GRAY"
        AFFICHER-MESSAGE "3. ARRETER UNE VM" "GRAY"
        AFFICHER-MESSAGE "4. REDEMARRER UNE VM" "GRAY"
        AFFICHER-MESSAGE "5. CREER UNE VM" "GRAY"
        AFFICHER-MESSAGE "6. SUPPRIMER UNE VM" "GRAY"
        AFFICHER-MESSAGE "7. EXPORTER UNE VM" "GRAY"
        AFFICHER-MESSAGE "8. CONFIGURER UNE VM" "GRAY"
        AFFICHER-MESSAGE "9. GERER LES RESEAUX D'UNE VM" "GRAY"
        AFFICHER-MESSAGE "R. RAFRAICHIR / RECENTER LE MENU" "GRAY"
        AFFICHER-MESSAGE "B. RETOUR MENU PRINCIPAL" "GRAY"
        $CHOIX_VM = READ-HOST "VOTRE CHOIX (1-9, R ou B)"
        SWITCH ($CHOIX_VM.TOUPPER()) {
            1 { LISTER-VM }
            2 { DEMARRER-VM }
            3 { ARRETER-VM }
            4 { REDEMARRER-VM }
            5 { CREER-VM }
            6 { SUPPRIMER-VM }
            7 { EXPORTER-VM }
            8 { CONFIGURER-VM }
            9 { GERER-RESEAU-VM }
            "B" { RETURN }
            "R" { CONTINUE }
            DEFAULT { AFFICHER-MESSAGE "CHOIX INVALIDE" "RED"; PAUSE }
        }
    } WHILE ($TRUE)
}

function GESTION-COMMUTATEUR {
    DO {
        Clear-Host
        AFFICHER-MESSAGE "GESTION COMMUTATEUR HYPER-V" "DARKMAGENTA"
        AFFICHER-MESSAGE "1. LISTER LES COMMUTATEURS" "GRAY"
        AFFICHER-MESSAGE "2. CREER UN COMMUTATEUR" "GRAY"
        AFFICHER-MESSAGE "3. SUPPRIMER UN COMMUTATEUR" "GRAY"
        AFFICHER-MESSAGE "R. RAFRAICHIR / RECENTER LE MENU" "GRAY"
        AFFICHER-MESSAGE "B. RETOUR MENU PRINCIPAL" "GRAY"
        $CHOIXC = READ-HOST "VOTRE CHOIX (1-3, R ou B)"
        SWITCH ($CHOIXC.TOUPPER()) {
            1 { LISTER-COMMUTATEUR }
            2 { CREER-COMMUTATEUR }
            3 { SUPPRIMER-COMMUTATEUR }
            "B" { RETURN }
            "R" { CONTINUE }
            DEFAULT { AFFICHER-MESSAGE "CHOIX INVALIDE" "RED"; PAUSE }
        }
    } WHILE ($TRUE)
}

function GESTION-CHECKPOINT {
    DO {
        Clear-Host
        AFFICHER-MESSAGE "GESTION DES CHECKPOINTS" "DARKBLUE"
        AFFICHER-MESSAGE "1. LISTER LES CHECKPOINTS" "WHITE"
        AFFICHER-MESSAGE "2. CREER UN CHECKPOINT" "WHITE"
        AFFICHER-MESSAGE "3. RESTAURER UN CHECKPOINT" "WHITE"
        AFFICHER-MESSAGE "4. SUPPRIMER UN CHECKPOINT" "WHITE"
        AFFICHER-MESSAGE "R. RAFRAICHIR / RECENTER LE MENU" "WHITE"
        AFFICHER-MESSAGE "B. RETOUR MENU PRINCIPAL" "WHITE"
        $CHOIX_SNAP = READ-HOST "VOTRE CHOIX (1-4, R ou B)"
        SWITCH ($CHOIX_SNAP.TOUPPER()) {
            1 { LISTER-CHECKPOINTS }
            2 { CREER-CHECKPOINT }
            3 { RESTAURER-CHECKPOINT }
            4 { SUPPRIMER-CHECKPOINT }
            "B" { RETURN }
            "R" { CONTINUE }
            DEFAULT { AFFICHER-MESSAGE "CHOIX INVALIDE" "RED"; PAUSE }
        }
    } WHILE ($TRUE)
}

# --- DEMARRAGE DU SCRIPT ---
VERIFIER-PRE-REQUIS
MENU-PRINCIPAL
