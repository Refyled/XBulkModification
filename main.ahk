; -------------------------------------------------------
; Initialisation
; -------------------------------------------------------

; Variables globales pour arreter le script
;zoom 110
CoordMode, Mouse, Screen  ;
global stopScript := false
Global enableLogging := false ; Permet de toggler le logging
Global successCount := 0
Global errorCount := 0



; -------------------------------------------------------
; Variables à modifier 
; -------------------------------------------------------

;la liste csv des modifications à effectuer
Global csvFile := "C:\Users\henri\Desktop\AHK\list.csv"
;le repertoire des images
Global imageDir := "C:\Users\henri\Desktop\AHK\Img\"
;le fichier log (si log est activé il sera créé automatiquement, ne sera pas utilisé sinon)
Global logFile := "C:\Users\henri\Desktop\AHK\log.csv"
;lien vers une page d'accueil neutre
Global neutralLink := "https://app.ximi.xelya.io/N106/List/Client?Mode=List&View=Default"
;lien vers une page d'édition'
Global editLink := "https://app.ximi.xelya.io/N106/Details/Intervention/"

; -------------------------------------------------------
; Logging
; -------------------------------------------------------

; Fonction pour ajouter une entree au log
Log(action, message) {
    if !enableLogging
        return
    File := FileOpen(logFile, "a")
    if !File {
        MsgBox, 16, Erreur, Impossible d'ouvrir le fichier de log.
        ExitApp
    }
    File.WriteLine(action ";" A_Now ";" message)
    File.Close()
}

; -------------------------------------------------------
; Manipulation csv
; -------------------------------------------------------

; Verifier si le fichier existe
if !FileExist(csvFile) {
    if enableLogging
        FileAppend, % "Erreur;" A_Now ";Le fichier CSV n'existe pas.`n", logFile
    ExitApp
}

; Fonction pour lire une ligne specifique y et retourner un tableau contenant les valeur de la colonne (y,1);(y,2);(y;3)...
Lire(csvFile, ligne) {
    Log("Lecture", "Lecture de la ligne " ligne)
    File := FileOpen(csvFile, "r")
    if !File {
        Log("Erreur", "Impossible d'ouvrir le fichier CSV")
        ExitApp
    }
    currentLine := 0
    while !File.AtEOF {
        line := File.ReadLine()
        currentLine++
        if (currentLine = ligne) {
            cols := StrSplit(line, ";")
            if (4 > cols.Length()) {
                File.Close()
                Log("Erreur", "Colonnes insuffisantes dans la ligne " ligne)
                ExitApp
            }
            File.Close()
            return cols
        }
    }
    File.Close()
    Log("Erreur", "La ligne " ligne " n'existe pas dans le fichier CSV")
    ExitApp
}

; Fonction pour compter les lignes dans le csv

CountLinesInCSV(csvFile) {
    lineCount := 0
    if !FileExist(csvFile) {
        Log("Erreur", "Le fichier n'existe pas")
        return -1 ; Retourne -1 en cas d'erreur
    }
    File := FileOpen(csvFile, "r")
    if !File {
        Log("Erreur", "Impossible d'ouvrir le fichier")
        return -1 ; Retourne -1 en cas d'erreur
    }
    while !File.AtEOF {
        File.ReadLine() ; Lit chaque ligne
        lineCount++ ; Incremente le compteur
    }
    File.Close() ; Ferme le fichier
    return lineCount
}

; Fonction pour ecrire une valeur dans une cellule specifique (ligne, colonne)
Ecrire(csvFile, ligne, colonne, valeur) {
    Log("Ecriture", "Ecriture dans la ligne " ligne ", colonne " colonne)
    File := FileOpen(csvFile, "r")
    if !File {
        Log("Erreur", "Impossible d'ouvrir le fichier CSV pour ecriture")
        ExitApp
    }
    newCsvContent := ""
    currentLine := 0
    while !File.AtEOF {
        line := RTrim(File.ReadLine(), "`r`n") ; Retirer les sauts de ligne superflus
        currentLine++
        if (currentLine = ligne) {
            cols := StrSplit(line, ";")
            if (colonne > cols.Length()) {
                File.Close()
                Log("Erreur", "La colonne " colonne " n'existe pas dans la ligne " ligne)
                ExitApp
            }
            cols[colonne] := valeur
            newLine := ""
            for index, value in cols {
                newLine .= value . (index < cols.Length() ? ";" : "")
            }
            newCsvContent .= newLine "`n"
        } else {
            newCsvContent .= line "`n"
        }
    }
    File.Close()
    newCsvContent := RTrim(newCsvContent, "`n") ; Retirer les sauts de ligne finaux superflus
    File := FileOpen(csvFile, "w")
    if !File {
        Log("Erreur", "Impossible d'ecrire dans le fichier CSV")
        ExitApp
    }
    File.Write(newCsvContent)
    File.Close()
}

; -------------------------------------------------------
; Fonctions utiles pour la recherche d'images/de coordonnées
; -------------------------------------------------------

; Fonction : Attendre qu'une image apparaisse avec un temps d'attente maximum : cherche une image sur l'écran, renvoie ses coordonnées si il trouve sous la forme {x: foundX, y: foundY}
FindImageTopLeft(imageName, timeout) {
    ; Construit le chemin complet de l'image
    imagePath := imageDir imageName

    ; Verifie si le chemin de l'image est valide
    if (!FileExist(imagePath)) {
        Log("Erreur", "L'image specifiee n'existe pas : " imagePath)
        return False
    }
    
    startTime := A_TickCount  ; Enregistre le temps de debut
    
    ; Attente de l'image tant qu'elle n'est pas trouvee et que le timeout n'est pas atteint
    Loop {
        ; Verifie si le script a ete arrete
        if (stopScript) {
            return False
        }
        
        ; Recherche de l'image sur tout l'ecran
        ImageSearch, foundX, foundY, 0, 0, A_ScreenWidth, A_ScreenHeight, %imagePath%
        
        ; Si l'image est trouvee, renvoie les coordonnees du coin superieur gauche
        if (!ErrorLevel) {
            return {x: foundX, y: foundY}
        }
        
        ; Verifie si le temps d'attente maximum est depasse
        if (A_TickCount - startTime > timeout) {
            return False
        }
        
        Sleep, 100 ; Pause de 100 ms avant de reessayer
    }
}

verif(coords) {
    ; Verifie si l'image a ete trouvee
    if (coords) {
        return, True
    } else if (stopScript) {
        Log("Info", "Recherche arretee par l'utilisateur")
        return, False
    } else {
        return, False
    }
}

; -------------------------------------------------------
; Fonctions utiles pour la manipulation du navigateur
; -------------------------------------------------------

; Fonction : ouvrir un lien

goToLink(link) {
    Log("Navigation", "Acces au lien : " link)
    Click 1365, 64
    Sleep, 50
    Clipboard := link
    Sleep, 50
    Send, ^v
    Sleep, 50
    Send, {Enter}
    coords := FindImageTopLeft("retour.png", 3000)
    if (verif(coords)){
        return
    }
    Click 1365, 64
    Sleep, 400
    Clipboard := link
    Sleep, 100
    Send, ^v
    Sleep, 400
    Send, {Enter}
    coords := FindImageTopLeft("retour.png", 7000)
    if (!verif(coords)){
        Log("Erreur", "Lien inaccessible ou mauvais navigateur")
        ExitApp
    }
    return
}

;cette fonction sert a bien changer de page en cas d'erreur, pour etre certain de ne pas editer par erreur l'entree en erreur avec les donnees de la suivante
goToNeutralLink() {
    Log("Navigation", "Acces au lien neutre")
    Click 1365, 64
    Sleep, 100
    link:= neutralLink
    Clipboard := link
    Send, ^v
    Sleep, 100
    Send, {Enter}
    Sleep, 300
    Send, {Enter}
    coords := FindImageTopLeft("creer.png", 7000)
    if (!verif(coords)){
        Log("Erreur", "Lien neutre inaccessible")
        ExitApp
    }
    return
}

; -------------------------------------------------------
; Fonction principale décrivant la logique de traitement d'une ligne du csv
; -------------------------------------------------------

;la fonction ouvre le csv, regarde si le statut est "td"
; si le statut est td 
; la fonction cherche le champs de début à modifier en cherchant l'image "debut"
; clique à côté de l'image une fois trouvées
; selectionne le contenu du champs en cliquant dessus 
; copie en mémoire la valeur de début par laquelle remplacer
; envoie controle v pour coller la nouvelle valeur
; fait de meme pour le champs d'heure de fin 
; cherche et clique sur l'image "enregistrer et retour"
; cherche et clique sur le bouton confirmer si, présent : si c'est le cas elle vérifie que l'on est revenu sur l'accueil et enregistre la ligne comme "done" dans le csv
; sinon elle cherche un bandeau d'erreur et enregistre la lgine comme en erreur dans le csv 



treatLine(line,csvFile) {
    Log("Traitement", "Traitement de la ligne " line)
    result := Lire(csvFile, line)
    statut := result[4]

    if (statut=="td")
    {
        idNum:= result[1]
        link:= editLink idNum "?Edit"
        do:=goToLink(link)

        coords := FindImageTopLeft("debut.png", 3000)
        if (!verif(coords)){
            Log("Erreur", "Champ 'debut' non trouve dans la ligne " line)
            errorCount++
            ExitApp
        }

        champsx:= coords.x + 19
        champsy:= coords.y + 19
        MouseMove, champsx, champsy
        Sleep, 30
        Click 3

        debut:= result[2]
        Clipboard := debut
        Sleep, 30
        Send, ^v
        Sleep, 30

        coords := FindImageTopLeft("fin.png", 3000)
        if (!verif(coords)){
            Log("Erreur", "Champ 'fin' non trouve dans la ligne " line)
            errorCount++
            ExitApp
        }

        champsx:= coords.x + 19
        champsy:= coords.y + 19
        MouseMove, champsx, champsy
        Sleep, 40
        Click 3

        fin:= result[3]
        Clipboard := fin
        Sleep, 30
        Send, ^v
        Sleep, 30

        coords := FindImageTopLeft("retour.png", timeout)
        if (!verif(coords)){ 
            Log("Erreur", "Bouton 'enregistrer et retour' non trouve")
            errorCount++
            ExitApp
        }
        MouseMove, coords.x, coords.y
        Sleep, 50
        Click

        confirmationLoop:

        coords := FindImageTopLeft("confirmer.png", 100)
        if (!((verif(coords))=False)){ 
            MouseMove, coords.x, coords.y, 0
            Sleep, 10
            Click
        }

        coords := FindImageTopLeft("erreur.png", 100)
        if (verif(coords)){ 
            Ecrire(csvFile, line, 4, "Erreur")
            Log("Erreur", "Erreur detectee dans la ligne " line)
            errorCount++
            goToNeutralLink()
            return
        }

        coords := FindImageTopLeft("creer.png", 100)
        if (!verif(coords)){
            Send, {Enter} 
            goto confirmationLoop
        } else {
            Ecrire(csvFile, line, 4, "ok")
            Log("Succes", "Ligne " line " traitee avec succes")
            successCount++
            return
        }

    } else {
        Log("Info", "Ligne " line " ignoree : statut different de 'td'")
        return
    }
    return
}

; -------------------------------------------------------
; controle
; -------------------------------------------------------

; en appuyant sur haut, on parcours le csv file et on traite ligne à ligne

~Up::
{
    MsgBox, BEGIN
    nbOfLines:=CountLinesInCSV(csvFile) - 1
    successCount := 0
    errorCount := 0
    
    k :=2
    loop %nbOfLines% 
    {
        treatLine(k,csvFile)
        k:=k+1
    }
    MsgBox, Traitement termine ! Lignes traitees : %nbOfLines%`nSucces : %successCount%`nErreurs : %errorCount%
    Log("Info", "Traitement termine. Succes: " successCount ", Erreurs: " errorCount)
    return
}

; en appuyant sur bas, on arrête le script

Down::
    MsgBox, FORCEEND
    stopScript := true  ; Change la variable d'arret
    ExitApp
return

#Persistent  ; Assure que le script reste actif
