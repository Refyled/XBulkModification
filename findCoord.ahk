#Persistent
#SingleInstance Force
CoordMode, Mouse, Screen  ; Utiliser les coordonnées de l'écran complet

isRunning := false  ; État du script (démarré ou non)

; Lancer le script avec Page Up
~Up::
    isRunning := true
    SetTimer, ShowMouseCoords, 100  ; Actualiser les coordonnées toutes les 100ms
    ToolTip, Script en cours... (Appuyez sur Page Down pour arrêter)
return

; Arrêter le script avec Page Down
Down::
    isRunning := false
    SetTimer, ShowMouseCoords, Off  ; Désactiver le timer
    ToolTip  ; Effacer le message à l'écran
    ExitApp  ; Quitter proprement le script
return

; Fonction pour afficher les coordonnées de la souris
ShowMouseCoords:
    if (isRunning)
    {
        MouseGetPos, x, y  ; Obtenir les positions X et Y de la souris
        ToolTip, Coordonnées de la souris :`nX: %x%`nY: %y%
    }
return