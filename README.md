# XBulkModification

# 🛠 Script AHK - Modification d'interventions dans Ximi via CSV

Ce script AutoHotkey permet de **modifier automatiquement les heures de début et de fin** d’interventions dans l’interface Ximi Mandataire à partir d’un **fichier CSV**.

Il interagit avec le navigateur via la reconnaissance d’images pour localiser les champs à modifier, et met à jour les entrées ligne par ligne.

---

## 📌 Fonctionnalités

- Lecture d’un fichier CSV ligne par ligne.
- Pour chaque ligne marquée `td` (à traiter) :
  - Accès à la fiche intervention via l’URL de Ximi.
  - Remplacement des champs "heure de début" et "heure de fin".
  - Sauvegarde de l’intervention.
  - Vérification du succès ou de l’échec.
  - Mise à jour de la ligne dans le CSV (`ok` ou `Erreur`).

---

## 📂 Structure du fichier CSV

Le fichier doit être **au format texte avec des séparateurs de colonnes en point-virgule `;`**.

Exemple :

```
id;heure_debut;heure_fin;statut
12345;08:00;09:00;td
67890;10:30;12:00;ok
```

- `id` : identifiant de l’intervention dans Ximi.
- `heure_debut` : nouvelle heure de début.
- `heure_fin` : nouvelle heure de fin.
- `statut` : doit être `td` pour que la ligne soit traitée.

---

## 🧰 Prérequis

### 1. Installer AutoHotkey

- Téléchargez et installez AHK depuis [https://www.autohotkey.com/](https://www.autohotkey.com/)

### 2. Télécharger et préparer le dossier

Votre dossier doit contenir :

```
📁 AHK/
├── main.ahk           # Le script principal
├── list.csv             # Le fichier CSV à traiter
├── log.csv              # (optionnel) journal des actions
└── Img/                 # Dossier d'images pour la reconnaissance
    ├── debut.png
    ├── fin.png
    ├── retour.png
    ├── confirmer.png
    ├── erreur.png
    └── creer.png
```

---

## ⚙️ Configuration du script

En haut du script, modifiez les variables suivantes avec vos chemins :

```ahk
Global csvFile := "C:\Chemin\Vers\list.csv"
Global imageDir := "C:\Chemin\Vers\Img\"
Global logFile := "C:\Chemin\Vers\log.csv"
Global neutralLink := "https://..."
Global editLink := "https://..."
```

---

## 📸 Captures d’écran à préparer

Les captures d’écran utilisées pour la reconnaissance d’image doivent être précises. Utilisez `Windows + Maj + S` pour capturer :

- Les libellés ou icônes associés aux champs (`debut.png`, `fin.png`)
- Les boutons (`retour.png`, `confirmer.png`, `creer.png`)
- Les messages d’erreur (`erreur.png`)

Attention a bien capturer le même format exact que les images originales

---

## 🏁 Utilisation

### Lancer le script :
1. Ouvrez un navigateur (Chrome, Firefox...) connecté à Ximi.
2. Double-cliquez sur `main.ahk` pour l'exécuter.
3. Appuyez sur la **flèche haut (↑)** pour démarrer le traitement ligne par ligne.

### Interrompre le script :
- Appuyez sur la **flèche bas (↓)** pour arrêter proprement.

---

## 🔍 Structure du code

Le script est structuré en plusieurs blocs :

- **Initialisation** : chemins, options globales, raccourcis.
- **Logging** : écrit un journal (`log.csv`) si activé.
- **Manipulation CSV** : fonctions pour lire, écrire et compter les lignes.
- **Reconnaissance d’image** : recherche d’éléments visuels à l’écran.
- **Navigation** : ouverture de liens dans le navigateur.
- **Traitement ligne** : logique de modification de l’intervention.
- **Contrôle clavier** : raccourcis pour lancer ou stopper.

---

## ℹ️ Remarques importantes

- Le script repose entièrement sur l’interface visuelle de Ximi : toute modification d’UI peut nécessiter de refaire les captures d’écran.
- Le temps de chargement de Ximi varie, les pauses (`Sleep`) sont prévues mais peuvent être ajustées.

---

## 📝 À venir 

- Interface graphique minimale pour indiquer la progression.
- Script de mise à jour des images.
