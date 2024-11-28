# Traitement Excel avec Calculs de Temps pour Prévision des Retards de Chargement

## Description

Ce programme est une application de bureau développée avec Python, permettant aux chefs d'équipe logistique de prévoir les retards de chargement en fonction des données d'un fichier Excel. Il calcule les heures nécessaires pour terminer les chargements et génère des rapports détaillant les temps de fin estimés, avec une mise en forme conditionnelle pour alerter visuellement sur les éventuels retards.

L'application permet de :
- Importer un fichier Excel contenant des informations sur les chargements.
- Saisir la production par heure (UVC/h) et la date/heure de début.
- Calculer le temps estimé nécessaire pour terminer les chargements.
- Ajouter un préfixe personnalisé au fichier de sortie.
- Visualiser un aperçu du fichier traité avec des couleurs indicatives (vert pour conforme, rouge pour retard, jaune pour délai proche).

## Fonctionnalités

- **Chargement d'un fichier Excel** : Sélectionnez un fichier Excel contenant des données de chargement.
- **Calcul de la durée des chargements** : En fonction des données d'entrée, le programme calcule les heures nécessaires pour terminer les chargements et génère un horaire de fin.
- **Mise en forme conditionnelle** : La couleur des cellules du fichier Excel est modifiée en fonction des résultats :
  - **Bleu** pour les sous-totaux.
  - **Rouge** pour les retards de chargement.
  - **Jaune** pour les chargements proches de leur délai.
- **Aperçu des données** : Affichage d'une vue préliminaire des résultats sous forme de tableau interactif.
- **Génération d'un fichier Excel de sortie** : Le fichier traité est enregistré avec un préfixe spécifié par l'utilisateur.

## Prérequis

L'application nécessite Python et certaines bibliothèques pour fonctionner correctement. Les bibliothèques nécessaires sont :

- `pandas` : Pour manipuler les données des fichiers Excel.
- `openpyxl` : Pour traiter les fichiers Excel `.xlsx`, `.xlsm`.
- `customtkinter` : Pour l'interface graphique.
- `tkinter` : Pour l'interface utilisateur (incluse par défaut dans Python).
- `xlrd` et `pyxlsb` : Pour la lecture de fichiers `.xls` et `.xlsb`.

## Installation des dépendances

Si vous exécutez l'application en mode développement ou en environnement Python, vous devrez installer les bibliothèques suivantes :

```bash
pip install pandas openpyxl customtkinter xlrd pyxlsb
```

## Utilisation

1. **Lancer l'application** :
   - Exécutez le fichier Python avec l'IDE de votre choix, ou créez un fichier exécutable `.exe` pour une utilisation sans installation de Python.

2. **Sélectionner un fichier Excel** :
   - Cliquez sur **"Choisir un fichier Excel"** pour importer un fichier de données de chargement.
   
3. **Saisir les informations** :
   - Entrez la production par heure (UVC/h).
   - Entrez la date et l'heure de début (format `YYYY-MM-DD HH:MM:SS`).
   - Entrez un préfixe pour le nom du fichier final.

4. **Lancer le traitement** :
   - Cliquez sur **"Lancer le traitement"** pour effectuer les calculs et générer le fichier Excel de sortie.

5. **Aperçu des données** :
   - Cliquez sur **"Aperçu du fichier"** pour afficher une fenêtre avec un tableau des résultats.

6. **Fichier de sortie** :
   - Le fichier généré sera enregistré dans le répertoire actuel, avec le préfixe choisi.

## Créer un fichier exécutable (.exe)

Si vous souhaitez exécuter l'application sans avoir besoin de Python ou de bibliothèques installées, vous pouvez créer un fichier `.exe` autonome à l'aide de PyInstaller.

### Étapes pour créer un `.exe` avec PyInstaller

1. **Installer PyInstaller** :
   Si vous ne l'avez pas déjà installé, vous pouvez le faire avec cette commande :
```bash
pip install pyinstaller
```

2. **Créer le fichier `.exe`** :
Exécutez cette commande dans le terminal, en remplaçant `nom_du_script.py` par le nom de votre fichier Python :

```bash
pyinstaller --onefile --windowed nom_du_script.py
```

- `--onefile` : Crée un fichier `.exe` unique.
- `--windowed` : Empêche l'affichage de la fenêtre de la console.

3. **Vérifier la sortie** :
Après l'exécution de la commande, le fichier `.exe` sera situé dans le dossier `dist`.

4. **Exécuter depuis une clé USB** :
Copiez le fichier `.exe` sur une clé USB, puis exécutez-le directement depuis n'importe quel ordinateur.

## Mise en forme conditionnelle

Les cellules du fichier Excel sont colorées en fonction des valeurs calculées :

- **Sous-total** : Les lignes contenant le sous-total sont coloriées en bleu.
- **Retards** : Si l'heure de fin dépasse la date de chargement, la cellule est coloriée en rouge.
- **Délai proche** : Si l'écart entre l'heure de fin et la date de chargement est inférieur ou égal à 30 minutes, la cellule est coloriée en jaune.

## Exemple d'utilisation

1. **Entrées** :
- Production par heure (UVC/h) : `500`
- Date/Heure de début : `2024-11-28 08:00:00`
- Préfixe du fichier final : `Prévision_Retards`

2. **Sortie** :
- Un fichier Excel sera généré avec un préfixe `Prévision_Retards` et contenant des données avec les heures de fin calculées et les mises en forme conditionnelles appliquées.

## Dépannage

- **Erreur lors du chargement du fichier Excel** : Vérifiez que le fichier est bien dans le format supporté (`.xlsx`, `.xls`, `.xlsm`, `.xlsb`).
- **Erreur de conversion de date** : Assurez-vous que la date et l'heure de début sont saisies dans le format correct (`YYYY-MM-DD HH:MM:SS`).
- **Problèmes d'interface** : Si l'application ne s'affiche pas correctement, vérifiez que la bibliothèque `customtkinter` est correctement installée.

## Licence

Ce projet est sous licence **MIT**. Vous êtes libre de l'utiliser, de le modifier et de le distribuer selon les termes de la licence.

---

© Yanis Bordonado - Tous droits réservés

