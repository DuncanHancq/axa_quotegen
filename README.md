# Générateur de Devis AXA

Ce projet permet aux utilisateurs de générer des devis en associant des champs d'un fichier Excel à des espaces réservés dans un modèle DOCX. Les devis générés sont ensuite compressés et disponibles en téléchargement.

## Fonctionnalités

- Téléchargement de fichiers Excel et DOCX.
- Association des colonnes Excel aux espaces réservés sur un template DOCX.
- Suggestion automatique des associations
- Génération de fichiers DOCX basés sur les associations et conversion en pdf. (todo)
- Téléchargement des fichiers PDF générés et trier sous forme d'archive ZIP. (fichiers DOCX pour le moment)

## Prérequis

- Node.js (v14 ou supérieur)
- npm (v6 ou supérieur)

## Installation

1. Installez les dépendances :
    ```sh
    npm install
    ```

## Lancement du Projet

1. Démarrez le serveur :
    ```sh
    node app.js
    ```

2. Ouvrez votre navigateur et accédez à :
    ```
    http://localhost:3000
    ```

## Utilisation

1. Téléchargez un fichier Excel et un modèle DOCX.
2. Sélectionnez la feuille du fichier Excel.
3. Associez les colonnes Excel aux espaces réservés DOCX.
4. Cliquez sur "Créer les devis" pour générer les devis.
5. Cliquez sur "Télécharger le fichier ZIP" pour télécharger les devis générés sous forme d'archive ZIP.

## Todo

- Conversion docx -> pdf
- Package "xlsx" à remplacer (problème de sécu)
- Front end
- Nettoyage de code
