# Générateur PDF Personnalisé pour AppSheet via Google Apps Script (v3 - Paramétré & Autonome)

Ce projet contient un script Google Apps Script **autonome (standalone)** et **générique** conçu pour être déclenché depuis une application AppSheet. Son objectif est de générer des fichiers PDF personnalisés en utilisant un modèle Google Docs et les données d'une ligne spécifique d'une feuille Google Sheets liée à AppSheet.

Il a été créé pour surmonter les limitations de la génération PDF native d'AppSheet et offrir un contrôle total sur la mise en page et l'apparence du document final. **Cette version du script est paramétrable : toute la configuration (ID du modèle, ID du fichier Sheet, dossier de destination, etc.) est passée directement depuis AppSheet lors de l'appel.**

**Dépôt privé destiné au suivi des évolutions et des demandes.**

## Problème Résolu

La génération de PDF native dans AppSheet offre peu de contrôle sur la mise en page fine. Cette solution permet d'utiliser toute la puissance de mise en page de Google Docs comme base pour les PDF, avec une configuration flexible définie directement dans AppSheet.

## Fonctionnalités Clés

*   Utilise un fichier Google Docs comme modèle de document.
*   Remplace dynamiquement des placeholders (marqueurs type `{{NomColonne}}`) dans le modèle par les données correspondantes d'une ligne Google Sheet.
*   Gère les placeholders dans le corps, l'en-tête et le pied de page du document.
*   Crée une copie du modèle pour chaque génération afin de ne pas altérer l'original.
*   Génère un fichier PDF à partir du document Google Docs rempli.
*   Sauvegarde le PDF généré dans un dossier Google Drive spécifié.
*   Nomme le fichier PDF selon un modèle configurable (ex: `BL-{{ID}}.pdf`) incluant des données de la ligne.
*   Gère automatiquement les conflits de noms de fichiers en ajoutant un suffixe numérique (ex: `NomFichier (1).pdf`).
*   **Script autonome :** Ne nécessite pas d'être lié à une feuille Sheet spécifique (créé via `script.google.com`).
*   **Entièrement configurable via 9 paramètres passés depuis AppSheet lors de l'appel**, incluant l'ID du fichier Google Sheet cible.
*   **Optionnel:** Met à jour une colonne spécifiée dans Google Sheet avec l'URL du PDF généré.
*   **Optionnel:** Supprime le fichier Google Docs intermédiaire après la génération du PDF.
*   Retourne le nom final du fichier PDF généré à AppSheet (peut être utilisé dans des étapes suivantes du Bot).

## Technologies Utilisées

*   **AppSheet:** Plateforme pour déclencher le script et **fournir les paramètres de configuration**.
*   **Google Apps Script (GAS):** Langage de script (utilisé ici en mode **autonome**).
*   **Google Docs:** Pour la création des modèles.
*   **Google Sheets:** Comme source de données.
*   **Google Drive:** Pour le stockage.

## Configuration Requise

### 1. Google Docs (Modèle)

1.  Créez votre modèle Google Doc.
2.  Insérez les placeholders `{{NomDeLaColonne}}` (correspondant aux en-têtes Google Sheet).
3.  **Notez l'ID du document modèle** (dans l'URL). **Requis pour AppSheet.**

### 2. Google Drive

1.  Créez un dossier de destination pour les PDF.
2.  **Notez l'ID de ce dossier** (dans l'URL). **Requis pour AppSheet.**

### 3. Google Sheets

1.  Identifiez le fichier Google Sheet contenant vos données AppSheet. **Notez l'ID de ce fichier Google Sheet** (dans l'URL). **Requis pour AppSheet.**
2.  Assurez-vous que la feuille (onglet) contient les colonnes nécessaires, y compris une **colonne clé unique**. Notez le **nom exact de la feuille (onglet)** et le **nom exact de l'en-tête de la colonne clé unique**. **Requis pour AppSheet.**
3.  **Optionnel:** Si vous voulez stocker l'URL du PDF, ajoutez une colonne (ex: `LienPDF`). Notez son **nom d'en-tête exact**. **Requis pour AppSheet si utilisé.**
4.  **Optionnel:** Si vous voulez stocker le nom du fichier PDF retourné, ajoutez une colonne texte (ex: `NomFichierPDF`).

### 4. Google Apps Script (Script Autonome)

1.  Allez sur `script.google.com` et créez un nouveau projet.
2.  Copiez le contenu du fichier de script `.js` de ce dépôt (la version autonome à 9 paramètres) et collez-le dans l'éditeur.
3.  **IMPORTANT :** Ce script est **autonome et générique**. Il n'y a **pas** d'objet `CONFIG` à modifier. Il utilise `SpreadsheetApp.openById()` car il ne sait pas nativement à quel fichier il est associé. La configuration se fait dans AppSheet.
4.  Sauvegardez le projet de script. Donnez-lui un nom clair (ex: "Générateur PDF Autonome Paramétrable").
5.  **(Recommandé)** Activez le moteur d'exécution V8 : `Paramètres du projet` (icône engrenage) > Cochez `Activer le moteur d'exécution Chrome V8`.
6.  **Autorisations :** La première fois que AppSheet appellera le script, Google demandera votre autorisation pour que le script accède à Drive, Docs et Sheets (il aura besoin d'accéder au fichier Sheet spécifié par son ID, au modèle Doc, et au dossier Drive). Accordez ces autorisations.

### 5. AppSheet (Configuration Cruciale de l'Appel - 9 Arguments !)

C'est ici que toute la configuration du script est définie. Soyez **extrêmement méticuleux** avec l'ordre et les guillemets.

1.  **Activer les Scripts :** `Security` > `Options` > Activez "Allow apps script execution...".
2.  **Créer une Automatisation (Bot) :**
    *   `Automation` > `Bots` > Créez un nouveau Bot.
    *   **Configurez l'Événement (Event) :** Définissez le déclencheur (Data Change, Action, Schedule...). Sélectionnez la Table et ajoutez une Condition si besoin.
    *   **Configurez le Processus (Process) :**
        *   Ajoutez une étape (`Add step`). Nommez-la clairement (ex: "Appel Script PDF Autonome").
        *   Choisissez le type d'étape : `Call a script`.
        *   **Sélectionnez votre Projet Script:** Choisissez le script autonome que vous venez de sauvegarder (ex: "Générateur PDF Autonome Paramétrable").
        *   **Nom de la Fonction :** Entrez **exactement** `generatePdfFromTemplate`.
        *   **Arguments de la Fonction :** Vous devez ajouter **9 arguments**, dans cet ordre précis.
            *   Cliquez sur `Add` 9 fois.
            *   Remplissez chaque champ comme suit :

            ---

            1.  **Argument 1 : `uniqueId`**
                *   **Description :** L'ID unique de la ligne AppSheet à traiter.
                *   **Valeur AppSheet :** `[VotreColonneCleUnique]` (Adaptez le nom de colonne). **Sans guillemets.**

            2.  **Argument 2 : `spreadsheetId` (NOUVEAU ET CRUCIAL)**
                *   **Description :** L'ID du fichier Google Sheet contenant les données.
                *   **Valeur AppSheet :** `"ID_DE_VOTRE_FICHIER_GOOGLE_SHEET"` (Collez l'ID du fichier noté à l'étape 3.1). **Avec guillemets doubles `"`**.
                *   *Exemple :* `"1zYxWvUtSrQpOnMlKjIhGfEdCbA098765"`

            3.  **Argument 3 : `templateDocId`**
                *   **Description :** L'ID de votre modèle Google Docs.
                *   **Valeur AppSheet :** `"ID_DE_VOTRE_MODELE_DOC"` (Collez l'ID du modèle). **Avec guillemets doubles `"`**.

            4.  **Argument 4 : `destinationFolderId`**
                *   **Description :** L'ID de votre dossier de destination Google Drive.
                *   **Valeur AppSheet :** `"ID_DE_VOTRE_DOSSIER_DRIVE"` (Collez l'ID du dossier). **Avec guillemets doubles `"`**.

            5.  **Argument 5 : `sheetName`**
                *   **Description :** Le nom exact de la feuille (onglet) dans le Google Sheet spécifié.
                *   **Valeur AppSheet :** `"NomDeVotreFeuille"` (Nom exact de l'onglet). **Avec guillemets doubles `"`**.

            6.  **Argument 6 : `uniqueIdColumnNameInSheet`**
                *   **Description :** Le nom exact de l'en-tête de la colonne clé unique dans Google Sheet.
                *   **Valeur AppSheet :** `"NomColonneCleDansSheet"` (En-tête exact). **Avec guillemets doubles `"`**.

            7.  **Argument 7 : `pdfLinkColumnName`**
                *   **Description :** Le nom exact de l'en-tête (Google Sheet) de la colonne pour l'URL du PDF. Vide si non utilisé.
                *   **Valeur AppSheet :** `"NomColonneLienPDFDansSheet"` ou `""`. **Avec guillemets doubles `"`**.

            8.  **Argument 8 : `pdfFilenameTemplate`**
                *   **Description :** Le modèle pour nommer le fichier PDF (avec `{{NomColonneSheet}}`).
                *   **Valeur AppSheet :** `"BL-{{ID Commande}}-{{NomClient}}.pdf"` (Adaptez !). **Avec guillemets doubles `"`**.

            9.  **Argument 9 : `deleteTempDocStr`**
                *   **Description :** Supprimer le Doc temporaire (`"true"`) ou le garder (`"false"`).
                *   **Valeur AppSheet :** `"false"` ou `"true"`. **Avec guillemets doubles `"`**.

            ---

            **Récapitulatif Visuel des 9 Arguments dans AppSheet :**

            ```
            1: [VotreColonneCleUnique]
            2: "ID_FICHIER_GOOGLE_SHEET"    <-- Le nouveau !
            3: "ID_MODELE_DOC"
            4: "ID_DOSSIER_DRIVE"
            5: "NOM_FEUILLE_SHEET"
            6: "NOM_COLONNE_CLE_SHEET"
            7: "NOM_COLONNE_LIEN_PDF_SHEET" (ou "")
            8: "TEMPLATE_NOM_FICHIER_{{Col1}}.pdf"
            9: "false" (ou "true")
            ```

            **ATTENTION (bis) :** L'ordre, les guillemets pour les textes fixes, et l'absence de guillemets pour l'expression `[Colonne]` sont **ABSOLUMENT CRITIQUES**.

        *   **(Optionnel) Capturer la sortie :** Cochez `Wait for return value`. Utilisez `[NomDeVotreEtapeScript].[Output]` pour récupérer le nom du fichier.
    *   **(Optionnel) Utiliser la sortie :** Ajoutez une étape `Data: Set column values` *après* pour stocker `[NomDeVotreEtapeScript].[Output]` dans une colonne.

3.  Sauvegardez votre application AppSheet.

## Utilisation

1.  Déclenchez l'événement configuré dans AppSheet.
2.  Synchronisez votre application.
3.  Le script s'exécute. Vérifiez le dossier Drive et les colonnes Sheet/AppSheet.
4.  **En cas de problème :**
    *   Vérifiez **les 9 arguments** dans AppSheet (ordre, guillemets, valeurs, **surtout le nouvel ID du fichier Sheet en 2ème position**).
    *   Vérifiez les permissions du script autonome pour accéder au fichier Sheet spécifié, au Doc modèle et au dossier Drive.
    *   Consultez les journaux d'exécution dans l'éditeur Google Apps Script (`Exécutions`).
    *   Consultez l'historique d'audit dans AppSheet (`Manage` > `Monitor` > `Audit History`).

## Suivi des Évolutions et Contributions (Interne)

*   Utiliser les **Issues** GitHub.
*   Utiliser les **branches** Git.
*   Utiliser les **Pull Requests**.
*   Rédiger des **messages de commit** clairs.

## Améliorations Futures Potentielles (To-Do)

*   [ ] Améliorer la gestion des erreurs (notifications).
*   [ ] Option pour envoyer par email.
*   [ ] Gérer des remplacements plus complexes.
*   [ ] Créer une feuille de configuration dans Google Sheets pour stocker les ID (modèle, dossier, fichier Sheet) et le template de nom, puis lire cette feuille depuis AppSheet pour simplifier la configuration de l'appel au script (passer moins d'arguments fixes).
