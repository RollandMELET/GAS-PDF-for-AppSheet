# Générateur PDF Personnalisé pour AppSheet via Google Apps Script (v2 - Paramétré)

Ce projet contient un script Google Apps Script **générique** conçu pour être déclenché depuis une application AppSheet. Son objectif est de générer des fichiers PDF personnalisés en utilisant un modèle Google Docs et les données d'une ligne spécifique d'une feuille Google Sheets liée à AppSheet.

Il a été créé pour surmonter les limitations de la génération PDF native d'AppSheet et offrir un contrôle total sur la mise en page et l'apparence du document final. **Cette version du script est paramétrable : toute la configuration (ID du modèle, dossier de destination, etc.) est passée directement depuis AppSheet lors de l'appel.**

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
*   **Entièrement configurable via les paramètres passés depuis AppSheet lors de l'appel.**
*   **Optionnel:** Met à jour une colonne spécifiée dans Google Sheet avec l'URL du PDF généré.
*   **Optionnel:** Supprime le fichier Google Docs intermédiaire après la génération du PDF.
*   Retourne le nom final du fichier PDF généré à AppSheet (peut être utilisé dans des étapes suivantes du Bot).

## Technologies Utilisées

*   **AppSheet:** Plateforme pour déclencher le script et **fournir les paramètres de configuration**.
*   **Google Apps Script (GAS):** Langage de script pour l'automatisation.
*   **Google Docs:** Pour la création des modèles.
*   **Google Sheets:** Comme source de données.
*   **Google Drive:** Pour le stockage.

## Configuration Requise

### 1. Google Docs (Modèle)

1.  Créez votre modèle Google Doc avec la mise en page souhaitée.
2.  Insérez les placeholders `{{NomDeLaColonne}}` où les données doivent apparaître. Le `NomDeLaColonne` doit correspondre **exactement** à l'en-tête de colonne dans Google Sheet.
3.  **Notez l'ID du document modèle** (dans l'URL : `/d/ID_DOCUMENT/edit`). **Vous en aurez besoin dans AppSheet.**

### 2. Google Drive

1.  Créez un dossier de destination pour les PDF.
2.  **Notez l'ID de ce dossier** (dans l'URL). **Vous en aurez besoin dans AppSheet.**

### 3. Google Sheets

1.  Assurez-vous que votre feuille contient les colonnes nécessaires, y compris une **colonne clé unique**. Notez le **nom exact de la feuille** et le **nom exact de la colonne clé unique**. **Vous en aurez besoin dans AppSheet.**
2.  **Optionnel:** Si vous voulez stocker l'URL du PDF, ajoutez une colonne (ex: `LienPDF`). Notez son **nom exact**. **Vous en aurez besoin dans AppSheet.**
3.  **Optionnel:** Si vous voulez stocker le nom du fichier PDF retourné par le script, ajoutez une colonne texte (ex: `NomFichierPDF`).

### 4. Google Apps Script

1.  Ouvrez votre Google Sheet associée à AppSheet.
2.  Allez dans `Outils` > `Éditeur de scripts`.
3.  Copiez le contenu du fichier de script `.js` de ce dépôt (la version paramétrée) et collez-le dans l'éditeur.
4.  **IMPORTANT :** Ce script est **générique**. **Il n'y a plus d'objet `CONFIG` à modifier directement dans le code du script.** Toute la configuration se fait dans AppSheet.
5.  Sauvegardez le projet de script (icône disquette). Donnez-lui un nom clair (ex: "Générateur PDF Paramétrable").
6.  **(Recommandé)** Activez le moteur d'exécution V8 : `Paramètres du projet` (icône engrenage) > Cochez `Activer le moteur d'exécution Chrome V8`.
7.  **Autorisations :** La première fois que AppSheet appellera le script, Google demandera votre autorisation pour que le script accède à Drive, Docs et Sheets en votre nom. Accordez ces autorisations.

### 5. AppSheet (Configuration Cruciale de l'Appel)

C'est ici que toute la configuration du script est définie. Soyez très méticuleux.

1.  **Activer les Scripts :** `Security` > `Options` > Activez "Allow apps script execution...".
2.  **Créer une Automatisation (Bot) :**
    *   `Automation` > `Bots` > Créez un nouveau Bot.
    *   **Configurez l'Événement (Event) :** Définissez ce qui déclenche le PDF (Data Change, Action, Schedule...). Sélectionnez la bonne Table et ajoutez une Condition si besoin.
    *   **Configurez le Processus (Process) :**
        *   Ajoutez une étape (`Add step`). Nommez-la clairement (ex: "Appel Script PDF Personnalisé").
        *   Choisissez le type d'étape : `Call a script`.
        *   **Sélectionnez votre Projet Script:** Choisissez le script que vous venez de sauvegarder (ex: "Générateur PDF Paramétrable").
        *   **Nom de la Fonction :** Entrez **exactement** `generatePdfFromTemplate`.
        *   **Arguments de la Fonction :** C'est la partie la plus importante ! Vous devez ajouter **8 arguments**, dans cet ordre précis.
            *   Cliquez sur `Add` 8 fois pour créer 8 champs d'arguments.
            *   Remplissez chaque champ comme suit :

            ---

            1.  **Argument 1 : `uniqueId`**
                *   **Description :** L'identifiant unique de la ligne AppSheet à traiter.
                *   **Valeur à entrer dans AppSheet :** `[VotreColonneCleUnique]`
                    *   *Explication :* Utilisez l'expression AppSheet pour récupérer la valeur de la colonne clé de la ligne qui a déclenché le bot. Remplacez `VotreColonneCleUnique` par le nom réel de votre colonne clé dans AppSheet. **Pas de guillemets ici.**

            2.  **Argument 2 : `templateDocId`**
                *   **Description :** L'ID de votre modèle Google Docs.
                *   **Valeur à entrer dans AppSheet :** `"ID_DE_VOTRE_MODELE_GOOGLE_DOC"`
                    *   *Explication :* Collez l'ID que vous avez noté à l'étape Google Docs. **Mettez cette valeur entre guillemets doubles `"`** car c'est une chaîne de texte fixe.
                    *   *Exemple :* `"123abcDEF456xyz"`

            3.  **Argument 3 : `destinationFolderId`**
                *   **Description :** L'ID de votre dossier de destination Google Drive.
                *   **Valeur à entrer dans AppSheet :** `"ID_DE_VOTRE_DOSSIER_DE_DESTINATION"`
                    *   *Explication :* Collez l'ID du dossier noté à l'étape Google Drive. **Mettez cette valeur entre guillemets doubles `"`**.
                    *   *Exemple :* `"789ghiJKL012uvw"`

            4.  **Argument 4 : `sheetName`**
                *   **Description :** Le nom exact de la feuille Google Sheet contenant les données.
                *   **Valeur à entrer dans AppSheet :** `"NomDeVotreFeuille"`
                    *   *Explication :* Entrez le nom exact de votre feuille (onglet en bas de Google Sheets). **Mettez cette valeur entre guillemets doubles `"`**.
                    *   *Exemple :* `"Commandes"`

            5.  **Argument 5 : `uniqueIdColumnNameInSheet`**
                *   **Description :** Le nom exact de la colonne clé unique, tel qu'il apparaît en **en-tête** dans votre feuille Google Sheet.
                *   **Valeur à entrer dans AppSheet :** `"NomColonneCleUnique"`
                    *   *Explication :* Entrez le nom de l'en-tête de la colonne clé dans Google Sheets. **Mettez cette valeur entre guillemets doubles `"`**.
                    *   *Exemple :* `"ID Commande"` (S'il y a un espace, mettez-le !) ou `"ID_Commande"`

            6.  **Argument 6 : `pdfLinkColumnName`**
                *   **Description :** Le nom exact de la colonne (en-tête Google Sheet) où stocker l'URL du PDF. Laissez vide si non utilisé.
                *   **Valeur à entrer dans AppSheet :** `"LienPDF"` ou `""`
                    *   *Explication :* Entrez le nom de l'en-tête de la colonne pour l'URL. **Mettez cette valeur entre guillemets doubles `"`**. Si vous n'utilisez pas cette fonctionnalité, mettez des guillemets vides : `""`.
                    *   *Exemple (utilisé) :* `"URL Facture"`
                    *   *Exemple (non utilisé) :* `""`

            7.  **Argument 7 : `pdfFilenameTemplate`**
                *   **Description :** Le modèle pour nommer le fichier PDF. Utilisez les placeholders `{{NomColonneSheet}}`.
                *   **Valeur à entrer dans AppSheet :** `"BL-{{ID Commande}}-{{NomClient}}.pdf"`
                    *   *Explication :* Définissez le format du nom de fichier. Les `{{NomColonneSheet}}` doivent correspondre aux en-têtes de votre Google Sheet. **Mettez toute la chaîne entre guillemets doubles `"`**.
                    *   *Exemple :* `"Facture_{{NumFacture}}_{{Date}}.pdf"`

            8.  **Argument 8 : `deleteTempDocStr`**
                *   **Description :** Indique s'il faut supprimer le fichier Google Doc temporaire (`true`) ou le conserver (`false`).
                *   **Valeur à entrer dans AppSheet :** `"false"` ou `"true"`
                    *   *Explication :* Entrez `true` ou `false` **sous forme de texte, entre guillemets doubles `"`**.
                    *   *Exemple (conserver le Doc) :* `"false"`
                    *   *Exemple (supprimer le Doc) :* `"true"`

            ---

            **Récapitulatif Visuel des Arguments dans AppSheet :**

            ```
            1: [VotreColonneCleUnique]
            2: "ID_MODELE_DOC"
            3: "ID_DOSSIER_DRIVE"
            4: "NOM_FEUILLE_SHEET"
            5: "NOM_COLONNE_CLE_SHEET"
            6: "NOM_COLONNE_LIEN_PDF_SHEET"  (ou "")
            7: "TEMPLATE_NOM_FICHIER_{{Colonne1}}_{{Colonne2}}.pdf"
            8: "false"  (ou "true")
            ```

            **ATTENTION :** L'ordre, les guillemets pour les textes fixes, et l'absence de guillemets pour l'expression `[Colonne]` sont **CRITIQUES**. Une erreur ici empêchera le script de fonctionner correctement.

        *   **(Optionnel) Capturer la sortie :** Cochez `Wait for return value`. La valeur retournée par le script (le nom final du fichier PDF) sera accessible dans les étapes suivantes via l'expression `[NomDeVotreEtapeScript].[Output]` (ex: `[Appel Script PDF Personnalisé].[Output]`).
    *   **(Optionnel) Utiliser la sortie :** Ajoutez une étape `Data: Set column values` *après* l'appel au script pour stocker le nom du fichier retourné dans votre colonne `NomFichierPDF` en utilisant l'expression `[Appel Script PDF Personnalisé].[Output]`.

3.  Sauvegardez votre application AppSheet.

## Utilisation

1.  Déclenchez l'événement configuré dans AppSheet.
2.  Synchronisez votre application.
3.  Le script s'exécute en arrière-plan. Vérifiez le dossier de destination Google Drive.
4.  Vérifiez la mise à jour éventuelle des colonnes dans Google Sheet/AppSheet.
5.  **En cas de problème :**
    *   Vérifiez **très attentivement les 8 arguments** configurés dans l'étape "Call a script" d'AppSheet (ordre, guillemets, valeurs exactes).
    *   Consultez les journaux d'exécution dans l'éditeur Google Apps Script (`Exécutions`).
    *   Consultez l'historique d'audit dans AppSheet (`Manage` > `Monitor` > `Audit History`).

## Suivi des Évolutions et Contributions (Interne)

*   Utiliser les **Issues** GitHub pour les demandes/bugs.
*   Utiliser les **branches** Git pour le développement.
*   Utiliser les **Pull Requests** pour la revue.
*   Rédiger des **messages de commit** clairs.

## Améliorations Futures Potentielles (To-Do)

*   [ ] Améliorer la gestion des erreurs (notifications).
*   [ ] Option pour envoyer par email.
*   [ ] Gérer des remplacements plus complexes (tableaux - difficile).
*   [ ] Créer une feuille de configuration dans Google Sheets pour stocker les paramètres (ID modèle, dossier, etc.) et les lire depuis AppSheet avant d'appeler le script, afin de simplifier la configuration de l'appel au script (moins d'arguments fixes à passer).
