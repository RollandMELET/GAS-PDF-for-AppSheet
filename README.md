# Générateur PDF Personnalisé pour AppSheet via Google Apps Script

Ce projet contient un script Google Apps Script conçu pour être déclenché depuis une application AppSheet. Son objectif est de générer des fichiers PDF personnalisés en utilisant un modèle Google Docs et les données d'une ligne spécifique d'une feuille Google Sheets liée à AppSheet.

Il a été créé pour surmonter les limitations de la génération PDF native d'AppSheet et offrir un contrôle total sur la mise en page et l'apparence du document final.

**Dépôt privé destiné au suivi des évolutions et des demandes.**

## Problème Résolu

La génération de PDF native dans AppSheet offre peu de contrôle sur la mise en page fine (polices spécifiques, positionnement exact des éléments, en-têtes/pieds de page complexes, etc.). Cette solution permet d'utiliser toute la puissance de mise en page de Google Docs comme base pour les PDF.

## Fonctionnalités Clés

*   Utilise un fichier Google Docs comme modèle de document.
*   Remplace dynamiquement des placeholders (marqueurs type `{{NomColonne}}`) dans le modèle par les données correspondantes d'une ligne Google Sheet.
*   Gère les placeholders dans le corps, l'en-tête et le pied de page du document.
*   Crée une copie du modèle pour chaque génération afin de ne pas altérer l'original.
*   Génère un fichier PDF à partir du document Google Docs rempli.
*   Sauvegarde le PDF généré dans un dossier Google Drive spécifié.
*   Nomme le fichier PDF selon un modèle configurable (ex: `BL-{{ID}}.pdf`) incluant des données de la ligne.
*   Gère automatiquement les conflits de noms de fichiers en ajoutant un suffixe numérique (ex: `NomFichier (1).pdf`, `NomFichier (2).pdf`) si un fichier du même nom existe déjà.
*   **Optionnel:** Met à jour une colonne spécifiée dans Google Sheet avec l'URL du PDF généré.
*   **Optionnel:** Supprime le fichier Google Docs intermédiaire après la génération du PDF.
*   Retourne le nom final du fichier PDF généré à AppSheet (peut être utilisé dans des étapes suivantes du Bot).

## Technologies Utilisées

*   **AppSheet:** Plateforme de création d'applications et interface utilisateur pour déclencher le script.
*   **Google Apps Script (GAS):** Langage de script pour l'automatisation dans l'écosystème Google.
*   **Google Docs:** Pour la création et la gestion des modèles de documents.
*   **Google Sheets:** Comme source de données pour AppSheet et le script.
*   **Google Drive:** Pour le stockage des modèles, des documents intermédiaires (optionnel) et des PDF finaux.

## Configuration Requise

Pour utiliser ce script, une configuration préalable est nécessaire dans plusieurs services Google et dans AppSheet.

### 1. Google Docs (Modèle)

1.  Créez un nouveau Google Doc qui servira de modèle.
2.  Mettez en page ce document exactement comme vous souhaitez que le PDF final apparaisse (logos, polices, couleurs, tableaux, etc.).
3.  Insérez des **placeholders** (marqueurs) aux endroits où les données dynamiques doivent être insérées. Utilisez le format `{{NomDeLaColonne}}`.
    *   **Important:** Le `NomDeLaColonne` doit correspondre **exactement** à l'en-tête de colonne dans votre Google Sheet (sensible à la casse, aux espaces - il est recommandé d'utiliser des noms de colonnes sans espaces ni caractères spéciaux dans Google Sheet, ex: `NomClient` plutôt que `Nom Client`).
4.  Notez l'**ID du document modèle**. Vous le trouverez dans l'URL du document (entre `/d/` et `/edit`).

### 2. Google Drive

1.  Créez un dossier dans Google Drive où les PDF générés seront sauvegardés.
2.  Notez l'**ID de ce dossier**. Vous le trouverez dans l'URL lorsque vous naviguez dans ce dossier.

### 3. Google Sheets

1.  Assurez-vous que votre feuille Google Sheet (utilisée par AppSheet) contient toutes les colonnes nécessaires correspondant aux placeholders de votre modèle Google Doc.
2.  Vérifiez que vous avez une colonne avec une **clé unique** pour chaque ligne (requis par AppSheet).
3.  **Optionnel:** Ajoutez une colonne pour stocker l'URL du PDF généré (ex: `LienPDF`, type `Url` ou `File` dans AppSheet).
4.  **Optionnel:** Ajoutez une colonne pour stocker le nom du fichier PDF généré (ex: `NomFichierPDF`, type `Text` dans AppSheet).

### 4. Google Apps Script

1.  Ouvrez votre Google Sheet.
2.  Allez dans `Outils` > `Éditeur de scripts`.
3.  Copiez le contenu du fichier `.js` de ce dépôt et collez-le dans l'éditeur, en remplaçant le code existant.
4.  **Configurez l'objet `CONFIG`** en haut du script avec vos propres informations :
    *   `TEMPLATE_DOC_ID`: ID de votre modèle Google Doc.
    *   `DESTINATION_FOLDER_ID`: ID de votre dossier de destination Google Drive.
    *   `SHEET_NAME`: Nom exact de votre feuille Google Sheet.
    *   `UNIQUE_ID_COLUMN_NAME`: Nom exact de votre colonne clé unique.
    *   `PDF_LINK_COLUMN_NAME`: Nom de la colonne pour l'URL du PDF (ou `''` si non utilisée).
    *   `PDF_FILENAME_TEMPLATE`: Modèle pour le nom du fichier PDF (utilisez les placeholders `{{NomColonne}}`).
    *   `DELETE_TEMP_DOC`: Mettez à `true` pour supprimer le Doc temporaire, `false` pour le conserver.
5.  Sauvegardez le projet de script (icône disquette).
6.  **(Recommandé)** Activez le moteur d'exécution V8 : `Paramètres du projet` (icône engrenage) > Cochez `Activer le moteur d'exécution Chrome V8`.
7.  **Autorisations :** La première fois que le script sera exécuté (ou lors de la configuration dans AppSheet), Google demandera l'autorisation d'accéder à Drive, Docs et Sheets. Accordez ces autorisations.

### 5. AppSheet

1.  **Activer les Scripts :** Dans l'éditeur AppSheet, allez dans `Security` > `Options` et activez "Allow apps script execution...".
2.  **Créer une Automatisation (Bot) :**
    *   Allez dans `Automation` > `Bots`.
    *   Créez un nouveau Bot (ex: "Générer PDF Personnalisé").
    *   **Configurez l'Événement (Event) :** Choisissez ce qui déclenche le PDF (ex: un changement de donnée - `Data Change` avec `Updates` et une condition sur un statut, ou une action utilisateur - `Action`). Sélectionnez la bonne table. Ajoutez une condition si nécessaire (ex: `[Statut]="Prêt"`).
    *   **Configurez le Processus (Process) :**
        *   Ajoutez une étape (`Add step`). Nommez-la (ex: "Appel Script PDF").
        *   Choisissez le type d'étape `Call a script`.
        *   Sélectionnez votre projet Google Apps Script.
        *   Nom de la fonction : `generatePdfFromTemplate`.
        *   **Arguments :** Ajoutez un argument et utilisez une expression pour passer la clé unique de la ligne, par exemple `[VotreColonneCleUnique]`.
        *   **(Optionnel) Capturer la sortie :** Cochez `Wait for return value`. La valeur retournée (le nom du fichier PDF) sera accessible dans les étapes suivantes via `[Appel Script PDF].[Output]`.
    *   **(Optionnel) Ajouter une étape suivante :** Si vous voulez stocker le nom du fichier retourné, ajoutez une étape `Data: Set column values` après l'appel au script, et mettez à jour votre colonne `NomFichierPDF` avec la valeur `[Appel Script PDF].[Output]`.
3.  Sauvegardez votre application AppSheet.

## Utilisation

1.  Déclenchez l'événement configuré dans AppSheet (ex: changez le statut d'une ligne, cliquez sur un bouton d'action).
2.  Synchronisez votre application AppSheet.
3.  Attendez quelques instants que le script s'exécute.
4.  Vérifiez le dossier de destination dans Google Drive pour le nouveau fichier PDF.
5.  Si configuré, vérifiez la mise à jour des colonnes `LienPDF` et/ou `NomFichierPDF` dans votre Google Sheet et AppSheet (après synchronisation).
6.  En cas de problème, vérifiez :
    *   Les journaux d'exécution dans l'éditeur Google Apps Script (`Exécutions`).
    *   L'historique d'audit dans AppSheet (`Manage` > `Monitor` > `Audit History`) pour voir si le Bot a été déclenché et si l'appel au script a réussi ou échoué.

## Suivi des Évolutions et Contributions (Interne)

*   **Demandes d'évolution / Bugs :** Utiliser les **Issues** GitHub de ce dépôt pour documenter les nouvelles fonctionnalités souhaitées ou les bugs rencontrés.
*   **Développement :** Utiliser des **branches** Git pour développer de nouvelles fonctionnalités ou corriger des bugs (ex: `feature/nouvelle-option`, `fix/bug-nom-fichier`).
*   **Intégration :** Utiliser les **Pull Requests** pour revoir les changements avant de les fusionner dans la branche principale (`main` ou `master`).
*   **Messages de Commit :** Rédiger des messages de commit clairs et descriptifs expliquant le *pourquoi* du changement.

## Améliorations Futures Potentielles (To-Do)

*   [ ] Améliorer la gestion des erreurs (ex: envoyer une notification en cas d'échec).
*   [ ] Ajouter une option pour envoyer directement le PDF par email.
*   [ ] Explorer des logiques de remplacement plus complexes (ex: tableaux dynamiques basés sur des tables enfants - difficile avec `replaceText`).
*   [ ] Créer une interface de configuration plus conviviale (ex: via une feuille Sheet dédiée).