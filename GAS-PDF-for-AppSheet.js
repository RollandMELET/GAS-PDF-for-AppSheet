// Version: 3.1 (Autonome, 9 paramètres, Logs améliorés, Sécurisé)

/**
 * @OnlyCurrentDoc // Peut être retiré si non pertinent pour standalone, mais ne gêne pas.
 * Activer le runtime V8 pour utiliser des fonctionnalités modernes.
 * Dans l'éditeur de script: Paramètres du projet > Cochez "Activer le moteur d'exécution Chrome V8"
 */

/**
 * Nettoie un nom de fichier des caractères interdits.
 * @param {string | null | undefined} filename Le nom de fichier potentiel.
 * @return {string} Le nom de fichier nettoyé, ou une chaîne vide si l'entrée est invalide.
 */
function sanitizeFilename(filename) {
    // Retourne une chaîne vide si l'entrée est null ou undefined pour éviter les erreurs
    if (filename === null || typeof filename === 'undefined') return '';
    // Convertit en string avant replace pour gérer d'autres types potentiels
    return filename.toString().replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * Trouve un nom de fichier unique dans un dossier donné en ajoutant (1), (2), etc. si nécessaire.
 * @param {GoogleAppsScript.Drive.Folder} folder Le dossier où vérifier l'unicité.
 * @param {string | null | undefined} desiredName Le nom de fichier initial souhaité (avec extension).
 * @return {string} Un nom de fichier garanti unique dans ce dossier.
 */
function getUniqueFilename(folder, desiredName) {
    // Fournir un nom par défaut si desiredName est vide, null ou undefined
    if (!desiredName) {
        desiredName = `Document_Sans_Nom_${new Date().getTime()}.pdf`;
        Logger.log(`Attention: Nom de fichier souhaité vide ou invalide, utilisation de: ${desiredName}`);
    } else {
        // S'assurer que desiredName est une chaîne
        desiredName = desiredName.toString();
    }

    let finalName = desiredName;
    let counter = 1;
    let files; // Déclarer files ici

    try {
      files = folder.getFilesByName(finalName); // Premier appel pour vérifier si le nom initial existe
    } catch (e) {
      Logger.log(`ERREUR critique lors de la recherche de fichier initial "${finalName}" dans getUniqueFilename: ${e}. Tentative avec nom basé sur le temps.`);
      // Retourner un nom basé sur le temps pour éviter un échec complet
      return sanitizeFilename(`Fichier_Erreur_${new Date().getTime()}.pdf`);
    }


    // Boucle while pour trouver un nom unique si le nom initial existe
    while (files.hasNext()) {
        let baseName = finalName;
        let extension = '';
        const dotIndex = finalName.lastIndexOf('.');
        // Sépare le nom de base et l'extension (si elle existe)
        if (dotIndex > 0 && dotIndex < finalName.length - 1) { // Vérifie que le point n'est pas le dernier caractère
            baseName = finalName.substring(0, dotIndex);
            extension = finalName.substring(dotIndex); // Inclut le point
        }

        // Gérer le cas où le nom contient déjà " (n)" pour éviter " (1) (1)"
        const match = baseName.match(/^(.*)\s\((\d+)\)$/);
        if (match) {
            baseName = match[1]; // Utiliser la base avant le suffixe (n)
            // Optionnellement, on pourrait essayer de continuer le compteur à partir de match[2] + 1,
            // mais recommencer avec `counter` est plus simple et robuste.
        }

        // Construire le nouveau nom potentiel avec le compteur
        finalName = `${baseName} (${counter})${extension}`;

        try {
          files = folder.getFilesByName(finalName); // Ré-évaluer l'existence avec le nouveau nom potentiel
        } catch (e) {
          Logger.log(`ERREUR critique lors de la recherche de fichier "${finalName}" dans la boucle while de getUniqueFilename: ${e}. Tentative avec nom basé sur le temps.`);
           // Retourner un nom basé sur le temps en cas d'erreur pendant la boucle
          return sanitizeFilename(`Fichier_Erreur_${baseName}_${new Date().getTime()}${extension}`);
        }

        counter++; // Incrémenter le compteur pour le prochain essai

        // Sécurité pour éviter une boucle infinie (par exemple, si Drive a un comportement inattendu)
        if (counter > 100) {
            Logger.log(`AVERTISSEMENT: Impossible de trouver un nom unique après 100 tentatives pour ${desiredName}. Utilisation d'un nom basé sur le temps pour éviter boucle infinie.`);
            finalName = sanitizeFilename(`${baseName}_${new Date().getTime()}${extension}`);
            break; // Sortir de la boucle while
        }
    }
    // Retourne le nom final trouvé (qui est garanti unique à ce moment)
    return finalName;
}


/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 * Tous les paramètres de configuration sont passés en arguments.
 * L'ORDRE DES PARAMÈTRES EST CRUCIAL lors de l'appel depuis AppSheet.
 *
 * @param {string} uniqueId La valeur de la clé unique de la ligne à traiter.
 * @param {string} spreadsheetId L'ID DU FICHIER Google Sheet contenant les données.
 * @param {string} templateDocId L'ID du document Google Docs servant de modèle.
 * @param {string} destinationFolderId L'ID du dossier Google Drive où sauvegarder le PDF.
 * @param {string} sheetName Le nom exact de la feuille (onglet) Google Sheet contenant les données.
 * @param {string} uniqueIdColumnNameInSheet Le nom exact de la colonne clé unique dans la feuille Sheet.
 * @param {string} pdfLinkColumnName Le nom de la colonne où stocker l'URL du PDF (laisser vide '' si non utilisé).
 * @param {string} pdfFilenameTemplate Le modèle pour le nom du fichier PDF (ex: "BL-{{ID}}.pdf").
 * @param {string} deleteTempDocStr Indique s'il faut supprimer le Doc temporaire ('true' ou 'false').
 *
 * @return {string} Le NOM FINAL du fichier PDF généré ou un message d'erreur propagé.
 */
function generatePdfFromTemplate(
    uniqueId,
    spreadsheetId,
    templateDocId,
    destinationFolderId,
    sheetName,
    uniqueIdColumnNameInSheet,
    pdfLinkColumnName,
    pdfFilenameTemplate,
    deleteTempDocStr
) {

    // --- Validation robuste des paramètres essentiels ---
    if (!uniqueId || String(uniqueId).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueId' est manquant ou vide.");
    if (!spreadsheetId || String(spreadsheetId).trim() === '') throw new Error("Erreur Script: L'argument 'spreadsheetId' (ID du fichier Sheet) est manquant ou vide.");
    if (!templateDocId || String(templateDocId).trim() === '') throw new Error("Erreur Script: L'argument 'templateDocId' est manquant ou vide.");
    if (!destinationFolderId || String(destinationFolderId).trim() === '') throw new Error("Erreur Script: L'argument 'destinationFolderId' est manquant ou vide.");
    if (!sheetName || String(sheetName).trim() === '') throw new Error("Erreur Script: L'argument 'sheetName' est manquant ou vide.");
    if (!uniqueIdColumnNameInSheet || String(uniqueIdColumnNameInSheet).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueIdColumnNameInSheet' est manquant ou vide.");
    if (!pdfFilenameTemplate || String(pdfFilenameTemplate).trim() === '') throw new Error("Erreur Script: L'argument 'pdfFilenameTemplate' est manquant ou vide.");

    // Conversion string 'true'/'false' en booléen (insensible à la casse et gère booléen direct aussi)
    const deleteTempDoc = (String(deleteTempDocStr).trim().toLowerCase() === 'true');

    Logger.log(`--- Début Génération PDF --- ID Ligne: ${uniqueId} ---`);
    Logger.log(`Paramètres Reçus: spreadsheetId=${spreadsheetId}, templateId=${templateDocId}, folderId=${destinationFolderId}, sheetName=${sheetName}, idColName=${uniqueIdColumnNameInSheet}, pdfLinkColName=${pdfLinkColumnName || 'N/A'}, nameTemplate=${pdfFilenameTemplate}, deleteTempDoc=${deleteTempDoc}`);

    let copiedFileId = null; // Pour pouvoir nettoyer en cas d'erreur précoce

    try { // --- Début du Try principal ---

        // 1. Accès aux données Google Sheet via son ID
        let ss;
        try {
            ss = SpreadsheetApp.openById(spreadsheetId);
            Logger.log(`Fichier Google Sheet ouvert avec succès (ID: ${spreadsheetId})`);
        } catch (e) {
            // Erreur spécifique si le fichier Sheet n'est pas accessible
            throw new Error(`Impossible d'ouvrir le fichier Google Sheet ID "${spreadsheetId}". Vérifiez l'ID et les permissions du script. Détails: ${e.message}`);
        }

        const sheet = ss.getSheetByName(sheetName);
        if (!sheet) {
            // Erreur si l'onglet n'existe pas dans le fichier Sheet
            throw new Error(`Feuille (onglet) "${sheetName}" non trouvée dans le fichier Sheet ID: ${spreadsheetId}. Vérifiez le nom exact (casse sensible).`);
        }
        Logger.log(`Accès réussi à la feuille "${sheetName}".`);

        const dataRange = sheet.getDataRange();
        const allData = dataRange.getValues(); // Récupère toutes les données [ligne][colonne]
        if (allData.length < 1) { // Vérifier s'il y a au moins une ligne (même juste les en-têtes)
           throw new Error(`La feuille "${sheetName}" est vide.`);
        }
        const headers = allData[0].map(h => h ? String(h).trim() : ''); // Nettoie les en-têtes (string, trim)
        Logger.log(`En-têtes trouvés: ${headers.join(' | ')}`);

        // Trouver l'index de la colonne clé (après nettoyage des en-têtes)
        const uniqueIdColNameClean = String(uniqueIdColumnNameInSheet).trim();
        const uniqueIdColIndex = headers.indexOf(uniqueIdColNameClean);
        if (uniqueIdColIndex === -1) {
            throw new Error(`Colonne clé "${uniqueIdColNameClean}" non trouvée dans les en-têtes nettoyés de la feuille "${sheetName}". Vérifiez le nom exact. En-têtes nettoyés: ${headers.join(', ')}`);
        }

        // Trouver l'index de la colonne lien PDF (optionnel)
        let pdfLinkColIndex = -1;
        const pdfLinkColNameClean = pdfLinkColumnName ? String(pdfLinkColumnName).trim() : '';
        if (pdfLinkColNameClean !== '') {
            pdfLinkColIndex = headers.indexOf(pdfLinkColNameClean);
            if (pdfLinkColIndex === -1) {
                Logger.log(`Attention: Colonne lien PDF "${pdfLinkColNameClean}" spécifiée mais non trouvée dans les en-têtes. Le lien ne sera pas enregistré.`);
            } else {
                 Logger.log(`Colonne lien PDF "${pdfLinkColNameClean}" trouvée à l'index ${pdfLinkColIndex}.`);
            }
        } else {
             Logger.log("Aucune colonne de lien PDF spécifiée.");
        }

        // 2. Trouver la ligne correspondant à l'ID unique
        let targetRowData = null;
        let targetRowIndexInData = -1; // Index dans le tableau allData (0-based)
        const uniqueIdStr = String(uniqueId); // Convertir l'ID reçu en string pour comparaison
        for (let i = 1; i < allData.length; i++) { // Commence à 1 pour ignorer les en-têtes
            const cellValue = allData[i][uniqueIdColIndex];
            // Comparaison après conversion en string des deux côtés
            if (cellValue !== null && cellValue !== undefined && String(cellValue) === uniqueIdStr) {
                targetRowData = allData[i];
                targetRowIndexInData = i;
                Logger.log(`Ligne trouvée pour ID "${uniqueIdStr}" à l'index de données ${i} (index feuille ${i + 1}).`);
                break; // Arrêter dès qu'on trouve la première correspondance
            }
        }
        if (!targetRowData) {
            throw new Error(`Aucune ligne trouvée avec l'ID unique "${uniqueIdStr}" dans la colonne "${uniqueIdColNameClean}" (index ${uniqueIdColIndex}). Vérifiez la valeur et le type de l'ID passé, et les données dans la feuille.`);
        }

        // 3. Préparer l'objet des placeholders { "{{Header}}": "Valeur" }
        const placeholders = {};
        headers.forEach((header, index) => {
            if (header !== '') { // Ne créer des placeholders que pour les en-têtes non vides
                const cellData = targetRowData[index];
                // Utilise une chaîne vide si la donnée est null ou undefined
                placeholders[`{{${header}}}`] = (cellData !== null && cellData !== undefined) ? cellData.toString() : '';
            }
        });

        // ===> LOGS AJOUTÉS POUR DÉBOGAGE DES DONNÉES ET PLACEHOLDERS <===
        try {
            // Log des données brutes (peut être tronqué par Logger si trop long)
            Logger.log(`Données de la ligne trouvée (index ${targetRowIndexInData}): ${JSON.stringify(targetRowData)}`);
            // Log des placeholders créés pour vérifier les clés et valeurs
            Logger.log(`Placeholders créés pour remplacement: ${JSON.stringify(placeholders, null, 2)}`);
        } catch (logError) {
            Logger.log(`AVERTISSEMENT: Erreur lors de la journalisation JSON des données/placeholders (taille/type?): ${logError}. Tentative de log partiel.`);
            try {
              Logger.log('Données brutes (10 premières colonnes): ' + targetRowData.slice(0, 10).join(' | '));
              Logger.log('Placeholders clés trouvées (' + Object.keys(placeholders).length + '): ' + Object.keys(placeholders).slice(0, 20).join(', ') + '...');
            } catch(fallbackLogError) {
              Logger.log('Erreur même lors du logging partiel des données/placeholders.');
            }
        }
        // ===> FIN DES LOGS AJOUTÉS <===

        // 4. Accéder au modèle et au dossier de destination
        let templateFile, destinationFolder;
        try {
            templateFile = DriveApp.getFileById(templateDocId);
            Logger.log(`Modèle Google Doc trouvé (ID: ${templateDocId}, Nom: ${templateFile.getName()})`);
        } catch (e) {
            throw new Error(`Impossible d'accéder au fichier modèle Google Docs ID "${templateDocId}". Vérifiez l'ID et les permissions. Détails: ${e.message}`);
        }
        try {
            destinationFolder = DriveApp.getFolderById(destinationFolderId);
            Logger.log(`Dossier de destination trouvé (ID: ${destinationFolderId}, Nom: ${destinationFolder.getName()})`);
        } catch (e) {
            throw new Error(`Impossible d'accéder au dossier de destination Google Drive ID "${destinationFolderId}". Vérifiez l'ID et les permissions. Détails: ${e.message}`);
        }

        // 4.1 Générer le nom de base souhaité pour le Doc temporaire
        let desiredTempDocName = `Temp_${pdfFilenameTemplate.replace(/\.pdf$/i, '')}`; // Enlever .pdf de la fin (insensible à la casse)
        for (const placeholderKey in placeholders) {
            // Regex pour remplacer globalement, en échappant les caractères spéciaux du placeholder
            const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
            desiredTempDocName = desiredTempDocName.replace(regex, placeholders[placeholderKey]);
        }
        desiredTempDocName = sanitizeFilename(desiredTempDocName) || `Temp_Doc_${uniqueId}`; // Nom par défaut si vide

        // 4.2 Obtenir un nom unique pour le Doc temporaire
        const uniqueTempDocName = getUniqueFilename(destinationFolder, desiredTempDocName);
        Logger.log(`Nom unique pour le Doc temporaire: ${uniqueTempDocName}`);

        // 4.3 Copier le modèle avec le nom unique
        const copiedFile = templateFile.makeCopy(uniqueTempDocName, destinationFolder);
        copiedFileId = copiedFile.getId(); // Stocker l'ID pour le nettoyage potentiel
        Logger.log(`Modèle copié avec succès (Nouveau Doc ID: ${copiedFileId})`);

        // 4.4 Ouvrir le document copié pour modification
        let copiedDoc;
        try {
           copiedDoc = DocumentApp.openById(copiedFileId);
        } catch (docOpenError) {
             throw new Error(`Impossible d'ouvrir le document temporaire copié ID: ${copiedFileId}. Détails: ${docOpenError.message}`);
             // Le nettoyage sera fait dans le bloc catch principal
        }
        Logger.log(`Document temporaire ouvert pour modification.`);

        // 5. Remplacer les placeholders dans la copie (Corps, En-tête, Pied de page)
        const body = copiedDoc.getBody();
        const header = copiedDoc.getHeader(); // Peut être null si pas d'en-tête
        const footer = copiedDoc.getFooter(); // Peut être null si pas de pied de page
        Logger.log("Début du remplacement des placeholders dans le document...");
        let replaceCount = 0;
        for (const placeholderKey in placeholders) {
            const value = placeholders[placeholderKey];
            // S'assurer que la valeur est une chaîne pour replaceText
            const valueStr = (value === null || typeof value === 'undefined') ? '' : value.toString();
            // Regex pour remplacer globalement, en échappant les caractères spéciaux
            const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
            try {
                let replacedInBody = body.replaceText(regex, valueStr);
                let replacedInHeader = header ? header.replaceText(regex, valueStr) : null;
                let replacedInFooter = footer ? footer.replaceText(regex, valueStr) : null;
                if (replacedInBody || replacedInHeader || replacedInFooter) {
                    replaceCount++;
                }
            } catch (replaceError) {
                Logger.log(`AVERTISSEMENT: Erreur lors du remplacement du placeholder ${placeholderKey}. Le script continue. Détails: ${replaceError.message}`);
            }
        }
        Logger.log(`Fin du remplacement. ${replaceCount} placeholders (au moins partiellement) remplacés.`);

        // 5.1 Sauvegarder et fermer le document modifié (critique avant de générer le PDF)
        copiedDoc.saveAndClose();
        Logger.log(`Document temporaire ${uniqueTempDocName} sauvegardé et fermé.`);

        // 6. Générer le nom final du fichier PDF à partir du template
        let desiredPdfName = pdfFilenameTemplate;
        for (const placeholderKey in placeholders) {
            const value = placeholders[placeholderKey];
            const valueStr = (value === null || typeof value === 'undefined') ? '' : value.toString();
            const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
            desiredPdfName = desiredPdfName.replace(regex, valueStr);
        }
        desiredPdfName = sanitizeFilename(desiredPdfName); // Nettoyer le nom de fichier résultant

        // 6.1 Vérification finale du nom du PDF (doit se terminer par .pdf)
        if (!desiredPdfName || !desiredPdfName.toLowerCase().endsWith('.pdf')) {
            Logger.log(`AVERTISSEMENT: Le nom de fichier PDF généré "${desiredPdfName}" est invalide ou manque l'extension .pdf. Utilisation d'un nom par défaut.`);
            // Utiliser l'ID unique pour garantir un nom de fichier au moins minimal et valide
            desiredPdfName = sanitizeFilename(`Document_${uniqueId}.pdf`);
             if (!desiredPdfName) { // Ultime fallback
                desiredPdfName = `Document_Fallback_${new Date().getTime()}.pdf`;
             }
        }

        // 6.2 Obtenir un nom de fichier unique pour le PDF dans le dossier
        const uniquePdfName = getUniqueFilename(destinationFolder, desiredPdfName);
        Logger.log(`Nom de fichier PDF final unique déterminé: ${uniquePdfName}`);

        // 7. Créer le fichier PDF à partir du Doc temporaire et le sauvegarder avec le nom unique
        let pdfFile, pdfUrl, pdfId;
        try {
            const tempDocFileForPdf = DriveApp.getFileById(copiedFileId); // Récupérer le fichier
            const pdfBlob = tempDocFileForPdf.getAs(MimeType.PDF);
            pdfFile = destinationFolder.createFile(pdfBlob).setName(uniquePdfName);
            pdfUrl = pdfFile.getUrl();
            pdfId = pdfFile.getId();
            Logger.log(`Fichier PDF généré avec succès : ${uniquePdfName} (ID: ${pdfId}), URL: ${pdfUrl}`);
        } catch (pdfError) {
             // Tenter de supprimer le Doc temporaire si l'option est activée, même si la création PDF échoue
             if (deleteTempDoc && copiedFileId) {
                try { DriveApp.getFileById(copiedFileId).setTrashed(true); Logger.log(`Nettoyage: Doc temporaire ${copiedFileId} supprimé après échec création PDF.`); } catch(cleanupError){}
             }
             throw new Error(`Erreur lors de la création ou de la sauvegarde du fichier PDF "${uniquePdfName}". Détails: ${pdfError.message}`);
        }

        // 8. Mettre à jour la feuille Google Sheet avec le lien PDF (si configuré et possible)
        // Vérifier si la colonne a été trouvée (index >= 0) et si la ligne cible est valide (index >= 1)
        if (pdfLinkColNameClean !== '' && pdfLinkColIndex !== -1 && targetRowIndexInData !== -1) {
            try {
                // +1 car getRange est 1-based, targetRowIndexInData est 0-based pour les données (mais correspond à ligne - 1)
                // => Ligne Sheet = targetRowIndexInData + 1
                // +1 car pdfLinkColIndex est 0-based, getRange est 1-based
                // => Colonne Sheet = pdfLinkColIndex + 1
                const targetCell = sheet.getRange(targetRowIndexInData + 1, pdfLinkColIndex + 1);
                targetCell.setValue(pdfUrl); // Ou pdfId si préféré
                SpreadsheetApp.flush(); // Forcer l'écriture immédiate pour éviter les délais
                Logger.log(`Lien PDF (URL) ajouté à la feuille "${sheetName}", cellule ${targetCell.getA1Notation()}, colonne "${pdfLinkColNameClean}".`);
            } catch (sheetUpdateError) {
                 Logger.log(`AVERTISSEMENT: Impossible de mettre à jour la colonne "${pdfLinkColNameClean}" dans la feuille Sheet. Détails: ${sheetUpdateError.message}`);
                 // Ne pas arrêter le script pour cette erreur optionnelle
            }
        } else if (pdfLinkColNameClean !== '') {
             // Log si la colonne ou la ligne n'ont pas été trouvées correctement
             Logger.log(`Le lien PDF n'a pas été enregistré car la colonne "${pdfLinkColNameClean}" n'a pas été trouvée (index: ${pdfLinkColIndex}) ou l'index de ligne cible est invalide (index données: ${targetRowIndexInData}).`);
        }

        // 9. Supprimer le Document Google Doc temporaire (si l'option est activée)
        if (deleteTempDoc) {
            try {
                // Utiliser l'ID stocké pour s'assurer qu'on cible le bon fichier
                const fileToDelete = DriveApp.getFileById(copiedFileId);
                const tempName = fileToDelete.getName(); // Récupérer le nom pour le log
                fileToDelete.setTrashed(true); // Met à la corbeille
                Logger.log(`Document Google Doc intermédiaire ${tempName} (ID: ${copiedFileId}) marqué pour suppression.`);
            } catch (deleteError) {
                // Logguer comme un avertissement si la suppression échoue
                Logger.log(`AVERTISSEMENT: Impossible de supprimer le document temporaire (ID: ${copiedFileId}). Il restera dans le dossier ${destinationFolder.getName()}. Détails: ${deleteError.message}`);
            }
        } else {
             // Logguer explicitement que la suppression n'a pas été faite
             Logger.log(`Le document temporaire (ID: ${copiedFileId}) n'a pas été supprimé (option désactivée).`);
        }

        // 10. Retourner le NOM FINAL et UNIQUE du fichier PDF généré à AppSheet
        Logger.log(`--- Fin Génération PDF --- Succès. Retournant le nom: ${uniquePdfName} ---`);
        return uniquePdfName;

    } catch (error) { // --- Fin du Try principal, début du Catch global ---
        Logger.log(`--- ERREUR GLOBALE --- ID Ligne: ${uniqueId} ---`);
        Logger.log(`Message Erreur: ${error.message}`);
        // Logguer la stack trace pour un débogage plus détaillé
        if (error.stack) {
            Logger.log(`Stack Trace: ${error.stack}`);
        }

        // Tentative de nettoyage du fichier temporaire si une erreur s'est produite APRES sa création
        if (copiedFileId && deleteTempDoc) {
           try {
               DriveApp.getFileById(copiedFileId).setTrashed(true);
               Logger.log(`Nettoyage après erreur: Doc temporaire ${copiedFileId} supprimé.`);
           } catch (cleanupError) {
               Logger.log(`Nettoyage après erreur: Impossible de supprimer le doc temporaire ${copiedFileId}. Détails: ${cleanupError.message}`);
           }
        }

        // Renvoyer une erreur claire à AppSheet pour indiquer l'échec
        // Le message original de l'erreur interne est souvent le plus utile pour le diagnostic
        throw new Error(`Erreur script: ${error.message}`);
    } // --- Fin du Catch principal ---
} // --- Fin de la fonction generatePdfFromTemplate ---

// ==========================================================================
// Fonction de test pour forcer la demande d'autorisation (si nécessaire)
// Exécutez cette fonction manuellement depuis l'éditeur de script
// pour accorder les permissions requises par le script principal.
// ==========================================================================
function testAutorisationsNecesaires() {
  let testResult = '';
  try {
    // Tester l'accès à un Spreadsheet via ID (remplacez par un ID valide pour vous)
    const testSheetId = "METTRE_ICI_UN_ID_DE_FICHIER_SHEET_VALIDE";
    const ss = SpreadsheetApp.openById(testSheetId);
    testResult += `Accès Sheet (${testSheetId}) OK. Nom: ${ss.getName()}\n`;

    // Tester l'accès à un Document via ID (remplacez par un ID valide pour vous)
    const testDocId = "METTRE_ICI_UN_ID_DE_FICHIER_DOC_VALIDE";
    const doc = DocumentApp.openById(testDocId);
    testResult += `Accès Doc (${testDocId}) OK. Nom: ${doc.getName()}\n`;

    // Tester l'accès à un dossier Drive via ID (remplacez par un ID valide pour vous)
    const testFolderId = "METTRE_ICI_UN_ID_DE_DOSSIER_DRIVE_VALIDE";
    const folder = DriveApp.getFolderById(testFolderId);
    testResult += `Accès Folder (${testFolderId}) OK. Nom: ${folder.getName()}\n`;

    // Tester la création/suppression d'un fichier temporaire dans ce dossier
    const tempFileName = `_test_autorisation_${new Date().getTime()}.txt`;
    const tempFile = folder.createFile(tempFileName, 'Test autorisation Drive');
    const tempFileId = tempFile.getId();
    testResult += `Création fichier test (${tempFileName}) OK.\n`;
    DriveApp.getFileById(tempFileId).setTrashed(true);
    testResult += `Suppression fichier test OK.\n`;

    Logger.log("Test des autorisations réussi :\n" + testResult);
    SpreadsheetApp.getUi().alert("Test des autorisations", "Toutes les autorisations nécessaires semblent accordées.\n\n" + testResult, SpreadsheetApp.getUi().ButtonSet.OK);

  } catch (e) {
    Logger.log(`ERREUR lors du test des autorisations : ${e}`);
    // Afficher l'erreur à l'utilisateur aussi
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<pre>${e.stack}</pre>`), `Erreur lors du test: ${e.message}`);
    // Relancer l'erreur pour l'enregistrer dans les exécutions
    throw new Error(`ERREUR test autorisations: ${e.message}`);
  }
}
// ==========================================================================
