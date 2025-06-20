// Version: 3.4 (Autonome, 9 paramètres, API AppSheet ou Google Sheets, Utilise getDisplayValues, Remplacement string simple, Logs+, Sécurisé)

/**
 * @OnlyCurrentDoc
 * Activer le runtime V8.
 */

/**
 * Nettoie un nom de fichier des caractères interdits.
 * @param {string | null | undefined} filename Le nom de fichier potentiel.
 * @return {string} Le nom de fichier nettoyé, ou une chaîne vide si l'entrée est invalide.
 */
function sanitizeFilename(filename) {
    if (filename === null || typeof filename === 'undefined') return '';
    return filename.toString().replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * Trouve un nom de fichier unique dans un dossier donné en ajoutant (1), (2), etc. si nécessaire.
 * @param {GoogleAppsScript.Drive.Folder} folder Le dossier où vérifier l'unicité.
 * @param {string | null | undefined} desiredName Le nom de fichier initial souhaité (avec extension).
 * @return {string} Un nom de fichier garanti unique dans ce dossier.
 */
function getUniqueFilename(folder, desiredName) {
    if (!desiredName) { desiredName = `Document_Sans_Nom_${new Date().getTime()}.pdf`; Logger.log(`Attention: Nom fichier vide/invalide, utilisation de: ${desiredName}`); } else { desiredName = desiredName.toString(); }
    let finalName = desiredName; let counter = 1; let files; try { files = folder.getFilesByName(finalName); } catch (e) { Logger.log(`ERREUR recherche fichier initial "${finalName}": ${e}`); return sanitizeFilename(`Fichier_Erreur_${new Date().getTime()}.pdf`); }
    while (files.hasNext()) { let baseName = finalName; let extension = ''; const dotIndex = finalName.lastIndexOf('.'); if (dotIndex > 0 && dotIndex < finalName.length - 1) { baseName = finalName.substring(0, dotIndex); extension = finalName.substring(dotIndex); } const match = baseName.match(/^(.*)\s\((\d+)\)$/); if (match) { baseName = match[1]; } finalName = `${baseName} (${counter})${extension}`; try { files = folder.getFilesByName(finalName); } catch (e) { Logger.log(`ERREUR recherche fichier "${finalName}" boucle: ${e}`); return sanitizeFilename(`Fichier_Erreur_${baseName}_${new Date().getTime()}${extension}`); } counter++; if (counter > 100) { Logger.log(`AVERTISSEMENT: 100 tentatives nom unique pour ${desiredName}. Nom basé sur temps utilisé.`); finalName = sanitizeFilename(`${baseName}_${new Date().getTime()}${extension}`); break; } } return finalName;
}


/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 * Lit les données affichées depuis Google Sheet.
 * @param {string} uniqueId La valeur de la clé unique de la ligne.
 * @param {string} spreadsheetId L'ID DU FICHIER Google Sheet.
 * @param {string} templateDocId L'ID du modèle Google Docs.
 * @param {string} destinationFolderId L'ID du dossier Google Drive.
 * @param {string} sheetName Le nom exact de la feuille (onglet).
 * @param {string} uniqueIdColumnNameInSheet Le nom exact de la colonne clé unique dans Sheet.
 * @param {string} pdfLinkColumnName Le nom de la colonne URL PDF (ou '').
 * @param {string} pdfFilenameTemplate Le modèle pour le nom du PDF.
 * @param {string} deleteTempDocStr Supprimer Doc temporaire ('true'/'false').
 * @return {string} Le NOM FINAL du fichier PDF généré.
 */
function generatePdfFromTemplate( uniqueId, spreadsheetId, templateDocId, destinationFolderId, sheetName, uniqueIdColumnNameInSheet, pdfLinkColumnName, pdfFilenameTemplate, deleteTempDocStr ) {

    // --- Validation robuste des paramètres essentiels ---
    if (!uniqueId || String(uniqueId).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueId' est manquant ou vide.");
    if (!spreadsheetId || String(spreadsheetId).trim() === '') throw new Error("Erreur Script: L'argument 'spreadsheetId' est manquant ou vide.");
    if (!templateDocId || String(templateDocId).trim() === '') throw new Error("Erreur Script: L'argument 'templateDocId' est manquant ou vide.");
    if (!destinationFolderId || String(destinationFolderId).trim() === '') throw new Error("Erreur Script: L'argument 'destinationFolderId' est manquant ou vide.");
    if (!sheetName || String(sheetName).trim() === '') throw new Error("Erreur Script: L'argument 'sheetName' est manquant ou vide.");
    if (!uniqueIdColumnNameInSheet || String(uniqueIdColumnNameInSheet).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueIdColumnNameInSheet' est manquant ou vide.");
    if (!pdfFilenameTemplate || String(pdfFilenameTemplate).trim() === '') throw new Error("Erreur Script: L'argument 'pdfFilenameTemplate' est manquant ou vide.");

    const deleteTempDoc = (String(deleteTempDocStr).trim().toLowerCase() === 'true');
    Logger.log(`--- Début Génération PDF --- ID Ligne: ${uniqueId} ---`);
    Logger.log(`Paramètres Reçus: spreadsheetId=${spreadsheetId}, templateId=${templateDocId}, folderId=${destinationFolderId}, sheetName=${sheetName}, idColName=${uniqueIdColumnNameInSheet}, pdfLinkColName=${pdfLinkColumnName || 'N/A'}, nameTemplate=${pdfFilenameTemplate}, deleteTempDoc=${deleteTempDoc}`);
    let copiedFileId = null;

    try { // --- Début du Try principal ---

        // 1. Accès aux données Google Sheet via son ID
        let ss; try { ss = SpreadsheetApp.openById(spreadsheetId); Logger.log(`Fichier Google Sheet ouvert (ID: ${spreadsheetId})`); } catch (e) { throw new Error(`Impossible d'ouvrir Sheet ID "${spreadsheetId}". Vérifiez ID/permissions. Détails: ${e.message}`); }
        const sheet = ss.getSheetByName(sheetName); if (!sheet) { throw new Error(`Feuille "${sheetName}" non trouvée dans Sheet ID: ${spreadsheetId}.`); } Logger.log(`Accès feuille "${sheetName}".`);
        const dataRange = sheet.getDataRange();

        // ===> CHANGEMENT MAJEUR ICI : Utiliser getDisplayValues() <===
        const allData = dataRange.getDisplayValues();
        Logger.log("Données lues avec getDisplayValues() (format affiché).");
        // ===> FIN DU CHANGEMENT MAJEUR <===

        if (allData.length < 1) { throw new Error(`Feuille "${sheetName}" vide.`); }
        const headers = allData[0].map(h => h ? String(h).trim() : ''); Logger.log(`En-têtes: ${headers.join(' | ')}`);
        const uniqueIdColNameClean = String(uniqueIdColumnNameInSheet).trim(); const uniqueIdColIndex = headers.indexOf(uniqueIdColNameClean); if (uniqueIdColIndex === -1) { throw new Error(`Colonne clé "${uniqueIdColNameClean}" non trouvée. En-têtes: ${headers.join(', ')}`); }
        let pdfLinkColIndex = -1; const pdfLinkColNameClean = pdfLinkColumnName ? String(pdfLinkColumnName).trim() : ''; if (pdfLinkColNameClean !== '') { pdfLinkColIndex = headers.indexOf(pdfLinkColNameClean); if (pdfLinkColIndex === -1) { Logger.log(`Attention: Colonne lien PDF "${pdfLinkColNameClean}" non trouvée.`); } else { Logger.log(`Col lien PDF "${pdfLinkColNameClean}" index ${pdfLinkColIndex}.`); } } else { Logger.log("Pas de colonne lien PDF spécifiée."); }

        // 2. Trouver la ligne
        let targetRowData = null; let targetRowIndexInData = -1; const uniqueIdStr = String(uniqueId);
        for (let i = 1; i < allData.length; i++) { const cellValue = allData[i][uniqueIdColIndex];
          // Comparaison sur les valeurs affichées (qui sont déjà des strings)
          if (cellValue !== null && cellValue !== undefined && cellValue === uniqueIdStr) { targetRowData = allData[i]; targetRowIndexInData = i; Logger.log(`Ligne trouvée pour ID "${uniqueIdStr}" index ${i}.`); break; } }
        if (!targetRowData) { throw new Error(`ID unique "${uniqueIdStr}" non trouvé dans colonne "${uniqueIdColNameClean}".`); }

        // 3. Préparer les placeholders (les valeurs sont déjà des strings)
        const placeholders = {}; headers.forEach((header, index) => { if (header !== '') { const cellData = targetRowData[index]; placeholders[`{{${header}}}`] = (cellData !== null && cellData !== undefined) ? cellData : ''; } }); // Directement la valeur (string) ou ''

        // ===> LOGS POUR DÉBOGAGE <===
        try { Logger.log(`Données ligne ${targetRowIndexInData} (Display Values): ${JSON.stringify(targetRowData)}`); Logger.log(`Placeholders créés: ${JSON.stringify(placeholders, null, 2)}`); } catch (logError) { Logger.log(`AVERTISSEMENT log JSON: ${logError}`); try { Logger.log('Données (10): ' + targetRowData.slice(0, 10).join(' | ')); Logger.log('Placeholders clés (' + Object.keys(placeholders).length + '): ' + Object.keys(placeholders).slice(0, 20).join(', ') + '...'); } catch(e){} }
        // ===> FIN LOGS <===

        // 4. Accéder modèle et dossier
        let templateFile, destinationFolder;
        try { templateFile = DriveApp.getFileById(templateDocId); Logger.log(`Modèle Doc trouvé (ID: ${templateDocId})`); } catch (e) { throw new Error(`Accès modèle Doc ID "${templateDocId}" échoué. Détails: ${e.message}`); }
        try { destinationFolder = DriveApp.getFolderById(destinationFolderId); Logger.log(`Dossier destination trouvé (ID: ${destinationFolderId})`); } catch (e) { throw new Error(`Accès dossier Drive ID "${destinationFolderId}" échoué. Détails: ${e.message}`); }

        // 4.1 Générer nom Doc temporaire
        let desiredTempDocName = `Temp_${pdfFilenameTemplate.replace(/\.pdf$/i, '')}`; for (const placeholderKey in placeholders) { const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); desiredTempDocName = desiredTempDocName.replace(regex, placeholders[placeholderKey]); } desiredTempDocName = sanitizeFilename(desiredTempDocName) || `Temp_Doc_${uniqueId}`;
        const uniqueTempDocName = getUniqueFilename(destinationFolder, desiredTempDocName); Logger.log(`Nom unique Doc temporaire: ${uniqueTempDocName}`);

        // 4.3 Copier modèle
        const copiedFile = templateFile.makeCopy(uniqueTempDocName, destinationFolder); copiedFileId = copiedFile.getId(); Logger.log(`Modèle copié (Doc ID: ${copiedFileId})`);

        // 4.4 Ouvrir copie
        let copiedDoc; try { copiedDoc = DocumentApp.openById(copiedFileId); } catch (docOpenError) { throw new Error(`Ouverture Doc temporaire ID ${copiedFileId} échouée. Détails: ${docOpenError.message}`); } Logger.log(`Doc temporaire ouvert.`);

        // 5. Remplacer placeholders (méthode string simple)
        const body = copiedDoc.getBody(); const header = copiedDoc.getHeader(); const footer = copiedDoc.getFooter();
        Logger.log("Début remplacement placeholders (méthode string simple)...");
        for (const placeholderKey in placeholders) {
            const valueStr = placeholders[placeholderKey]; // C'est déjà une string
            const searchString = placeholderKey;
            try {
                 //Logger.log(`Tentative remplacement : "${searchString}" par "${valueStr}"`); // Peut être réactivé si besoin, mais très verbeux
                 body.replaceText(searchString, valueStr);
                 if (header) header.replaceText(searchString, valueStr);
                 if (footer) footer.replaceText(searchString, valueStr);
            } catch (replaceError) {
                Logger.log(`AVERTISSEMENT: Erreur replaceText pour ${searchString}. Continue. Détails: ${replaceError.message}`);
            }
        }
        Logger.log("Fin tentatives remplacement (méthode string simple).");

        // 5.1 Sauvegarder et fermer
        copiedDoc.saveAndClose(); Logger.log(`Doc temporaire ${uniqueTempDocName} sauvegardé et fermé.`);

        // 6. Générer nom PDF final
        let desiredPdfName = pdfFilenameTemplate; for (const placeholderKey in placeholders) { const valueStr = placeholders[placeholderKey]; const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); desiredPdfName = desiredPdfName.replace(regex, valueStr); } desiredPdfName = sanitizeFilename(desiredPdfName);
        if (!desiredPdfName || !desiredPdfName.toLowerCase().endsWith('.pdf')) { Logger.log(`AVERTISSEMENT: Nom PDF généré "${desiredPdfName}" invalide. Nom défaut utilisé.`); desiredPdfName = sanitizeFilename(`Document_${uniqueId}.pdf`); if (!desiredPdfName) { desiredPdfName = `Document_Fallback_${new Date().getTime()}.pdf`; } }
        const uniquePdfName = getUniqueFilename(destinationFolder, desiredPdfName); Logger.log(`Nom PDF final unique: ${uniquePdfName}`);

        // 7. Créer PDF
        let pdfFile, pdfUrl, pdfId; try { const tempDocFileForPdf = DriveApp.getFileById(copiedFileId); const pdfBlob = tempDocFileForPdf.getAs(MimeType.PDF); pdfFile = destinationFolder.createFile(pdfBlob).setName(uniquePdfName); pdfUrl = pdfFile.getUrl(); pdfId = pdfFile.getId(); Logger.log(`PDF généré: ${uniquePdfName} (ID: ${pdfId}), URL: ${pdfUrl}`); } catch (pdfError) { if (deleteTempDoc && copiedFileId) { try { DriveApp.getFileById(copiedFileId).setTrashed(true); Logger.log(`Nettoyage: Doc temp ${copiedFileId} supprimé après échec PDF.`); } catch(e){} } throw new Error(`Erreur création/sauvegarde PDF "${uniquePdfName}". Détails: ${pdfError.message}`); }

        // 8. Mettre à jour Sheet
        if (pdfLinkColNameClean !== '' && pdfLinkColIndex !== -1 && targetRowIndexInData !== -1) { try { const targetCell = sheet.getRange(targetRowIndexInData + 1, pdfLinkColIndex + 1); targetCell.setValue(pdfUrl); SpreadsheetApp.flush(); Logger.log(`Lien PDF ajouté feuille "${sheetName}", cellule ${targetCell.getA1Notation()}, col "${pdfLinkColNameClean}".`); } catch (sheetUpdateError) { Logger.log(`AVERTISSEMENT: Échec màj col ${pdfLinkColNameClean} Sheet. Détails: ${sheetUpdateError.message}`); } } else if (pdfLinkColNameClean !== '') { Logger.log(`Lien PDF non enregistré (col: ${pdfLinkColIndex}, ligne: ${targetRowIndexInData}).`); }

        // 9. Supprimer Doc temporaire
        if (deleteTempDoc) { try { const fileToDelete = DriveApp.getFileById(copiedFileId); const tempName = fileToDelete.getName(); fileToDelete.setTrashed(true); Logger.log(`Doc temp ${tempName} (ID: ${copiedFileId}) marqué pour suppression.`); } catch (deleteError) { Logger.log(`AVERTISSEMENT: Échec suppression Doc temp (ID: ${copiedFileId}). Détails: ${deleteError.message}`); } } else { Logger.log(`Doc temp (ID: ${copiedFileId}) non supprimé (option false).`); }

        // 10. Retourner nom PDF
        Logger.log(`--- Fin Génération PDF --- Succès. Retournant nom: ${uniquePdfName} ---`);
        return uniquePdfName;

    } catch (error) { // --- Catch global ---
        Logger.log(`--- ERREUR GLOBALE --- ID Ligne: ${uniqueId} --- Message: ${error.message}`);
        if (error.stack) { Logger.log(`Stack Trace: ${error.stack}`); }
        if (copiedFileId && deleteTempDoc) { try { DriveApp.getFileById(copiedFileId).setTrashed(true); Logger.log(`Nettoyage après erreur: Doc temp ${copiedFileId} supprimé.`); } catch (e) { Logger.log(`Nettoyage après erreur: Échec suppression Doc temp ${copiedFileId}. Détails: ${e.message}`); } }
        throw new Error(`Erreur script: ${error.message}`);
    } // --- Fin Catch global ---
} // --- Fin fonction generatePdfFromTemplate ---

// ==========================================================================
// Fonction utilitaire pour récupérer les données depuis l'API AppSheet
// ==========================================================================
/**
 * Récupère les données d'un enregistrement via l'API AppSheet.
 * @param {string} appId ID de l'application AppSheet.
 * @param {string} accessKey Clé d'accès API AppSheet.
 * @param {string} tableName Nom de la table AppSheet.
 * @param {string} uniqueIdColumnName Nom de la colonne ID unique dans AppSheet.
 * @param {string} uniqueId Valeur de l'ID unique à rechercher.
 * @return {Object} Les données de l'enregistrement trouvé et l'index si mise à jour nécessaire.
 * @throws {Error} Si la requête échoue ou si l'enregistrement n'est pas trouvé.
 */
function fetchDataFromAppSheetAPI(appId, accessKey, tableName, uniqueIdColumnName, uniqueId) {
    Logger.log(`Récupération données via API AppSheet - AppID: ${appId}, Table: ${tableName}, Colonne ID: ${uniqueIdColumnName}, Valeur ID: ${uniqueId}`);
    
    // Construction de l'URL de l'API
    const apiUrl = `https://api.appsheet.com/api/v2/apps/${appId}/tables/${tableName}/Action`;
    Logger.log(`URL API: ${apiUrl}`);
    
    // Construction du corps de la requête avec filtre sur l'ID unique
    const requestBody = {
        Action: "Find",
        Properties: {
            Filter: `${uniqueIdColumnName} = '${uniqueId}'`
        }
    };
    
    // Configuration de la requête HTTP
    const options = {
        method: 'post',
        contentType: 'application/json',
        headers: {
            'ApplicationAccessKey': accessKey
        },
        payload: JSON.stringify(requestBody),
        muteHttpExceptions: true
    };
    
    try {
        // Exécution de la requête HTTP
        Logger.log('Envoi requête API AppSheet...');
        const response = UrlFetchApp.fetch(apiUrl, options);
        
        // Analyse de la réponse
        const responseCode = response.getResponseCode();
        const responseText = response.getContentText();
        Logger.log(`Réponse API - Code: ${responseCode}`);
        
        // Vérification du code de réponse
        if (responseCode !== 200) {
            Logger.log(`Erreur API AppSheet - Code: ${responseCode}, Détails: ${responseText}`);
            throw new Error(`Erreur API AppSheet (${responseCode}): ${responseText}`);
        }
        
        // Parsing de la réponse JSON
        const responseData = JSON.parse(responseText);
        Logger.log(`Réponse API AppSheet parsée: ${JSON.stringify(responseData).substring(0, 200)}...`);
        
        // Extraction des données de l'enregistrement
        let recordData;
        if (Array.isArray(responseData)) {
            // Format tableau direct
            recordData = responseData.length > 0 ? responseData[0] : null;
        } else if (responseData.Rows && Array.isArray(responseData.Rows)) {
            // Format avec Rows et Properties
            recordData = responseData.Rows.length > 0 ? responseData.Rows[0] : null;
        } else {
            // Format non reconnu
            Logger.log(`Format de réponse API non reconnu: ${responseText.substring(0, 200)}...`);
            throw new Error('Format de réponse API AppSheet non reconnu');
        }
        
        // Vérification qu'un enregistrement a été trouvé
        if (!recordData) {
            Logger.log(`Aucun enregistrement trouvé pour ID: ${uniqueId}`);
            throw new Error(`Aucun enregistrement trouvé dans la table '${tableName}' avec ${uniqueIdColumnName} = '${uniqueId}'`);
        }
        
        // Nettoyage et conversion des données
        const cleanedData = {};
        for (const key in recordData) {
            if (Object.prototype.hasOwnProperty.call(recordData, key)) {
                // Convertir null ou undefined en chaîne vide
                cleanedData[key] = (recordData[key] !== null && recordData[key] !== undefined) 
                    ? String(recordData[key]) 
                    : '';
            }
        }
        
        Logger.log(`Données récupérées pour ID ${uniqueId}: ${Object.keys(cleanedData).join(', ')}`);
        return {
            data: cleanedData,
            recordId: uniqueId
        };
    } catch (error) {
        Logger.log(`ERREUR lors de la récupération API AppSheet: ${error.message}`);
        if (error.stack) {
            Logger.log(`Stack trace: ${error.stack}`);
        }
        throw new Error(`Erreur lors de la récupération des données via API AppSheet: ${error.message}`);
    }
}

// ==========================================================================
// Fonction principale avec API AppSheet
// ==========================================================================
/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 *
// Version: 3.4 (Autonome, 9 paramètres, API AppSheet ou Google Sheets, Utilise getDisplayValues, Remplacement string simple, Logs+, Sécurisé)

/**
 * @OnlyCurrentDoc
 * Activer le runtime V8.
 */

/**
 * Nettoie un nom de fichier des caractères interdits.
 * @param {string | null | undefined} filename Le nom de fichier potentiel.
 * @return {string} Le nom de fichier nettoyé, ou une chaîne vide si l'entrée est invalide.
 */
function sanitizeFilename(filename) {
    if (filename === null || typeof filename === 'undefined') return '';
    return filename.toString().replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * Trouve un nom de fichier unique dans un dossier donné en ajoutant (1), (2), etc. si nécessaire.
 * @param {GoogleAppsScript.Drive.Folder} folder Le dossier où vérifier l'unicité.
 * @param {string | null | undefined} desiredName Le nom de fichier initial souhaité (avec extension).
 * @return {string} Un nom de fichier garanti unique dans ce dossier.
 */
function getUniqueFilename(folder, desiredName) {
    if (!desiredName) { desiredName = `Document_Sans_Nom_${new Date().getTime()}.pdf`; Logger.log(`Attention: Nom fichier vide/invalide, utilisation de: ${desiredName}`); } else { desiredName = desiredName.toString(); }
    let finalName = desiredName; let counter = 1; let files; try { files = folder.getFilesByName(finalName); } catch (e) { Logger.log(`ERREUR recherche fichier initial "${finalName}": ${e}`); return sanitizeFilename(`Fichier_Erreur_${new Date().getTime()}.pdf`); }
    while (files.hasNext()) { let baseName = finalName; let extension = ''; const dotIndex = finalName.lastIndexOf('.'); if (dotIndex > 0 && dotIndex < finalName.length - 1) { baseName = finalName.substring(0, dotIndex); extension = finalName.substring(dotIndex); } const match = baseName.match(/^(.*)\s\((\d+)\)$/); if (match) { baseName = match[1]; } finalName = `${baseName} (${counter})${extension}`; try { files = folder.getFilesByName(finalName); } catch (e) { Logger.log(`ERREUR recherche fichier "${finalName}" boucle: ${e}`); return sanitizeFilename(`Fichier_Erreur_${baseName}_${new Date().getTime()}${extension}`); } counter++; if (counter > 100) { Logger.log(`AVERTISSEMENT: 100 tentatives nom unique pour ${desiredName}. Nom basé sur temps utilisé.`); finalName = sanitizeFilename(`${baseName}_${new Date().getTime()}${extension}`); break; } } return finalName;
}


/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 * Lit les données affichées depuis Google Sheet.
 * @param {string} uniqueId La valeur de la clé unique de la ligne.
 * @param {string} spreadsheetId L'ID DU FICHIER Google Sheet.
 * @param {string} templateDocId L'ID du modèle Google Docs.
 * @param {string} destinationFolderId L'ID du dossier Google Drive.
 * @param {string} sheetName Le nom exact de la feuille (onglet).
 * @param {string} uniqueIdColumnNameInSheet Le nom exact de la colonne clé unique dans Sheet.
 * @param {string} pdfLinkColumnName Le nom de la colonne URL PDF (ou '').
 * @param {string} pdfFilenameTemplate Le modèle pour le nom du PDF.
 * @param {string} deleteTempDocStr Supprimer Doc temporaire ('true'/'false').
 * @return {string} Le NOM FINAL du fichier PDF généré.
 */
function generatePdfFromTemplate( uniqueId, spreadsheetId, templateDocId, destinationFolderId, sheetName, uniqueIdColumnNameInSheet, pdfLinkColumnName, pdfFilenameTemplate, deleteTempDocStr ) {

    // --- Validation robuste des paramètres essentiels ---
    if (!uniqueId || String(uniqueId).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueId' est manquant ou vide.");
    if (!spreadsheetId || String(spreadsheetId).trim() === '') throw new Error("Erreur Script: L'argument 'spreadsheetId' est manquant ou vide.");
    if (!templateDocId || String(templateDocId).trim() === '') throw new Error("Erreur Script: L'argument 'templateDocId' est manquant ou vide.");
    if (!destinationFolderId || String(destinationFolderId).trim() === '') throw new Error("Erreur Script: L'argument 'destinationFolderId' est manquant ou vide.");
    if (!sheetName || String(sheetName).trim() === '') throw new Error("Erreur Script: L'argument 'sheetName' est manquant ou vide.");
    if (!uniqueIdColumnNameInSheet || String(uniqueIdColumnNameInSheet).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueIdColumnNameInSheet' est manquant ou vide.");
    if (!pdfFilenameTemplate || String(pdfFilenameTemplate).trim() === '') throw new Error("Erreur Script: L'argument 'pdfFilenameTemplate' est manquant ou vide.");

    const deleteTempDoc = (String(deleteTempDocStr).trim().toLowerCase() === 'true');
    Logger.log(`--- Début Génération PDF --- ID Ligne: ${uniqueId} ---`);
    Logger.log(`Paramètres Reçus: spreadsheetId=${spreadsheetId}, templateId=${templateDocId}, folderId=${destinationFolderId}, sheetName=${sheetName}, idColName=${uniqueIdColumnNameInSheet}, pdfLinkColName=${pdfLinkColumnName || 'N/A'}, nameTemplate=${pdfFilenameTemplate}, deleteTempDoc=${deleteTempDoc}`);
    let copiedFileId = null;

    try { // --- Début du Try principal ---

        // 1. Accès aux données Google Sheet via son ID
        let ss; try { ss = SpreadsheetApp.openById(spreadsheetId); Logger.log(`Fichier Google Sheet ouvert (ID: ${spreadsheetId})`); } catch (e) { throw new Error(`Impossible d'ouvrir Sheet ID "${spreadsheetId}". Vérifiez ID/permissions. Détails: ${e.message}`); }
        const sheet = ss.getSheetByName(sheetName); if (!sheet) { throw new Error(`Feuille "${sheetName}" non trouvée dans Sheet ID: ${spreadsheetId}.`); } Logger.log(`Accès feuille "${sheetName}".`);
        const dataRange = sheet.getDataRange();

        // ===> CHANGEMENT MAJEUR ICI : Utiliser getDisplayValues() <===
        const allData = dataRange.getDisplayValues();
        Logger.log("Données lues avec getDisplayValues() (format affiché).");
        // ===> FIN DU CHANGEMENT MAJEUR <===

        if (allData.length < 1) { throw new Error(`Feuille "${sheetName}" vide.`); }
        const headers = allData[0].map(h => h ? String(h).trim() : ''); Logger.log(`En-têtes: ${headers.join(' | ')}`);
        const uniqueIdColNameClean = String(uniqueIdColumnNameInSheet).trim(); const uniqueIdColIndex = headers.indexOf(uniqueIdColNameClean); if (uniqueIdColIndex === -1) { throw new Error(`Colonne clé "${uniqueIdColNameClean}" non trouvée. En-têtes: ${headers.join(', ')}`); }
        let pdfLinkColIndex = -1; const pdfLinkColNameClean = pdfLinkColumnName ? String(pdfLinkColumnName).trim() : ''; if (pdfLinkColNameClean !== '') { pdfLinkColIndex = headers.indexOf(pdfLinkColNameClean); if (pdfLinkColIndex === -1) { Logger.log(`Attention: Colonne lien PDF "${pdfLinkColNameClean}" non trouvée.`); } else { Logger.log(`Col lien PDF "${pdfLinkColNameClean}" index ${pdfLinkColIndex}.`); } } else { Logger.log("Pas de colonne lien PDF spécifiée."); }

        // 2. Trouver la ligne
        let targetRowData = null; let targetRowIndexInData = -1; const uniqueIdStr = String(uniqueId);
        for (let i = 1; i < allData.length; i++) { const cellValue = allData[i][uniqueIdColIndex];
          // Comparaison sur les valeurs affichées (qui sont déjà des strings)
          if (cellValue !== null && cellValue !== undefined && cellValue === uniqueIdStr) { targetRowData = allData[i]; targetRowIndexInData = i; Logger.log(`Ligne trouvée pour ID "${uniqueIdStr}" index ${i}.`); break; } }
        if (!targetRowData) { throw new Error(`ID unique "${uniqueIdStr}" non trouvé dans colonne "${uniqueIdColNameClean}".`); }

        // 3. Préparer les placeholders (les valeurs sont déjà des strings)
        const placeholders = {}; headers.forEach((header, index) => { if (header !== '') { const cellData = targetRowData[index]; placeholders[`{{${header}}}`] = (cellData !== null && cellData !== undefined) ? cellData : ''; } }); // Directement la valeur (string) ou ''

        // ===> LOGS POUR DÉBOGAGE <===
        try { Logger.log(`Données ligne ${targetRowIndexInData} (Display Values): ${JSON.stringify(targetRowData)}`); Logger.log(`Placeholders créés: ${JSON.stringify(placeholders, null, 2)}`); } catch (logError) { Logger.log(`AVERTISSEMENT log JSON: ${logError}`); try { Logger.log('Données (10): ' + targetRowData.slice(0, 10).join(' | ')); Logger.log('Placeholders clés (' + Object.keys(placeholders).length + '): ' + Object.keys(placeholders).slice(0, 20).join(', ') + '...'); } catch(e){} }
        // ===> FIN LOGS <===

        // 4. Accéder modèle et dossier
        let templateFile, destinationFolder;
        try { templateFile = DriveApp.getFileById(templateDocId); Logger.log(`Modèle Doc trouvé (ID: ${templateDocId})`); } catch (e) { throw new Error(`Accès modèle Doc ID "${templateDocId}" échoué. Détails: ${e.message}`); }
        try { destinationFolder = DriveApp.getFolderById(destinationFolderId); Logger.log(`Dossier destination trouvé (ID: ${destinationFolderId})`); } catch (e) { throw new Error(`Accès dossier Drive ID "${destinationFolderId}" échoué. Détails: ${e.message}`); }

        // 4.1 Générer nom Doc temporaire
        let desiredTempDocName = `Temp_${pdfFilenameTemplate.replace(/\.pdf$/i, '')}`; for (const placeholderKey in placeholders) { const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); desiredTempDocName = desiredTempDocName.replace(regex, placeholders[placeholderKey]); } desiredTempDocName = sanitizeFilename(desiredTempDocName) || `Temp_Doc_${uniqueId}`;
        const uniqueTempDocName = getUniqueFilename(destinationFolder, desiredTempDocName); Logger.log(`Nom unique Doc temporaire: ${uniqueTempDocName}`);

        // 4.3 Copier modèle
        const copiedFile = templateFile.makeCopy(uniqueTempDocName, destinationFolder); copiedFileId = copiedFile.getId(); Logger.log(`Modèle copié (Doc ID: ${copiedFileId})`);

        // 4.4 Ouvrir copie
        let copiedDoc; try { copiedDoc = DocumentApp.openById(copiedFileId); } catch (docOpenError) { throw new Error(`Ouverture Doc temporaire ID ${copiedFileId} échouée. Détails: ${docOpenError.message}`); } Logger.log(`Doc temporaire ouvert.`);

        // 5. Remplacer placeholders (méthode string simple)
        const body = copiedDoc.getBody(); const header = copiedDoc.getHeader(); const footer = copiedDoc.getFooter();
        Logger.log("Début remplacement placeholders (méthode string simple)...");
        for (const placeholderKey in placeholders) {
            const valueStr = placeholders[placeholderKey]; // C'est déjà une string
            const searchString = placeholderKey;
            try {
                 //Logger.log(`Tentative remplacement : "${searchString}" par "${valueStr}"`); // Peut être réactivé si besoin, mais très verbeux
                 body.replaceText(searchString, valueStr);
                 if (header) header.replaceText(searchString, valueStr);
                 if (footer) footer.replaceText(searchString, valueStr);
            } catch (replaceError) {
                Logger.log(`AVERTISSEMENT: Erreur replaceText pour ${searchString}. Continue. Détails: ${replaceError.message}`);
            }
        }
        Logger.log("Fin tentatives remplacement (méthode string simple).");

        // 5.1 Sauvegarder et fermer
        copiedDoc.saveAndClose(); Logger.log(`Doc temporaire ${uniqueTempDocName} sauvegardé et fermé.`);

        // 6. Générer nom PDF final
        let desiredPdfName = pdfFilenameTemplate; for (const placeholderKey in placeholders) { const valueStr = placeholders[placeholderKey]; const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); desiredPdfName = desiredPdfName.replace(regex, valueStr); } desiredPdfName = sanitizeFilename(desiredPdfName);
        if (!desiredPdfName || !desiredPdfName.toLowerCase().endsWith('.pdf')) { Logger.log(`AVERTISSEMENT: Nom PDF généré "${desiredPdfName}" invalide. Nom défaut utilisé.`); desiredPdfName = sanitizeFilename(`Document_${uniqueId}.pdf`); if (!desiredPdfName) { desiredPdfName = `Document_Fallback_${new Date().getTime()}.pdf`; } }
        const uniquePdfName = getUniqueFilename(destinationFolder, desiredPdfName); Logger.log(`Nom PDF final unique: ${uniquePdfName}`);

        // 7. Créer PDF
        let pdfFile, pdfUrl, pdfId; try { const tempDocFileForPdf = DriveApp.getFileById(copiedFileId); const pdfBlob = tempDocFileForPdf.getAs(MimeType.PDF); pdfFile = destinationFolder.createFile(pdfBlob).setName(uniquePdfName); pdfUrl = pdfFile.getUrl(); pdfId = pdfFile.getId(); Logger.log(`PDF généré: ${uniquePdfName} (ID: ${pdfId}), URL: ${pdfUrl}`); } catch (pdfError) { if (deleteTempDoc && copiedFileId) { try { DriveApp.getFileById(copiedFileId).setTrashed(true); Logger.log(`Nettoyage: Doc temp ${copiedFileId} supprimé après échec PDF.`); } catch(e){} } throw new Error(`Erreur création/sauvegarde PDF "${uniquePdfName}". Détails: ${pdfError.message}`); }

        // 8. Mettre à jour Sheet
        if (pdfLinkColNameClean !== '' && pdfLinkColIndex !== -1 && targetRowIndexInData !== -1) { try { const targetCell = sheet.getRange(targetRowIndexInData + 1, pdfLinkColIndex + 1); targetCell.setValue(pdfUrl); SpreadsheetApp.flush(); Logger.log(`Lien PDF ajouté feuille "${sheetName}", cellule ${targetCell.getA1Notation()}, col "${pdfLinkColNameClean}".`); } catch (sheetUpdateError) { Logger.log(`AVERTISSEMENT: Échec màj col ${pdfLinkColNameClean} Sheet. Détails: ${sheetUpdateError.message}`); } } else if (pdfLinkColNameClean !== '') { Logger.log(`Lien PDF non enregistré (col: ${pdfLinkColIndex}, ligne: ${targetRowIndexInData}).`); }

        // 9. Supprimer Doc temporaire
        if (deleteTempDoc) { try { const fileToDelete = DriveApp.getFileById(copiedFileId); const tempName = fileToDelete.getName(); fileToDelete.setTrashed(true); Logger.log(`Doc temp ${tempName} (ID: ${copiedFileId}) marqué pour suppression.`); } catch (deleteError) { Logger.log(`AVERTISSEMENT: Échec suppression Doc temp (ID: ${copiedFileId}). Détails: ${deleteError.message}`); } } else { Logger.log(`Doc temp (ID: ${copiedFileId}) non supprimé (option false).`); }

        // 10. Retourner nom PDF
        Logger.log(`--- Fin Génération PDF --- Succès. Retournant nom: ${uniquePdfName} ---`);
        return uniquePdfName;

    } catch (error) { // --- Catch global ---
        Logger.log(`--- ERREUR GLOBALE --- ID Ligne: ${uniqueId} --- Message: ${error.message}`);
        if (error.stack) { Logger.log(`Stack Trace: ${error.stack}`); }
        if (copiedFileId && deleteTempDoc) { try { DriveApp.getFileById(copiedFileId).setTrashed(true); Logger.log(`Nettoyage après erreur: Doc temp ${copiedFileId} supprimé.`); } catch (e) { Logger.log(`Nettoyage après erreur: Échec suppression Doc temp ${copiedFileId}. Détails: ${e.message}`); } }
        throw new Error(`Erreur script: ${error.message}`);
    } // --- Fin Catch global ---
} // --- Fin fonction generatePdfFromTemplate ---

// ==========================================================================
// Fonction principale avec API AppSheet
// ==========================================================================
/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 * Utilise l'API AppSheet pour récupérer les données.
 * @param {string} uniqueId La valeur de la clé unique de l'enregistrement.
 * @param {string} appsheetAppId L'ID de l'application AppSheet.
 * @param {string} appsheetAccessKey La clé d'accès API AppSheet.
 * @param {string} tableName Le nom de la table AppSheet.
 * @param {string} templateDocId L'ID du modèle Google Docs.
 * @param {string} destinationFolderId L'ID du dossier Google Drive.
 * @param {string} uniqueIdColumnName Le nom de la colonne clé unique dans AppSheet.
 * @param {string} pdfFilenameTemplate Le modèle pour le nom du PDF.
 * @param {string} deleteTempDocStr Supprimer Doc temporaire ('true'/'false').
 * @return {string} Le NOM FINAL du fichier PDF généré.
 */
function generatePdfFromTemplateAPI_AppSheetUse(uniqueId, appsheetAppId, appsheetAccessKey, tableName, templateDocId, destinationFolderId, uniqueIdColumnName, pdfFilenameTemplate, deleteTempDocStr) {
    
    // --- Validation robuste des paramètres essentiels ---
    if (!uniqueId || String(uniqueId).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueId' est manquant ou vide.");
    if (!appsheetAppId || String(appsheetAppId).trim() === '') throw new Error("Erreur Script: L'argument 'appsheetAppId' est manquant ou vide.");
    if (!appsheetAccessKey || String(appsheetAccessKey).trim() === '') throw new Error("Erreur Script: L'argument 'appsheetAccessKey' est manquant ou vide.");
    if (!tableName || String(tableName).trim() === '') throw new Error("Erreur Script: L'argument 'tableName' est manquant ou vide.");
    if (!templateDocId || String(templateDocId).trim() === '') throw new Error("Erreur Script: L'argument 'templateDocId' est manquant ou vide.");
    if (!destinationFolderId || String(destinationFolderId).trim() === '') throw new Error("Erreur Script: L'argument 'destinationFolderId' est manquant ou vide.");
    if (!uniqueIdColumnName || String(uniqueIdColumnName).trim() === '') throw new Error("Erreur Script: L'argument 'uniqueIdColumnName' est manquant ou vide.");
    if (!pdfFilenameTemplate || String(pdfFilenameTemplate).trim() === '') throw new Error("Erreur Script: L'argument 'pdfFilenameTemplate' est manquant ou vide.");

    const deleteTempDoc = (String(deleteTempDocStr).trim().toLowerCase() === 'true');
    Logger.log(`--- Début Génération PDF via API AppSheet --- ID: ${uniqueId} ---`);
    Logger.log(`Paramètres Reçus: appId=${appsheetAppId}, tableName=${tableName}, uniqueId=${uniqueId}, templateId=${templateDocId}, folderId=${destinationFolderId}, idColName=${uniqueIdColumnName}, nameTemplate=${pdfFilenameTemplate}, deleteTempDoc=${deleteTempDoc}`);
    let copiedFileId = null;

    try { // --- Début du Try principal ---

        // 1. Récupérer les données via l'API AppSheet
        Logger.log(`Récupération des données pour ID ${uniqueId} depuis l'API AppSheet...`);
        const recordData = fetchDataFromAppSheetAPI(appsheetAppId, appsheetAccessKey, tableName, uniqueIdColumnName, uniqueId);
        
        if (!recordData || !recordData.data) {
            throw new Error(`Aucune donnée récupérée pour l'ID ${uniqueId}.`);
        }
        
        Logger.log(`Données récupérées avec succès pour ID ${uniqueId}.`);

        // 2. Préparer les placeholders
        const placeholders = {};
        for (const key in recordData.data) {
            if (Object.prototype.hasOwnProperty.call(recordData.data, key)) {
                // Le format API est déjà { "nom": "valeur" }, nous voulons { "{{nom}}": "valeur" }
                placeholders[`{{${key}}}`] = recordData.data[key];
            }
        }

        // ===> LOGS POUR DÉBOGAGE <===
        try { 
            Logger.log(`Données récupérées: ${JSON.stringify(recordData.data)}`); 
            Logger.log(`Placeholders créés: ${JSON.stringify(placeholders, null, 2)}`); 
        } catch (logError) { 
            Logger.log(`AVERTISSEMENT log JSON: ${logError}`); 
            try { 
                Logger.log('Placeholders clés (' + Object.keys(placeholders).length + '): ' + Object.keys(placeholders).slice(0, 20).join(', ') + '...'); 
            } catch(e){} 
        }
        // ===> FIN LOGS <===

        // 3. Accéder au modèle et au dossier
        let templateFile, destinationFolder;
        try { 
            templateFile = DriveApp.getFileById(templateDocId); 
            Logger.log(`Modèle Doc trouvé (ID: ${templateDocId})`); 
        } catch (e) { 
            throw new Error(`Accès modèle Doc ID "${templateDocId}" échoué. Détails: ${e.message}`); 
        }
        
        try { 
            destinationFolder = DriveApp.getFolderById(destinationFolderId); 
            Logger.log(`Dossier destination trouvé (ID: ${destinationFolderId})`); 
        } catch (e) { 
            throw new Error(`Accès dossier Drive ID "${destinationFolderId}" échoué. Détails: ${e.message}`); 
        }

        // 4. Générer nom Doc temporaire
        let desiredTempDocName = `Temp_${pdfFilenameTemplate.replace(/\.pdf$/i, '')}`; 
        for (const placeholderKey in placeholders) { 
            const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); 
            desiredTempDocName = desiredTempDocName.replace(regex, placeholders[placeholderKey]); 
        } 
        desiredTempDocName = sanitizeFilename(desiredTempDocName) || `Temp_Doc_${uniqueId}`;
        const uniqueTempDocName = getUniqueFilename(destinationFolder, desiredTempDocName); 
        Logger.log(`Nom unique Doc temporaire: ${uniqueTempDocName}`);

        // 5. Copier modèle
        const copiedFile = templateFile.makeCopy(uniqueTempDocName, destinationFolder); 
        copiedFileId = copiedFile.getId(); 
        Logger.log(`Modèle copié (Doc ID: ${copiedFileId})`);

        // 6. Ouvrir copie
        let copiedDoc; 
        try { 
            copiedDoc = DocumentApp.openById(copiedFileId); 
        } catch (docOpenError) { 
            throw new Error(`Ouverture Doc temporaire ID ${copiedFileId} échouée. Détails: ${docOpenError.message}`); 
        } 
        Logger.log(`Doc temporaire ouvert.`);

        // 7. Remplacer placeholders (méthode string simple)
        const body = copiedDoc.getBody(); 
        const header = copiedDoc.getHeader(); 
        const footer = copiedDoc.getFooter();
        
        Logger.log("Début remplacement placeholders (méthode string simple)...");
        for (const placeholderKey in placeholders) {
            const valueStr = placeholders[placeholderKey];
            const searchString = placeholderKey;
            try {
                body.replaceText(searchString, valueStr);
                if (header) header.replaceText(searchString, valueStr);
                if (footer) footer.replaceText(searchString, valueStr);
            } catch (replaceError) {
                Logger.log(`AVERTISSEMENT: Erreur replaceText pour ${searchString}. Continue. Détails: ${replaceError.message}`);
            }
        }
        Logger.log("Fin tentatives remplacement (méthode string simple).");

        // 8. Sauvegarder et fermer
        copiedDoc.saveAndClose(); 
        Logger.log(`Doc temporaire ${uniqueTempDocName} sauvegardé et fermé.`);

        // 9. Générer nom PDF final
        let desiredPdfName = pdfFilenameTemplate; 
        for (const placeholderKey in placeholders) { 
            const valueStr = placeholders[placeholderKey]; 
            const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); 
            desiredPdfName = desiredPdfName.replace(regex, valueStr); 
        } 
        desiredPdfName = sanitizeFilename(desiredPdfName);
        
        if (!desiredPdfName || !desiredPdfName.toLowerCase().endsWith('.pdf')) { 
            Logger.log(`AVERTISSEMENT: Nom PDF généré "${desiredPdfName}" invalide. Nom défaut utilisé.`); 
            desiredPdfName = sanitizeFilename(`Document_${uniqueId}.pdf`); 
            if (!desiredPdfName) { 
                desiredPdfName = `Document_Fallback_${new Date().getTime()}.pdf`; 
            } 
        }
        
        const uniquePdfName = getUniqueFilename(destinationFolder, desiredPdfName); 
        Logger.log(`Nom PDF final unique: ${uniquePdfName}`);

        // 10. Créer PDF
        let pdfFile, pdfUrl, pdfId; 
        try { 
            const tempDocFileForPdf = DriveApp.getFileById(copiedFileId); 
            const pdfBlob = tempDocFileForPdf.getAs(MimeType.PDF); 
            pdfFile = destinationFolder.createFile(pdfBlob).setName(uniquePdfName); 
            pdfUrl = pdfFile.getUrl(); 
            pdfId = pdfFile.getId(); 
            Logger.log(`PDF généré: ${uniquePdfName} (ID: ${pdfId}), URL: ${pdfUrl}`); 
        } catch (pdfError) { 
            if (deleteTempDoc && copiedFileId) { 
                try { 
                    DriveApp.getFileById(copiedFileId).setTrashed(true); 
                    Logger.log(`Nettoyage: Doc temp ${copiedFileId} supprimé après échec PDF.`); 
                } catch(e){} 
            } 
            throw new Error(`Erreur création/sauvegarde PDF "${uniquePdfName}". Détails: ${pdfError.message}`); 
        }

        // 11. Supprimer Doc temporaire
        if (deleteTempDoc) { 
            try { 
                const fileToDelete = DriveApp.getFileById(copiedFileId); 
                const tempName = fileToDelete.getName(); 
                fileToDelete.setTrashed(true); 
                Logger.log(`Doc temp ${tempName} (ID: ${copiedFileId}) marqué pour suppression.`); 
            } catch (deleteError) { 
                Logger.log(`AVERTISSEMENT: Échec suppression Doc temp (ID: ${copiedFileId}). Détails: ${deleteError.message}`); 
            } 
        } else { 
            Logger.log(`Doc temp (ID: ${copiedFileId}) non supprimé (option false).`); 
        }

        // 12. Retourner nom PDF
        Logger.log(`--- Fin Génération PDF API AppSheet --- Succès. Retournant nom: ${uniquePdfName} ---`);
        return uniquePdfName;

    } catch (error) { // --- Catch global ---
        Logger.log(`--- ERREUR GLOBALE API AppSheet --- ID: ${uniqueId} --- Message: ${error.message}`);
        if (error.stack) { 
            Logger.log(`Stack Trace: ${error.stack}`); 
        }
        if (copiedFileId && deleteTempDoc) { 
            try { 
                DriveApp.getFileById(copiedFileId).setTrashed(true); 
                Logger.log(`Nettoyage après erreur: Doc temp ${copiedFileId} supprimé.`); 
            } catch (e) { 
                Logger.log(`Nettoyage après erreur: Échec suppression Doc temp ${copiedFileId}. Détails: ${e.message}`); 
            } 
        }
        throw new Error(`Erreur script API AppSheet: ${error.message}`);
    } // --- Fin Catch global ---
} // --- Fin fonction generatePdfFromTemplateAPI_AppSheetUse ---

// ==========================================================================
// Fonction de test pour forcer la demande d'autorisation (si nécessaire)
// ==========================================================================
function testAutorisationsNecesaires() { let testResult = ''; try { const testSheetId = "METTRE_ICI_UN_ID_DE_FICHIER_SHEET_VALIDE"; const ss = SpreadsheetApp.openById(testSheetId); testResult += `Accès Sheet (${testSheetId}) OK. Nom: ${ss.getName()}\n`; const testDocId = "METTRE_ICI_UN_ID_DE_FICHIER_DOC_VALIDE"; const doc = DocumentApp.openById(testDocId); testResult += `Accès Doc (${testDocId}) OK. Nom: ${doc.getName()}\n`; const testFolderId = "METTRE_ICI_UN_ID_DE_DOSSIER_DRIVE_VALIDE"; const folder = DriveApp.getFolderById(testFolderId); testResult += `Accès Folder (${testFolderId}) OK. Nom: ${folder.getName()}\n`; const tempFileName = `_test_autorisation_${new Date().getTime()}.txt`; const tempFile = folder.createFile(tempFileName, 'Test autorisation Drive'); const tempFileId = tempFile.getId(); testResult += `Création fichier test (${tempFileName}) OK.\n`; DriveApp.getFileById(tempFileId).setTrashed(true); testResult += `Suppression fichier test OK.\n`; Logger.log("Test autorisations OK :\n" + testResult); SpreadsheetApp.getUi().alert("Test Autorisations", "OK:\n" + testResult, SpreadsheetApp.getUi().ButtonSet.OK); } catch (e) { Logger.log(`ERREUR Test autorisations : ${e}\n${e.stack}`); SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(`<pre>${e.stack}</pre>`), `Erreur Test: ${e.message}`); throw new Error(`ERREUR test autorisations: ${e.message}`); } }
// ==========================================================================
