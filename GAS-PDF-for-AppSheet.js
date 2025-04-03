/**
 * @OnlyCurrentDoc // Peut être retiré si non pertinent pour standalone, mais ne gêne pas.
 * Activer le runtime V8 pour utiliser des fonctionnalités modernes.
 * Dans l'éditeur de script: Paramètres du projet > Cochez "Activer le moteur d'exécution Chrome V8"
 */

// ... (fonctions sanitizeFilename et getUniqueFilename restent identiques) ...
function sanitizeFilename(filename) {
    if (!filename) return '';
    return filename.replace(/[\\/:*?"<>|]/g, '_');
}

function getUniqueFilename(folder, desiredName) {
    if (!desiredName) {
        desiredName = `Document_Sans_Nom_${new Date().getTime()}.pdf`;
        Logger.log(`Attention: Nom de fichier souhaité vide, utilisation de: ${desiredName}`);
    }
    let finalName = desiredName;
    let counter = 1;
    let files = folder.getFilesByName(finalName);
    while (files.hasNext()) {
        let baseName = finalName;
        let extension = '';
        const dotIndex = finalName.lastIndexOf('.');
        if (dotIndex > 0) {
            baseName = finalName.substring(0, dotIndex);
            extension = finalName.substring(dotIndex);
        }
        const match = baseName.match(/^(.*)\s\((\d+)\)$/);
        if (match) {
            baseName = match[1];
        }
        finalName = `${baseName} (${counter})${extension}`;
        files = folder.getFilesByName(finalName);
        counter++;
        if (counter > 100) {
            Logger.log(`ERREUR: Impossible de trouver un nom unique après 100 tentatives pour ${desiredName}`);
            finalName = `${baseName}_${new Date().getTime()}${extension}`;
            break;
        }
    }
    return finalName;
}


/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 * Tous les paramètres de configuration sont passés en arguments.
 * L'ORDRE DES PARAMÈTRES EST CRUCIAL.
 *
 * @param {string} uniqueId La valeur de la clé unique de la ligne à traiter.
 * @param {string} spreadsheetId L'ID DU FICHIER Google Sheet contenant les données. <<< NOUVEAU PARAMÈTRE
 * @param {string} templateDocId L'ID du document Google Docs servant de modèle.
 * @param {string} destinationFolderId L'ID du dossier Google Drive où sauvegarder le PDF.
 * @param {string} sheetName Le nom exact de la feuille (onglet) Google Sheet contenant les données.
 * @param {string} uniqueIdColumnNameInSheet Le nom exact de la colonne clé unique dans la feuille Sheet.
 * @param {string} pdfLinkColumnName Le nom de la colonne où stocker l'URL du PDF (laisser vide '' si non utilisé).
 * @param {string} pdfFilenameTemplate Le modèle pour le nom du fichier PDF (ex: "BL-{{ID}}.pdf").
 * Tparam {string} deleteTempDocStr Indique s'il faut supprimer le Doc temporaire ('true' ou 'false').
 *
 * @return {string} Le NOM FINAL du fichier PDF généré ou un message d'erreur.
 */
function generatePdfFromTemplate(
    uniqueId,
    spreadsheetId, // <<< NOUVEAU PARAMÈTRE AJOUTÉ ICI
    templateDocId,
    destinationFolderId,
    sheetName,
    uniqueIdColumnNameInSheet,
    pdfLinkColumnName,
    pdfFilenameTemplate,
    deleteTempDocStr
) {

  // --- Validation des paramètres essentiels ---
  if (!uniqueId) throw new Error("Erreur Script: L'argument 'uniqueId' est manquant ou vide.");
  if (!spreadsheetId) throw new Error("Erreur Script: L'argument 'spreadsheetId' (ID du fichier Sheet) est manquant."); // <<< VALIDATION AJOUTÉE
  if (!templateDocId) throw new Error("Erreur Script: L'argument 'templateDocId' est manquant.");
  if (!destinationFolderId) throw new Error("Erreur Script: L'argument 'destinationFolderId' est manquant.");
  if (!sheetName) throw new Error("Erreur Script: L'argument 'sheetName' est manquant.");
  if (!uniqueIdColumnNameInSheet) throw new Error("Erreur Script: L'argument 'uniqueIdColumnNameInSheet' est manquant.");
  if (!pdfFilenameTemplate) throw new Error("Erreur Script: L'argument 'pdfFilenameTemplate' est manquant.");

  const deleteTempDoc = (deleteTempDocStr === 'true' || deleteTempDocStr === true);

  Logger.log(`Début génération PDF pour ID: ${uniqueId}`);
  Logger.log(`Paramètres reçus: spreadsheetId=${spreadsheetId}, templateId=${templateDocId}, folderId=${destinationFolderId}, sheet=${sheetName}, idCol=${uniqueIdColumnNameInSheet}, linkCol=${pdfLinkColumnName}, nameTemplate=${pdfFilenameTemplate}, deleteDoc=${deleteTempDoc}`); // <<< LOG MIS À JOUR

  try {
    // 1. Accès aux données Sheet via ID
    let ss;
    try {
        // <<< CHANGEMENT MAJEUR ICI : Utilisation de openById au lieu de getActiveSpreadsheet
        ss = SpreadsheetApp.openById(spreadsheetId);
    } catch (e) {
         throw new Error(`Impossible d'ouvrir le fichier Google Sheet avec l'ID "${spreadsheetId}". Vérifiez l'ID et les permissions du script pour accéder à ce fichier. Détails: ${e.message}`);
    }

    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Feuille (onglet) "${sheetName}" non trouvée dans le fichier Sheet ID: ${spreadsheetId}.`);

    // ... (Le reste du code reste identique : dataRange, allData, headers, recherche de ligne, placeholders, etc.) ...

    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    const headers = allData[0];

    const uniqueIdColIndex = headers.indexOf(uniqueIdColumnNameInSheet);
    if (uniqueIdColIndex === -1) throw new Error(`Colonne clé "${uniqueIdColumnNameInSheet}" non trouvée dans la feuille "${sheetName}".`);

    let pdfLinkColIndex = -1;
    if (pdfLinkColumnName) {
        pdfLinkColIndex = headers.indexOf(pdfLinkColumnName);
        if (pdfLinkColIndex === -1) {
            Logger.log(`Attention: Colonne lien PDF "${pdfLinkColumnName}" spécifiée mais non trouvée.`);
        }
    }

    let targetRowData = null;
    let targetRowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      const cellValue = allData[i][uniqueIdColIndex];
      if (cellValue !== null && cellValue !== undefined && cellValue.toString() === uniqueId.toString()) {
        targetRowData = allData[i];
        targetRowIndex = i;
        break;
      }
    }
    if (!targetRowData) throw new Error(`Aucune ligne trouvée avec l'ID unique "${uniqueId}" dans la colonne "${uniqueIdColumnNameInSheet}".`);

    const placeholders = {};
    headers.forEach((header, index) => {
      if (header) {
          placeholders[`{{${header}}}`] = targetRowData[index] !== null && targetRowData[index] !== undefined ? targetRowData[index].toString() : '';
      }
    });

    let templateFile, destinationFolder;
    try { templateFile = DriveApp.getFileById(templateDocId); } catch (e) { throw new Error(`Impossible d'accéder au fichier modèle ID "${templateDocId}". Détails: ${e.message}`); }
    try { destinationFolder = DriveApp.getFolderById(destinationFolderId); } catch (e) { throw new Error(`Impossible d'accéder au dossier destination ID "${destinationFolderId}". Détails: ${e.message}`); }

    let desiredTempDocName = `Temp_${pdfFilenameTemplate.replace(/\.pdf$/i, '')}`;
    for (const placeholderKey in placeholders) { const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); desiredTempDocName = desiredTempDocName.replace(regex, placeholders[placeholderKey]); }
    desiredTempDocName = sanitizeFilename(desiredTempDocName) || `Temp_Doc_${uniqueId}`;
    const uniqueTempDocName = getUniqueFilename(destinationFolder, desiredTempDocName);

    const copiedFile = templateFile.makeCopy(uniqueTempDocName, destinationFolder);
    const copiedDoc = DocumentApp.openById(copiedFile.getId());
    const copiedFileId = copiedFile.getId();

    const body = copiedDoc.getBody(); const header = copiedDoc.getHeader(); const footer = copiedDoc.getFooter();
    for (const placeholderKey in placeholders) {
        const value = placeholders[placeholderKey];
        const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        try { body.replaceText(regex, value); if (header) header.replaceText(regex, value); if (footer) footer.replaceText(regex, value); } catch (replaceError) { Logger.log(`Avertissement: Erreur remplacement placeholder ${placeholderKey}. Détails: ${replaceError.message}`); }
    }
    copiedDoc.saveAndClose();

    let desiredPdfName = pdfFilenameTemplate;
    for (const placeholderKey in placeholders) { const value = placeholders[placeholderKey]; const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g'); desiredPdfName = desiredPdfName.replace(regex, value); }
    desiredPdfName = sanitizeFilename(desiredPdfName);
     if (!desiredPdfName || !desiredPdfName.toLowerCase().endsWith('.pdf')) { Logger.log(`Attention: Nom de fichier généré "${desiredPdfName}" invalide. Utilisation nom défaut.`); desiredPdfName = `Document_${uniqueId}.pdf`; }
    const uniquePdfName = getUniqueFilename(destinationFolder, desiredPdfName);

    const pdfBlob = copiedFile.getAs(MimeType.PDF);
    const pdfFile = destinationFolder.createFile(pdfBlob).setName(uniquePdfName);
    const pdfUrl = pdfFile.getUrl(); const pdfId = pdfFile.getId();
    Logger.log(`PDF généré: ${uniquePdfName}, URL: ${pdfUrl}`);

    if (pdfLinkColumnName && pdfLinkColIndex !== -1 && targetRowIndex !== -1) {
      sheet.getRange(targetRowIndex + 1, pdfLinkColIndex + 1).setValue(pdfUrl);
      SpreadsheetApp.flush();
      Logger.log(`Lien PDF (URL) ajouté ligne ${targetRowIndex + 1}, colonne ${pdfLinkColumnName}`);
    }

    if (deleteTempDoc) {
      try { DriveApp.getFileById(copiedFileId).setTrashed(true); Logger.log(`Doc temporaire ${copiedFile.getName()} supprimé.`); } catch (deleteError) { Logger.log(`Avertissement: Impossible supprimer doc temporaire ${copiedFile.getName()}. Détails: ${deleteError.message}`); }
    }

    Logger.log(`Retournant nom fichier: ${uniquePdfName}`);
    return uniquePdfName;

  } catch (error) {
    Logger.log(`ERREUR génération PDF pour ID ${uniqueId}: ${error.message} \n Stack: ${error.stack}`);
    throw new Error(`Erreur script: ${error.message}`);
  }
}
