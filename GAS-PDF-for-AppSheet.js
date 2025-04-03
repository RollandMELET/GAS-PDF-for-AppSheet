/**
 * @OnlyCurrentDoc
 * Activer le runtime V8 pour utiliser des fonctionnalités modernes comme String.prototype.replaceAll
 * Dans l'éditeur de script: Paramètres du projet > Cochez "Activer le moteur d'exécution Chrome V8"
 */

// --- CONFIGURATION ---
const CONFIG = {
  TEMPLATE_DOC_ID: 'ID_DE_VOTRE_MODELE_GOOGLE_DOC', // <<< REMPLACEZ CECI
  DESTINATION_FOLDER_ID: 'ID_DE_VOTRE_DOSSIER_DE_DESTINATION', // <<< REMPLACEZ CECI
  SHEET_NAME: 'NomDeVotreFeuille', // <<< REMPLACEZ CECI
  UNIQUE_ID_COLUMN_NAME: 'NomColonneCleUnique', // <<< REMPLACEZ CECI
  PDF_LINK_COLUMN_NAME: 'LienPDF', // <<< REMPLACEZ CECI (ou mettre '')
  PDF_FILENAME_TEMPLATE: 'BL-{{NomColonneCleUnique}}.pdf', // <<< ADAPTEZ CE MODELE
  DELETE_TEMP_DOC: false // Mettre à true pour supprimer le Doc temporaire
};
// --- FIN DE LA CONFIGURATION ---

/**
 * Nettoie un nom de fichier des caractères interdits.
 * @param {string} filename Le nom de fichier potentiel.
 * @return {string} Le nom de fichier nettoyé.
 */
function sanitizeFilename(filename) {
  return filename.replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * Trouve un nom de fichier unique dans un dossier donné en ajoutant (1), (2), etc. si nécessaire.
 * @param {GoogleAppsScript.Drive.Folder} folder Le dossier où vérifier l'unicité.
 * @param {string} desiredName Le nom de fichier initial souhaité (avec extension).
 * @return {string} Un nom de fichier garanti unique dans ce dossier.
 */
function getUniqueFilename(folder, desiredName) {
  let finalName = desiredName;
  let counter = 1;
  let files = folder.getFilesByName(finalName); // Utiliser let pour pouvoir réassigner

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
    files = folder.getFilesByName(finalName); // Ré-évaluer avec le nouveau nom potentiel
    counter++;
  }
  return finalName;
}


/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 * @param {string} uniqueId La valeur de la clé unique de la ligne à traiter.
 * @return {string} Le NOM FINAL du fichier PDF généré ou un message d'erreur.
 */
function generatePdfFromTemplate(uniqueId) {
  // ... (début de la fonction: validation config, accès sheet, trouver ligne, placeholders... reste identique) ...
  // Valider la configuration de base
  if (!CONFIG.TEMPLATE_DOC_ID || !CONFIG.DESTINATION_FOLDER_ID || !CONFIG.SHEET_NAME || !CONFIG.UNIQUE_ID_COLUMN_NAME || !CONFIG.PDF_FILENAME_TEMPLATE) {
     Logger.log("Erreur de configuration : Une ou plusieurs constantes CONFIG sont manquantes.");
     throw new Error("Erreur de configuration du script. Vérifiez les constantes.");
  }

  try {
    // 1. Accès aux données Sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    if (!sheet) throw new Error(`Feuille "${CONFIG.SHEET_NAME}" non trouvée.`);

    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    const headers = allData[0];

    const uniqueIdColIndex = headers.indexOf(CONFIG.UNIQUE_ID_COLUMN_NAME);
    if (uniqueIdColIndex === -1) throw new Error(`Colonne clé "${CONFIG.UNIQUE_ID_COLUMN_NAME}" non trouvée.`);
    let pdfLinkColIndex = CONFIG.PDF_LINK_COLUMN_NAME ? headers.indexOf(CONFIG.PDF_LINK_COLUMN_NAME) : -1;
     if (CONFIG.PDF_LINK_COLUMN_NAME && pdfLinkColIndex === -1) {
        Logger.log(`Attention: Colonne lien PDF "${CONFIG.PDF_LINK_COLUMN_NAME}" non trouvée.`);
     }


    // 2. Trouver la ligne
    let targetRowData = null;
    let targetRowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      if (allData[i][uniqueIdColIndex] !== null && allData[i][uniqueIdColIndex] !== undefined &&
          allData[i][uniqueIdColIndex].toString() === uniqueId.toString()) {
        targetRowData = allData[i];
        targetRowIndex = i;
        break;
      }
    }
    if (!targetRowData) throw new Error(`Aucune ligne trouvée avec l'ID unique "${uniqueId}" dans la colonne "${CONFIG.UNIQUE_ID_COLUMN_NAME}".`);

    // 3. Préparer les placeholders
    const placeholders = {};
    headers.forEach((header, index) => {
      placeholders[`{{${header}}}`] = targetRowData[index] !== null && targetRowData[index] !== undefined ? targetRowData[index].toString() : '';
    });

    // 4. Préparer le fichier temporaire
    const templateFile = DriveApp.getFileById(CONFIG.TEMPLATE_DOC_ID);
    const destinationFolder = DriveApp.getFolderById(CONFIG.DESTINATION_FOLDER_ID);

    let desiredTempDocName = `Temp_${CONFIG.PDF_FILENAME_TEMPLATE.replace(/\.pdf$/i, '')}`;
    for (const placeholderKey in placeholders) {
        const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        desiredTempDocName = desiredTempDocName.replace(regex, placeholders[placeholderKey]);
    }
    desiredTempDocName = sanitizeFilename(desiredTempDocName);
    const uniqueTempDocName = getUniqueFilename(destinationFolder, desiredTempDocName);

    const copiedFile = templateFile.makeCopy(uniqueTempDocName, destinationFolder);
    const copiedDoc = DocumentApp.openById(copiedFile.getId());
    const copiedFileId = copiedFile.getId();

    // 5. Remplacer les placeholders dans la copie
    const body = copiedDoc.getBody();
    const header = copiedDoc.getHeader();
    const footer = copiedDoc.getFooter();
    for (const placeholderKey in placeholders) {
        const value = placeholders[placeholderKey];
        const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        body.replaceText(regex, value);
        if (header) header.replaceText(regex, value);
        if (footer) footer.replaceText(regex, value);
    }
    copiedDoc.saveAndClose();

    // 6. Préparer le fichier PDF final
    let desiredPdfName = CONFIG.PDF_FILENAME_TEMPLATE;
    for (const placeholderKey in placeholders) {
        const value = placeholders[placeholderKey];
        const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        desiredPdfName = desiredPdfName.replace(regex, value);
    }
    desiredPdfName = sanitizeFilename(desiredPdfName);
    const uniquePdfName = getUniqueFilename(destinationFolder, desiredPdfName); // <<< C'est cette variable que nous voulons retourner

    // 7. Créer le PDF avec le nom unique
    const pdfBlob = copiedFile.getAs(MimeType.PDF);
    const pdfFile = destinationFolder.createFile(pdfBlob).setName(uniquePdfName);
    const pdfUrl = pdfFile.getUrl(); // Nous avons toujours l'URL ici si besoin
    const pdfId = pdfFile.getId();

    Logger.log(`PDF généré avec succès : ${uniquePdfName}, URL: ${pdfUrl}`); // Log toujours l'URL aussi pour le debug

    // 8. Mettre à jour la feuille Sheet (Optionnel) - Met toujours à jour avec l'URL
    if (CONFIG.PDF_LINK_COLUMN_NAME && pdfLinkColIndex !== -1 && targetRowIndex !== -1) {
      sheet.getRange(targetRowIndex + 1, pdfLinkColIndex + 1).setValue(pdfUrl); // On continue de stocker l'URL dans la feuille
      SpreadsheetApp.flush();
      Logger.log(`Lien PDF (URL) ajouté à la ligne ${targetRowIndex + 1}, colonne ${CONFIG.PDF_LINK_COLUMN_NAME}`);
    }

    // 9. Supprimer le Doc temporaire (Optionnel)
    if (CONFIG.DELETE_TEMP_DOC) {
      DriveApp.getFileById(copiedFileId).setTrashed(true);
      Logger.log(`Document Google Doc intermédiaire ${copiedFile.getName()} supprimé.`);
    }

    // 10. Retourner le NOM FINAL du fichier PDF généré
    Logger.log(`Retournant le nom de fichier à AppSheet: ${uniquePdfName}`);
    return uniquePdfName; // <<< MODIFICATION ICI: Retourne le nom du fichier

  } catch (error) {
    Logger.log(`Erreur lors de la génération du PDF pour ID ${uniqueId}: ${error.message} \n Stack: ${error.stack}`);
    throw new Error(`Erreur script : ${error.message}`);
  }
}