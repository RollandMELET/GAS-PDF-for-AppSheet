/**
 * @OnlyCurrentDoc
 * Activer le runtime V8 pour utiliser des fonctionnalités modernes.
 * Dans l'éditeur de script: Paramètres du projet > Cochez "Activer le moteur d'exécution Chrome V8"
 */

/**
 * Nettoie un nom de fichier des caractères interdits.
 * @param {string} filename Le nom de fichier potentiel.
 * @return {string} Le nom de fichier nettoyé.
 */
function sanitizeFilename(filename) {
  if (!filename) return '';
  return filename.replace(/[\\/:*?"<>|]/g, '_');
}

/**
 * Trouve un nom de fichier unique dans un dossier donné en ajoutant (1), (2), etc. si nécessaire.
 * @param {GoogleAppsScript.Drive.Folder} folder Le dossier où vérifier l'unicité.
 * @param {string} desiredName Le nom de fichier initial souhaité (avec extension).
 * @return {string} Un nom de fichier garanti unique dans ce dossier.
 */
function getUniqueFilename(folder, desiredName) {
  if (!desiredName) {
      // Générer un nom par défaut si desiredName est vide/null
      desiredName = `Document_Sans_Nom_${new Date().getTime()}.pdf`;
      Logger.log(`Attention: Nom de fichier souhaité vide, utilisation de: ${desiredName}`);
  }

  let finalName = desiredName;
  let counter = 1;
  let files = folder.getFilesByName(finalName); // Utiliser let pour pouvoir réassigner

  // Boucle while pour trouver un nom unique
  while (files.hasNext()) {
    let baseName = finalName;
    let extension = '';
    const dotIndex = finalName.lastIndexOf('.');
    if (dotIndex > 0) {
      baseName = finalName.substring(0, dotIndex);
      extension = finalName.substring(dotIndex);
    }

    // Gérer le cas où le nom contient déjà " (n)" pour éviter " (1) (1)"
    const match = baseName.match(/^(.*)\s\((\d+)\)$/);
    if (match) {
      baseName = match[1]; // Utiliser la base avant le suffixe (n)
    }

    // Construire le nouveau nom potentiel
    finalName = `${baseName} (${counter})${extension}`;
    files = folder.getFilesByName(finalName); // Ré-évaluer avec le nouveau nom potentiel
    counter++;

    // Sécurité pour éviter boucle infinie (très improbable mais sait-on jamais)
    if (counter > 100) {
        Logger.log(`ERREUR: Impossible de trouver un nom unique après 100 tentatives pour ${desiredName}`);
        finalName = `${baseName}_${new Date().getTime()}${extension}`; // Nom unique basé sur le temps
        break;
    }
  }
  return finalName;
}


/**
 * Fonction principale appelée par AppSheet pour générer le PDF.
 * Tous les paramètres de configuration sont passés en arguments.
 * L'ORDRE DES PARAMÈTRES EST CRUCIAL LORS DE L'APPEL DEPUIS APPSHEET.
 *
 * @param {string} uniqueId La valeur de la clé unique de la ligne à traiter.
 * @param {string} templateDocId L'ID du document Google Docs servant de modèle.
 * @param {string} destinationFolderId L'ID du dossier Google Drive où sauvegarder le PDF.
 * @param {string} sheetName Le nom exact de la feuille Google Sheet contenant les données.
 * @param {string} uniqueIdColumnNameInSheet Le nom exact de la colonne clé unique dans la feuille Sheet.
 * @param {string} pdfLinkColumnName Le nom de la colonne où stocker l'URL du PDF (laisser vide '' si non utilisé).
 * @param {string} pdfFilenameTemplate Le modèle pour le nom du fichier PDF (ex: "BL-{{ID}}.pdf").
 * @param {string} deleteTempDocStr Indique s'il faut supprimer le Doc temporaire ('true' ou 'false').
 *
 * @return {string} Le NOM FINAL du fichier PDF généré ou un message d'erreur.
 */
function generatePdfFromTemplate(
    uniqueId,
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
  if (!templateDocId) throw new Error("Erreur Script: L'argument 'templateDocId' est manquant.");
  if (!destinationFolderId) throw new Error("Erreur Script: L'argument 'destinationFolderId' est manquant.");
  if (!sheetName) throw new Error("Erreur Script: L'argument 'sheetName' est manquant.");
  if (!uniqueIdColumnNameInSheet) throw new Error("Erreur Script: L'argument 'uniqueIdColumnNameInSheet' est manquant.");
  if (!pdfFilenameTemplate) throw new Error("Erreur Script: L'argument 'pdfFilenameTemplate' est manquant.");

  // Convertir la chaîne 'true'/'false' en booléen
  const deleteTempDoc = (deleteTempDocStr === 'true' || deleteTempDocStr === true); // Gère string et boolean au cas où

  Logger.log(`Début génération PDF pour ID: ${uniqueId}`);
  Logger.log(`Paramètres reçus: templateId=${templateDocId}, folderId=${destinationFolderId}, sheet=${sheetName}, idCol=${uniqueIdColumnNameInSheet}, linkCol=${pdfLinkColumnName}, nameTemplate=${pdfFilenameTemplate}, deleteDoc=${deleteTempDoc}`);

  try {
    // 1. Accès aux données Sheet
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(sheetName);
    if (!sheet) throw new Error(`Feuille "${sheetName}" non trouvée.`);

    const dataRange = sheet.getDataRange();
    const allData = dataRange.getValues();
    const headers = allData[0];

    const uniqueIdColIndex = headers.indexOf(uniqueIdColumnNameInSheet);
    if (uniqueIdColIndex === -1) throw new Error(`Colonne clé "${uniqueIdColumnNameInSheet}" non trouvée dans la feuille "${sheetName}".`);

    let pdfLinkColIndex = -1;
    if (pdfLinkColumnName) { // Vérifier si une colonne est spécifiée
        pdfLinkColIndex = headers.indexOf(pdfLinkColumnName);
        if (pdfLinkColIndex === -1) {
            Logger.log(`Attention: Colonne lien PDF "${pdfLinkColumnName}" spécifiée mais non trouvée.`);
            // On continue, mais on ne pourra pas écrire le lien
        }
    }

    // 2. Trouver la ligne
    let targetRowData = null;
    let targetRowIndex = -1;
    for (let i = 1; i < allData.length; i++) {
      // Vérifier que la cellule n'est pas vide avant toString()
      const cellValue = allData[i][uniqueIdColIndex];
      if (cellValue !== null && cellValue !== undefined && cellValue.toString() === uniqueId.toString()) {
        targetRowData = allData[i];
        targetRowIndex = i;
        break;
      }
    }
    if (!targetRowData) throw new Error(`Aucune ligne trouvée avec l'ID unique "${uniqueId}" dans la colonne "${uniqueIdColumnNameInSheet}".`);

    // 3. Préparer les placeholders
    const placeholders = {};
    headers.forEach((header, index) => {
      // Gérer le cas où le nom de header est vide (possible dans Sheets)
      if (header) {
          placeholders[`{{${header}}}`] = targetRowData[index] !== null && targetRowData[index] !== undefined ? targetRowData[index].toString() : '';
      }
    });

    // 4. Préparer le fichier temporaire
    let templateFile, destinationFolder;
    try {
        templateFile = DriveApp.getFileById(templateDocId);
    } catch (e) {
        throw new Error(`Impossible d'accéder au fichier modèle avec l'ID "${templateDocId}". Vérifiez l'ID et les permissions. Détails: ${e.message}`);
    }
     try {
        destinationFolder = DriveApp.getFolderById(destinationFolderId);
    } catch (e) {
        throw new Error(`Impossible d'accéder au dossier de destination avec l'ID "${destinationFolderId}". Vérifiez l'ID et les permissions. Détails: ${e.message}`);
    }


    // Générer le nom de base souhaité pour le Doc temporaire
    let desiredTempDocName = `Temp_${pdfFilenameTemplate.replace(/\.pdf$/i, '')}`; // Enlever .pdf
    for (const placeholderKey in placeholders) {
        const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        desiredTempDocName = desiredTempDocName.replace(regex, placeholders[placeholderKey]);
    }
    desiredTempDocName = sanitizeFilename(desiredTempDocName) || `Temp_Doc_${uniqueId}`; // Nom par défaut si vide

    // Obtenir un nom unique pour le Doc temporaire
    const uniqueTempDocName = getUniqueFilename(destinationFolder, desiredTempDocName);

    // Copier le modèle avec le nom unique
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
        try { // Ajouter un try/catch autour des remplacements individuels peut aider au debug
           body.replaceText(regex, value);
           if (header) header.replaceText(regex, value);
           if (footer) footer.replaceText(regex, value);
        } catch (replaceError) {
            Logger.log(`Avertissement: Erreur lors du remplacement du placeholder ${placeholderKey} avec la valeur "${value}". Détails: ${replaceError.message}`);
            // Continuer avec les autres remplacements
        }
    }
    copiedDoc.saveAndClose();

    // 6. Préparer le fichier PDF final
    let desiredPdfName = pdfFilenameTemplate;
    for (const placeholderKey in placeholders) {
        const value = placeholders[placeholderKey];
        const regex = new RegExp(placeholderKey.replace(/[-\/\\^$*+?.()|[\]{}]/g, '\\$&'), 'g');
        desiredPdfName = desiredPdfName.replace(regex, value);
    }
    desiredPdfName = sanitizeFilename(desiredPdfName);
     if (!desiredPdfName || !desiredPdfName.toLowerCase().endsWith('.pdf')) {
         Logger.log(`Attention: Le nom de fichier généré "${desiredPdfName}" est invalide ou manque l'extension .pdf. Utilisation d'un nom par défaut.`);
         desiredPdfName = `Document_${uniqueId}.pdf`;
     }

    // Obtenir un nom unique pour le PDF
    const uniquePdfName = getUniqueFilename(destinationFolder, desiredPdfName);

    // 7. Créer le PDF avec le nom unique
    const pdfBlob = copiedFile.getAs(MimeType.PDF);
    const pdfFile = destinationFolder.createFile(pdfBlob).setName(uniquePdfName);
    const pdfUrl = pdfFile.getUrl();
    const pdfId = pdfFile.getId();

    Logger.log(`PDF généré avec succès : ${uniquePdfName}, URL: ${pdfUrl}`);

    // 8. Mettre à jour la feuille Sheet (Optionnel)
    if (pdfLinkColumnName && pdfLinkColIndex !== -1 && targetRowIndex !== -1) {
      // +1 car getRange est basé sur 1, targetRowIndex est basé sur 0
      // +1 car pdfLinkColIndex est basé sur 0, getRange basé sur 1
      sheet.getRange(targetRowIndex + 1, pdfLinkColIndex + 1).setValue(pdfUrl);
      SpreadsheetApp.flush(); // Forcer l'écriture
      Logger.log(`Lien PDF (URL) ajouté à la ligne ${targetRowIndex + 1}, colonne ${pdfLinkColumnName}`);
    }

    // 9. Supprimer le Doc temporaire (Optionnel)
    if (deleteTempDoc) {
      try {
          DriveApp.getFileById(copiedFileId).setTrashed(true);
          Logger.log(`Document Google Doc intermédiaire ${copiedFile.getName()} supprimé.`);
      } catch (deleteError) {
           Logger.log(`Avertissement: Impossible de supprimer le document temporaire ${copiedFile.getName()} (ID: ${copiedFileId}). Détails: ${deleteError.message}`);
      }
    }

    // 10. Retourner le NOM FINAL du fichier PDF généré
    Logger.log(`Retournant le nom de fichier à AppSheet: ${uniquePdfName}`);
    return uniquePdfName;

  } catch (error) {
    Logger.log(`ERREUR lors de la génération du PDF pour ID ${uniqueId}: ${error.message} \n Stack: ${error.stack}`);
    // Renvoyer l'erreur pour qu'AppSheet puisse la capturer
    throw new Error(`Erreur script: ${error.message}`);
  }
}
