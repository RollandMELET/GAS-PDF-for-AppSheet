/**
 * Suite de tests automatisés pour les fonctions de génération PDF
 * Tests.js - À ajouter au projet Google Apps Script
 */

// ==========================================================================
// Configuration des tests
// ==========================================================================
const TEST_CONFIG = {
  // Configuration pour les tests d'API AppSheet
  apiTest: {
    appId: "VOTRE_APP_ID_APPSHEET", // À remplacer
    accessKey: "VOTRE_CLE_ACCES_API", // À remplacer
    tableName: "NomDeVotreTable", // À remplacer
    uniqueId: "ID_TEST", // ID d'un enregistrement existant pour le test
    uniqueIdColumnName: "ID", // Nom de la colonne ID dans AppSheet
  },
  
  // Configuration pour les tests Google Sheet
  sheetTest: {
    spreadsheetId: "VOTRE_SPREADSHEET_ID", // À remplacer
    sheetName: "NomDeLaFeuille", // À remplacer
    uniqueIdColumnName: "ID", // Nom de la colonne ID dans Sheet
    uniqueId: "ID_TEST", // ID d'un enregistrement existant pour le test
    pdfLinkColumnName: "LienPDF", // Optionnel, colonne pour stocker l'URL du PDF
  },
  
  // Configuration commune
  common: {
    templateDocId: "VOTRE_TEMPLATE_DOC_ID", // À remplacer
    destinationFolderId: "VOTRE_DOSSIER_TEST_ID", // À remplacer
    pdfFilenameTemplate: "Test-{{ID}}.pdf",
    cleanupAfterTests: true, // Supprimer les fichiers générés par les tests
  }
};

// ==========================================================================
// Fonctions utilitaires pour les tests
// ==========================================================================

/**
 * Enregistre les résultats de test dans un fichier texte.
 * @param {string} testName Nom du test exécuté
 * @param {string} result Résultat (Succès/Échec)
 * @param {string} message Message détaillé
 * @param {Error} error Erreur éventuelle
 */
function logTestResult(testName, result, message, error = null) {
  const timestamp = new Date().toISOString();
  let logMessage = `[${timestamp}] [${result}] ${testName}: ${message}`;
  
  if (error) {
    logMessage += `\nERREUR: ${error.message}`;
    if (error.stack) {
      logMessage += `\nStack: ${error.stack}`;
    }
  }
  
  console.log(logMessage);
  
  // Optionnel: Écrire dans un fichier log dans Drive
  try {
    const folder = DriveApp.getFolderById(TEST_CONFIG.common.destinationFolderId);
    const logFile = folder.createFile(`TestLog_${new Date().getTime()}.txt`, logMessage);
    console.log(`Log enregistré: ${logFile.getUrl()}`);
  } catch (e) {
    console.error("Impossible d'enregistrer le log:", e);
  }
}

/**
 * Vérifie l'existence d'un fichier dans un dossier.
 * @param {string} folderId ID du dossier à vérifier
 * @param {string} partialFileName Partie du nom du fichier à rechercher
 * @returns {GoogleAppsScript.Drive.File|null} Le fichier trouvé ou null
 */
function findFileInFolder(folderId, partialFileName) {
  const folder = DriveApp.getFolderById(folderId);
  const files = folder.getFiles();
  
  while (files.hasNext()) {
    const file = files.next();
    if (file.getName().indexOf(partialFileName) !== -1) {
      return file;
    }
  }
  
  return null;
}

/**
 * Nettoie les fichiers de test.
 * @param {string[]} fileIds Liste des IDs de fichiers à supprimer
 */
function cleanupTestFiles(fileIds) {
  if (!TEST_CONFIG.common.cleanupAfterTests) return;
  
  for (const fileId of fileIds) {
    try {
      DriveApp.getFileById(fileId).setTrashed(true);
      console.log(`Fichier nettoyé: ${fileId}`);
    } catch (e) {
      console.error(`Erreur nettoyage ${fileId}:`, e);
    }
  }
}

// ==========================================================================
// Tests de la fonction API AppSheet
// ==========================================================================

/**
 * Teste la fonction fetchDataFromAppSheetAPI.
 * @returns {boolean} Vrai si le test réussit, faux sinon
 */
function testFetchDataFromAppSheetAPI() {
  const testName = "Test fetchDataFromAppSheetAPI";
  try {
    // Appel de la fonction à tester
    const result = fetchDataFromAppSheetAPI(
      TEST_CONFIG.apiTest.appId,
      TEST_CONFIG.apiTest.accessKey,
      TEST_CONFIG.apiTest.tableName,
      TEST_CONFIG.apiTest.uniqueIdColumnName,
      TEST_CONFIG.apiTest.uniqueId
    );
    
    // Vérification du résultat
    if (!result) {
      throw new Error("Résultat vide");
    }
    
    if (!result.data) {
      throw new Error("Données manquantes dans le résultat");
    }
    
    if (String(result.recordId) !== String(TEST_CONFIG.apiTest.uniqueId)) {
      throw new Error(`ID incorrect: ${result.recordId} vs ${TEST_CONFIG.apiTest.uniqueId}`);
    }
    
    // Vérifier que les données contiennent au moins quelques champs
    const dataKeys = Object.keys(result.data);
    if (dataKeys.length === 0) {
      throw new Error("Aucune donnée récupérée");
    }
    
    logTestResult(testName, "SUCCÈS", `Données récupérées correctement. Champs: ${dataKeys.join(', ')}`);
    return true;
  } catch (error) {
    logTestResult(testName, "ÉCHEC", "Erreur lors du test", error);
    return false;
  }
}

/**
 * Teste la génération complète via API AppSheet.
 * @returns {boolean} Vrai si le test réussit, faux sinon
 */
function testGeneratePdfFromTemplateAPI_AppSheetUse() {
  const testName = "Test generatePdfFromTemplateAPI_AppSheetUse";
  const filesToCleanup = [];
  
  try {
    // Appel de la fonction à tester
    const pdfName = generatePdfFromTemplateAPI_AppSheetUse(
      TEST_CONFIG.apiTest.uniqueId,
      TEST_CONFIG.apiTest.appId,
      TEST_CONFIG.apiTest.accessKey,
      TEST_CONFIG.apiTest.tableName,
      TEST_CONFIG.common.templateDocId,
      TEST_CONFIG.common.destinationFolderId,
      TEST_CONFIG.apiTest.uniqueIdColumnName,
      TEST_CONFIG.common.pdfFilenameTemplate,
      "true" // Supprimer doc temporaire
    );
    
    // Vérification du résultat
    if (!pdfName) {
      throw new Error("Nom de fichier PDF non retourné");
    }
    
    // Vérifier que le PDF existe bien
    const pdfFile = findFileInFolder(TEST_CONFIG.common.destinationFolderId, pdfName);
    if (!pdfFile) {
      throw new Error(`Fichier PDF "${pdfName}" non trouvé dans le dossier de destination`);
    }
    
    filesToCleanup.push(pdfFile.getId());
    
    // Vérifier que le nom du fichier contient l'ID (remplacé correctement)
    if (pdfName.indexOf(TEST_CONFIG.apiTest.uniqueId) === -1) {
      throw new Error(`Le nom du fichier "${pdfName}" ne contient pas l'ID ${TEST_CONFIG.apiTest.uniqueId}`);
    }
    
    logTestResult(testName, "SUCCÈS", `PDF "${pdfName}" généré avec succès via API AppSheet. URL: ${pdfFile.getUrl()}`);
    return true;
  } catch (error) {
    logTestResult(testName, "ÉCHEC", "Erreur lors du test", error);
    return false;
  } finally {
    if (filesToCleanup.length > 0) {
      cleanupTestFiles(filesToCleanup);
    }
  }
}

// ==========================================================================
// Tests de la fonction Google Sheet originale (pour comparaison)
// ==========================================================================

/**
 * Teste la génération via Google Sheet directement (fonction originale).
 * @returns {boolean} Vrai si le test réussit, faux sinon
 */
function testGeneratePdfFromTemplate() {
  const testName = "Test generatePdfFromTemplate (original)";
  const filesToCleanup = [];
  
  try {
    // Appel de la fonction à tester
    const pdfName = generatePdfFromTemplate(
      TEST_CONFIG.sheetTest.uniqueId,
      TEST_CONFIG.sheetTest.spreadsheetId,
      TEST_CONFIG.common.templateDocId,
      TEST_CONFIG.common.destinationFolderId,
      TEST_CONFIG.sheetTest.sheetName,
      TEST_CONFIG.sheetTest.uniqueIdColumnName,
      TEST_CONFIG.sheetTest.pdfLinkColumnName,
      TEST_CONFIG.common.pdfFilenameTemplate,
      "true" // Supprimer doc temporaire
    );
    
    // Vérification du résultat
    if (!pdfName) {
      throw new Error("Nom de fichier PDF non retourné");
    }
    
    // Vérifier que le PDF existe bien
    const pdfFile = findFileInFolder(TEST_CONFIG.common.destinationFolderId, pdfName);
    if (!pdfFile) {
      throw new Error(`Fichier PDF "${pdfName}" non trouvé dans le dossier de destination`);
    }
    
    filesToCleanup.push(pdfFile.getId());
    
    // Vérifier que le nom du fichier contient l'ID (remplacé correctement)
    if (pdfName.indexOf(TEST_CONFIG.sheetTest.uniqueId) === -1) {
      throw new Error(`Le nom du fichier "${pdfName}" ne contient pas l'ID ${TEST_CONFIG.sheetTest.uniqueId}`);
    }
    
    logTestResult(testName, "SUCCÈS", `PDF "${pdfName}" généré avec succès via Sheet. URL: ${pdfFile.getUrl()}`);
    return true;
  } catch (error) {
    logTestResult(testName, "ÉCHEC", "Erreur lors du test", error);
    return false;
  } finally {
    if (filesToCleanup.length > 0) {
      cleanupTestFiles(filesToCleanup);
    }
  }
}

// ==========================================================================
// Fonction principale d'exécution des tests
// ==========================================================================

/**
 * Exécute tous les tests.
 * Fonction à déclencher manuellement.
 */
function runAllTests() {
  console.log("=== DÉBUT DES TESTS ===");
  console.log(`Date/heure: ${new Date().toLocaleString()}`);
  
  let results = {
    passed: 0,
    failed: 0,
    total: 0
  };
  
  // Test de la fonction fetchDataFromAppSheetAPI
  if (testFetchDataFromAppSheetAPI()) {
    results.passed++;
  } else {
    results.failed++;
  }
  results.total++;
  
  // Test de la fonction generatePdfFromTemplateAPI_AppSheetUse
  if (testGeneratePdfFromTemplateAPI_AppSheetUse()) {
    results.passed++;
  } else {
    results.failed++;
  }
  results.total++;
  
  // Test de la fonction originale (pour comparaison)
  if (testGeneratePdfFromTemplate()) {
    results.passed++;
  } else {
    results.failed++;
  }
  results.total++;
  
  // Résumé
  console.log(`=== FIN DES TESTS ===`);
  console.log(`Tests exécutés: ${results.total}`);
  console.log(`Tests réussis: ${results.passed}`);
  console.log(`Tests échoués: ${results.failed}`);
  console.log(`Taux de succès: ${(results.passed / results.total * 100).toFixed(2)}%`);
  
  return {
    timestamp: new Date().toISOString(),
    results: results
  };
}

/**
 * Compare les résultats des deux méthodes de génération PDF.
 * Génère un PDF avec chaque méthode et compare les fichiers produits.
 */
function compareGenerationMethods() {
  const testName = "Comparaison des méthodes";
  const filesToCleanup = [];
  
  try {
    // 1. Générer PDF avec la méthode API
    const apiPdfName = generatePdfFromTemplateAPI_AppSheetUse(
      TEST_CONFIG.apiTest.uniqueId,
      TEST_CONFIG.apiTest.appId,
      TEST_CONFIG.apiTest.accessKey,
      TEST_CONFIG.apiTest.tableName,
      TEST_CONFIG.common.templateDocId,
      TEST_CONFIG.common.destinationFolderId,
      TEST_CONFIG.apiTest.uniqueIdColumnName,
      TEST_CONFIG.common.pdfFilenameTemplate,
      "true"
    );
    
    const apiPdfFile = findFileInFolder(TEST_CONFIG.common.destinationFolderId, apiPdfName);
    if (!apiPdfFile) {
      throw new Error("PDF via API non généré");
    }
    filesToCleanup.push(apiPdfFile.getId());
    const apiPdfSize = apiPdfFile.getSize();
    
    // 2. Générer PDF avec la méthode Sheet
    const sheetPdfName = generatePdfFromTemplate(
      TEST_CONFIG.sheetTest.uniqueId,
      TEST_CONFIG.sheetTest.spreadsheetId,
      TEST_CONFIG.common.templateDocId,
      TEST_CONFIG.common.destinationFolderId,
      TEST_CONFIG.sheetTest.sheetName,
      TEST_CONFIG.sheetTest.uniqueIdColumnName,
      TEST_CONFIG.sheetTest.pdfLinkColumnName,
      TEST_CONFIG.common.pdfFilenameTemplate,
      "true"
    );
    
    const sheetPdfFile = findFileInFolder(TEST_CONFIG.common.destinationFolderId, sheetPdfName);
    if (!sheetPdfFile) {
      throw new Error("PDF via Sheet non généré");
    }
    filesToCleanup.push(sheetPdfFile.getId());
    const sheetPdfSize = sheetPdfFile.getSize();
    
    // 3. Comparer les résultats
    // Note: La comparaison parfaite bit à bit est difficile sans bibliothèque spécifique
    // Nous comparons donc des caractéristiques de base comme la taille
    const sizeDiff = Math.abs(apiPdfSize - sheetPdfSize);
    const sizeDiffPercent = (sizeDiff / sheetPdfSize) * 100;
    
    // Si la différence de taille est inférieure à 5%, considérons les fichiers comme similaires
    const areFilesEquivalent = sizeDiffPercent < 5;
    
    if (areFilesEquivalent) {
      logTestResult(testName, "SUCCÈS", `Les PDF sont équivalents (différence taille: ${sizeDiffPercent.toFixed(2)}%). API: ${apiPdfFile.getUrl()}, Sheet: ${sheetPdfFile.getUrl()}`);
      return true;
    } else {
      throw new Error(`Différence significative entre les PDF: ${sizeDiffPercent.toFixed(2)}%. API: ${apiPdfSize} octets, Sheet: ${sheetPdfSize} octets.`);
    }
  } catch (error) {
    logTestResult(testName, "ÉCHEC", "Erreur lors de la comparaison", error);
    return false;
  } finally {
    if (filesToCleanup.length > 0) {
      cleanupTestFiles(filesToCleanup);
    }
  }
}
