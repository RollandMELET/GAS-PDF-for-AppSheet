// Fonction de test pour l'API AppSheet

/**
 * Fonction de test pour la fonction generatePdfFromTemplateAPI_AppSheetUse
 * Elle utilise les paramètres précédemment fournis.
 */
function testAppSheetAPIFunctionWithProvidedParams() {
  const uniqueId = "ba154bac"; // ID unique valide fourni
  const appsheetAppId = "63a906af-52b2-4d40-884f-18e95ca79c45";
  const appsheetAccessKey = "V2-nvnfP-sBEaK-EGTnO-zkXL5-pe6R9-dRm6i-F3djI-spKMq";
  const tableName = "PAC";
  const templateDocId = "1t6ex7K3qTtfwOA2tdsA4Q8jnb1xW8k1RX_39Xu3j1w0"; // ID du modèle DocG
  const destinationFolderId = "1tXLU6t-i-h3FFqws6QjIZciyMMxCvHjv"; // ID du dossier destination
  const uniqueIdColumnName = "PAC_ID"; // Nom de la colonne ID unique
  const pdfFilenameTemplate = "Test-API-AppSheet-{{PAC_N_OF}}-{{PAC_sGTIN}}.pdf"; // Modèle pour le nom du PDF
  const deleteTempDocStr = "false"; // Ne pas supprimer le doc temporaire pour vérification
  
  Logger.log("Début du test de la fonction generatePdfFromTemplateAPI_AppSheetUse avec les paramètres fournis");
  
  try {
    const result = generatePdfFromTemplateAPI_AppSheetUse(
      uniqueId,
      appsheetAppId,
      appsheetAccessKey,
      tableName,
      templateDocId,
      destinationFolderId,
      uniqueIdColumnName,
      pdfFilenameTemplate,
      deleteTempDocStr
    );
    
    Logger.log("TEST RÉUSSI! PDF généré avec succès. Nom du fichier : " + result);
    return result;
  } catch (error) {
    Logger.log("TEST ÉCHOUÉ! Erreur : " + error.message);
    if (error.stack) {
      Logger.log("Stack trace : " + error.stack);
    }
    throw error;
  }
}

// Fonction de test pour tester uniquement la récupération des données via l'API AppSheet
function testFetchDataOnly() {
  const appsheetAppId = "63a906af-52b2-4d40-884f-18e95ca79c45";
  const appsheetAccessKey = "V2-nvnfP-sBEaK-EGTnO-zkXL5-pe6R9-dRm6i-F3djI-spKMq";
  const tableName = "PAC";
  const uniqueIdColumnName = "PAC_ID";
  const uniqueId = "ba154bac";
  
  Logger.log("Test de récupération des données uniquement via l'API AppSheet");
  
  try {
    const recordData = fetchDataFromAppSheetAPI(
      appsheetAppId,
      appsheetAccessKey,
      tableName,
      uniqueIdColumnName,
      uniqueId
    );
    
    Logger.log("TEST RÉUSSI! Données récupérées : " + JSON.stringify(recordData, null, 2));
    return recordData;
  } catch (error) {
    Logger.log("TEST ÉCHOUÉ! Erreur récupération données : " + error.message);
    if (error.stack) {
      Logger.log("Stack trace : " + error.stack);
    }
    throw error;
  }
}
