// Code.gs - Côté serveur de l'application

// Fonction doGet requise pour le déploiement web
function doGet(e) {
  try {
    Logger.log("Début de l'exécution de doGet()");
    Logger.log("Paramètres de la requête : " + JSON.stringify(e));
    
    // Vérifier si l'utilisateur est autorisé
    const autorisation = verifierAutorisation();
    
    // Si l'utilisateur n'est pas autorisé, afficher une page d'erreur
    if (!autorisation.autorise) {
      Logger.log("Accès refusé pour l'utilisateur: " + autorisation.email);
      
      // Créer une page d'erreur personnalisée
      var htmlOutput = HtmlService.createHtmlOutput(
        `<html>
        <head>
          <base target="_top">
          <meta charset="utf-8">
          <meta name="viewport" content="width=device-width, initial-scale=1">
          <title>Accès refusé</title>
          <link href="https://cdn.jsdelivr.net/npm/bootstrap@5.3.0/dist/css/bootstrap.min.css" rel="stylesheet">
          <style>
            body {
              font-family: 'Inter', -apple-system, BlinkMacSystemFont, 'Segoe UI', Roboto, sans-serif;
              background-color: #f5f7ff;
              height: 100vh;
              display: flex;
              justify-content: center;
              align-items: center;
              padding: 20px;
            }
            
            .error-card {
              max-width: 500px;
              border-radius: 12px;
              overflow: hidden;
              box-shadow: 0 4px 20px rgba(0, 0, 0, 0.1);
            }
            
            .error-header {
              background-color: #f72585;
              color: white;
              padding: 20px;
            }
            
            .error-body {
              background-color: white;
              padding: 30px;
            }
            
            .error-icon {
              font-size: 48px;
              margin-bottom: 20px;
            }
          </style>
        </head>
        <body>
          <div class="error-card">
            <div class="error-header">
              <h3><i class="bi bi-shield-lock"></i> Accès refusé</h3>
            </div>
            <div class="error-body">
              <div class="text-center error-icon text-danger">
                <i class="bi bi-lock-fill"></i>
              </div>
              <div class="alert alert-danger">
                <strong>${autorisation.message}</strong>
              </div>
              <p class="mb-3">
                Votre adresse email: <strong>${autorisation.email}</strong>
              </p>
              <p class="text-muted">
                Si vous pensez qu'il s'agit d'une erreur, veuillez contacter le gestionnaire de l'application.
              </p>
              <div class="d-grid gap-2">
                <a href="javascript:window.close();" class="btn btn-outline-secondary">Fermer</a>
              </div>
            </div>
          </div>
          
          <!-- Bootstrap Icons -->
          <link rel="stylesheet" href="https://cdn.jsdelivr.net/npm/bootstrap-icons@1.8.0/font/bootstrap-icons.css">
        </body>
        </html>`
      );
      
      htmlOutput.setTitle('Accès refusé');
      return htmlOutput;
    }
    
    // Si l'utilisateur est autorisé, continuer avec le chargement normal
    Logger.log("Accès autorisé pour l'utilisateur: " + autorisation.email);
    
    // Déterminer quelle page servir en fonction de l'URL
    var template = e && e.parameter && e.parameter.page ? e.parameter.page : "TableauDeBord";
    Logger.log("Template à servir : " + template);
    
    // Générer le HTML approprié
    var htmlOutput;
    if (template === "test") {
      Logger.log("Génération de la page de test");
      htmlOutput = HtmlService.createHtmlOutputFromFile('TestSimple');
      htmlOutput.setTitle('Test de connexion');
    } else {
      Logger.log("Génération du tableau de bord");
      htmlOutput = HtmlService.createHtmlOutputFromFile('TableauDeBord');
      htmlOutput.setTitle('Tableau de Bord Rendement');
    }
    
    // Configurer les options de sécurité
    htmlOutput.setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
    
    Logger.log("Fin de l'exécution de doGet() - succès");
    return htmlOutput;
    
  } catch (e) {
    Logger.log("ERREUR dans doGet : " + e.toString());
    Logger.log("Stack trace : " + e.stack);
    
    // Créer une page d'erreur
    var htmlOutput = HtmlService.createHtmlOutput(
      '<html><body>' +
      '<h1>Erreur</h1>' +
      '<p>Une erreur s\'est produite : ' + e.toString() + '</p>' +
      '<p><a href="javascript:history.back()">Retour</a></p>' +
      '</body></html>'
    );
    
    htmlOutput.setTitle('Erreur');
    return htmlOutput;
  }
}

/**
 * Vérifie si l'utilisateur est autorisé à accéder à l'application
 * @return {Object} - Objet contenant le statut d'autorisation et des informations sur l'utilisateur
 */
function verifierAutorisation() {
  try {
    // Obtenir l'email de l'utilisateur actuel
    const userEmail = Session.getEffectiveUser().getEmail();
    
    // Ouvrir le classeur et la feuille des autorisations
    const spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetAuth = spreadsheet.getSheetByName('Autorisation');
    
    // Si la feuille n'existe pas, journaliser l'erreur
    if (!sheetAuth) {
      Logger.log("ERREUR: La feuille 'Autorisation' n'existe pas dans le classeur.");
      return {
        autorise: false,
        email: userEmail,
        message: "Configuration incorrecte: La feuille des autorisations n'existe pas."
      };
    }
    
    // Récupérer toutes les données de la feuille
    const range = sheetAuth.getDataRange();
    const values = range.getValues();
    
    // Identifier l'index de la colonne des emails (supposons que c'est la première colonne)
    let emailColIndex = 0;
    
    // Si la feuille contient un en-tête, vérifier la colonne des emails
    if (values.length > 0 && values[0].length > 0) {
      // Parcourir la première ligne pour trouver une colonne nommée "Email" ou similaire
      for (let i = 0; i < values[0].length; i++) {
        const header = String(values[0][i] || "").toLowerCase();
        if (header === "email" || header === "courriel" || header === "mail" || header === "adresse" || header === "utilisateur") {
          emailColIndex = i;
          break;
        }
      }
    }
    
    // Vérifier si l'email de l'utilisateur est dans la liste
    let isAuthorized = false;
    
    // Parcourir toutes les lignes (en sautant l'en-tête si présent)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Vérifier si la ligne contient suffisamment de colonnes
      if (row.length > emailColIndex) {
        const emailCell = String(row[emailColIndex] || "").trim().toLowerCase();
        
        // Comparer avec l'email de l'utilisateur (insensible à la casse)
        if (emailCell === userEmail.toLowerCase()) {
          isAuthorized = true;
          break;
        }
      }
    }
    
    // Journaliser le résultat
    Logger.log(`Vérification d'autorisation pour ${userEmail}: ${isAuthorized ? "Autorisé" : "Non autorisé"}`);
    
    // Retourner le résultat
    return {
      autorise: isAuthorized,
      email: userEmail,
      message: isAuthorized ? "Utilisateur autorisé" : "Vous n'êtes pas autorisé à accéder à cette application. Veuillez contacter le gestionnaire de l'application pour demander l'accès."
    };
    
  } catch (e) {
    Logger.log("ERREUR dans verifierAutorisation: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    
    return {
      autorise: false,
      email: Session.getEffectiveUser().getEmail(),
      message: "Une erreur est survenue lors de la vérification des autorisations: " + e.toString()
    };
  }
}

/**
 * Obtenir la liste des utilisateurs autorisés
 * @return {Object} - Liste des utilisateurs autorisés
 */
function obtenirUtilisateursAutorises() {
  try {
    // Ouvrir le classeur et la feuille des autorisations
    const spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetAuth = spreadsheet.getSheetByName('Autorisation');
    
    // Si la feuille n'existe pas, créer la feuille avec un en-tête
    if (!sheetAuth) {
      Logger.log("La feuille 'Autorisation' n'existe pas. Création de la feuille.");
      const newSheet = spreadsheet.insertSheet('Autorisation');
      
      // Ajouter les en-têtes
      newSheet.getRange(1, 1, 1, 3).setValues([['Email', 'Nom', 'Date d\'ajout']]);
      newSheet.getRange(1, 1, 1, 3).setFontWeight('bold');
      
      // Formater la feuille
      newSheet.setFrozenRows(1);
      newSheet.autoResizeColumns(1, 3);
      
      return {
        success: true,
        utilisateurs: []
      };
    }
    
    // Récupérer toutes les données de la feuille
    const range = sheetAuth.getDataRange();
    const values = range.getValues();
    
    // Si la feuille est vide, ajouter les en-têtes
    if (values.length === 0) {
      sheetAuth.getRange(1, 1, 1, 3).setValues([['Email', 'Nom', 'Date d\'ajout']]);
      sheetAuth.getRange(1, 1, 1, 3).setFontWeight('bold');
      sheetAuth.setFrozenRows(1);
      
      return {
        success: true,
        utilisateurs: []
      };
    }
    
    // Identifier les indices des colonnes
    let emailColIndex = 0;
    let nomColIndex = 1;
    let dateColIndex = 2;
    
    // Si la feuille contient un en-tête, trouver les colonnes appropriées
    if (values.length > 0) {
      for (let i = 0; i < values[0].length; i++) {
        const header = String(values[0][i] || "").toLowerCase();
        if (header === "email" || header === "courriel" || header === "mail") {
          emailColIndex = i;
        } else if (header === "nom" || header === "utilisateur" || header === "name") {
          nomColIndex = i;
        } else if (header.includes("date") || header.includes("ajout")) {
          dateColIndex = i;
        }
      }
    }
    
    // Créer une liste d'utilisateurs (en sautant la ligne d'en-tête)
    const utilisateurs = [];
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      
      // Vérifier si la ligne contient suffisamment de colonnes
      if (row.length > Math.max(emailColIndex, nomColIndex, dateColIndex)) {
        // Créer un objet utilisateur
        const utilisateur = {
          email: row[emailColIndex] || "",
          nom: row[nomColIndex] || "",
          dateAjout: row[dateColIndex] || ""
        };
        
        // Formater la date si c'est un objet Date
        if (utilisateur.dateAjout instanceof Date) {
          utilisateur.dateAjout = Utilities.formatDate(utilisateur.dateAjout, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
        }
        
        // Ajouter l'utilisateur à la liste
        utilisateurs.push(utilisateur);
      }
    }
    
    return {
      success: true,
      utilisateurs: utilisateurs
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirUtilisateursAutorises: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    
    return {
      success: false,
      message: "Erreur lors de la récupération des utilisateurs: " + e.toString(),
      utilisateurs: []
    };
  }
}

/**
 * Ajouter un nouvel utilisateur autorisé
 * @param {string} email - Adresse email de l'utilisateur
 * @param {string} nom - Nom de l'utilisateur
 * @return {Object} - Résultat de l'opération
 */
function ajouterUtilisateurAutorise(email, nom) {
  try {
    // Valider l'email
    if (!email || !validateEmail(email)) {
      return {
        success: false,
        message: "Adresse email invalide"
      };
    }
    
    // Ouvrir le classeur et la feuille des autorisations
    const spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    let sheetAuth = spreadsheet.getSheetByName('Autorisation');
    
    // Si la feuille n'existe pas, la créer avec un en-tête
    if (!sheetAuth) {
      Logger.log("La feuille 'Autorisation' n'existe pas. Création de la feuille.");
      sheetAuth = spreadsheet.insertSheet('Autorisation');
      
      // Ajouter les en-têtes
      sheetAuth.getRange(1, 1, 1, 3).setValues([['Email', 'Nom', 'Date d\'ajout']]);
      sheetAuth.getRange(1, 1, 1, 3).setFontWeight('bold');
      
      // Formater la feuille
      sheetAuth.setFrozenRows(1);
      sheetAuth.autoResizeColumns(1, 3);
    }
    
    // Récupérer toutes les données de la feuille
    const range = sheetAuth.getDataRange();
    const values = range.getValues();
    
    // Si la feuille est vide (pas d'en-tête), ajouter les en-têtes
    if (values.length === 0) {
      sheetAuth.getRange(1, 1, 1, 3).setValues([['Email', 'Nom', 'Date d\'ajout']]);
      sheetAuth.getRange(1, 1, 1, 3).setFontWeight('bold');
      sheetAuth.setFrozenRows(1);
    }
    
    // Vérifier si l'email existe déjà (ignorer la première ligne d'en-tête)
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString().toLowerCase() === email.toLowerCase()) {
        return {
          success: false,
          message: "Cet utilisateur est déjà autorisé"
        };
      }
    }
    
    // Ajouter le nouvel utilisateur
    const now = new Date();
    const dateFormatted = Utilities.formatDate(now, Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm:ss");
    sheetAuth.appendRow([email, nom || "", now]);
    
    // Formater la ligne ajoutée
    sheetAuth.autoResizeColumns(1, 3);
    
    return {
      success: true,
      message: "Utilisateur ajouté avec succès"
    };
    
  } catch (e) {
    Logger.log("ERREUR dans ajouterUtilisateurAutorise: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    
    return {
      success: false,
      message: "Erreur lors de l'ajout de l'utilisateur: " + e.toString()
    };
  }
}

/**
 * Supprimer un utilisateur autorisé
 * @param {string} email - Adresse email de l'utilisateur à supprimer
 * @return {Object} - Résultat de l'opération
 */
function supprimerUtilisateurAutorise(email) {
  try {
    // Valider l'email
    if (!email) {
      return {
        success: false,
        message: "Adresse email manquante"
      };
    }
    
    // Normaliser l'email (minuscules)
    const emailNormalized = email.toLowerCase();
    
    // Empêcher la suppression de l'utilisateur actuel
    const currentUserEmail = Session.getEffectiveUser().getEmail().toLowerCase();
    if (emailNormalized === currentUserEmail) {
      return {
        success: false,
        message: "Vous ne pouvez pas supprimer votre propre accès"
      };
    }
    
    // Ouvrir le classeur et la feuille des autorisations
    const spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    const spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    const sheetAuth = spreadsheet.getSheetByName('Autorisation');
    
    // Si la feuille n'existe pas, retourner une erreur
    if (!sheetAuth) {
      return {
        success: false,
        message: "La feuille des autorisations n'existe pas"
      };
    }
    
    // Récupérer toutes les données de la feuille
    const range = sheetAuth.getDataRange();
    const values = range.getValues();
    
    // Chercher l'utilisateur (ignorer la première ligne d'en-tête)
    let rowToDelete = -1;
    for (let i = 1; i < values.length; i++) {
      if (values[i][0] && values[i][0].toString().toLowerCase() === emailNormalized) {
        rowToDelete = i + 1; // +1 car les indices de lignes dans Sheets commencent à 1
        break;
      }
    }
    
    // Si l'utilisateur n'a pas été trouvé
    if (rowToDelete === -1) {
      return {
        success: false,
        message: "Utilisateur non trouvé"
      };
    }
    
    // Supprimer la ligne
    sheetAuth.deleteRow(rowToDelete);
    
    return {
      success: true,
      message: "Utilisateur supprimé avec succès"
    };
    
  } catch (e) {
    Logger.log("ERREUR dans supprimerUtilisateurAutorise: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    
    return {
      success: false,
      message: "Erreur lors de la suppression de l'utilisateur: " + e.toString()
    };
  }
}

// Fonction utilitaire pour valider un email
function validateEmail(email) {
  const re = /^(([^<>()\[\]\\.,;:\s@"]+(\.[^<>()\[\]\\.,;:\s@"]+)*)|(".+"))@((\[[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\.[0-9]{1,3}\])|(([a-zA-Z\-0-9]+\.)+[a-zA-Z]{2,}))$/;
  return re.test(String(email).toLowerCase());
}

// Fonction globale appelée lors de l'ouverture de la feuille
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  ui.createMenu('Analyse Rendement')
    .addItem('Ouvrir le tableau de bord', 'ouvrirTableauDeBord')
    .addToUi();
}

// Fonction qui ouvre l'application web
function ouvrirTableauDeBord() {
  var html = HtmlService.createHtmlOutputFromFile('TableauDeBord')
    .setWidth(1200)
    .setHeight(800)
    .setTitle('Tableau de Bord Rendement');
  SpreadsheetApp.getUi().showModalDialog(html, 'Tableau de Bord Rendement');
}

/**
 * Fonction pour récupérer la liste des opérateurs qui ont des aléas
 * @param {Object} filtres - Les filtres actifs (dateDebut, dateFin, equipe, poste)
 * @return {Object} - Un objet avec un champ success et un tableau d'opérateurs
 */
function obtenirOperateursAvecAleas(filtres) {
  try {
    Logger.log("Démarrage de obtenirOperateursAvecAleas() avec filtres : " + JSON.stringify(filtres));
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var ssAleas = spreadsheet.getSheetByName('Aleas');
   
    if (!ssAleas) return { success: false, message: "Feuille Aléas non trouvée" };
    
    // Récupérer toutes les données
    const data = ssAleas.getDataRange().getValues();
    if (data.length <= 1) return { success: true, operateurs: [] }; // Retourner un tableau vide si seulement l'en-tête
    
    // Déterminer les indices des colonnes - ajustez ces indices selon la structure de votre feuille
    const dateColIndex = 4;       // Colonne de la date 
    const operateurColIndex = 0;  // Colonne de l'opérateur 
    const posteColIndex = 3;      // Colonne du poste 
    const equipeColIndex = 2;     // Colonne de l'équipe 
    
    // Filtrer les données selon les critères
    let filteredData = data.slice(1); // Exclure l'en-tête
    
    // Convertir les filtres de date une seule fois
    let dateDebut = null;
    let dateFin = null;
    
    if (filtres) {
      if (filtres.dateDebut) {
        dateDebut = new Date(filtres.dateDebut);
        Logger.log("Date début convertie: " + dateDebut);
      }
      
      if (filtres.dateFin) {
        dateFin = new Date(filtres.dateFin);
        // Définir l'heure à la fin de la journée pour inclure toute la journée
        dateFin.setHours(23, 59, 59, 999);
        Logger.log("Date fin convertie: " + dateFin);
      }
    }
    
    // Appliquer les filtres si présents
    if (filtres) {
      if (dateDebut || dateFin) {
        filteredData = filteredData.filter(row => {
          let cellDate = null;
          
          // Si la date est au format JJ/MM/AAAA dans votre feuille
          if (typeof row[dateColIndex] === 'string') {
            const parts = row[dateColIndex].split('/');
            if (parts.length === 3) {
              // Format DD/MM/YYYY -> new Date(YYYY, MM-1, DD)
              cellDate = new Date(parts[2], parseInt(parts[1]) - 1, parseInt(parts[0]));
            }
          }
          // Si la date est déjà un objet Date
          else if (row[dateColIndex] instanceof Date) {
            cellDate = new Date(row[dateColIndex]);
          }
          
          // Si on n'a pas pu obtenir une date valide, ignorer cette ligne
          if (!cellDate || isNaN(cellDate.getTime())) {
            return false;
          }
          
          // Appliquer les filtres de date
          if (dateDebut && cellDate < dateDebut) {
            return false;
          }
          
          if (dateFin && cellDate > dateFin) {
            return false;
          }
          
          return true;
        });
        
        Logger.log("Après filtrage par date: " + filteredData.length + " lignes");
      }
      
      if (filtres.poste) {
        filteredData = filteredData.filter(row => row[posteColIndex] === filtres.poste);
        Logger.log("Après filtrage par poste: " + filteredData.length + " lignes");
      }
      
      if (filtres.equipe) {
        filteredData = filteredData.filter(row => row[equipeColIndex] === filtres.equipe);
        Logger.log("Après filtrage par équipe: " + filteredData.length + " lignes");
      }
    }
    
    // Extraire la liste unique des opérateurs ayant des aléas
    const operateurs = [...new Set(filteredData.map(row => row[operateurColIndex]))];
    Logger.log("Nombre d'opérateurs avec aléas trouvés: " + operateurs.length);
    
    return { success: true, operateurs: operateurs };
  } catch (error) {
    Logger.log("ERREUR dans obtenirOperateursAvecAleas: " + error.toString());
    return { success: false, message: "Erreur: " + error.toString() };
  }
}

/**
 * Fonction pour obtenir la durée des aléas par opérateur
 * @param {Array} operateurs - Liste des opérateurs à traiter
 * @param {Object} filtres - Filtres à appliquer (date, équipe, poste)
 * @return {Object} - Objet contenant les durées d'aléas par opérateur
 */
function obtenirDureesAleasParOperateur(operateurs, filtres) {
  try {
    Logger.log("Récupération des durées d'aléas par opérateur");
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // En-têtes
    
    // Trouver les indices des colonnes
    var operateurIndex = -1;
    var dateIndex = -1;
    var dureeIndex = -1;
    var posteIndex = -1;
    var equipeIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase();
      if (header === "opérateur" || header === "operateur") {
        operateurIndex = i;
      } else if (header === "date") {
        dateIndex = i;
      } else if (header === "durée" || header === "duree") {
        dureeIndex = i;
      } else if (header === "poste") {
        posteIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      }
    }
    
    if (operateurIndex < 0 || dureeIndex < 0) {
      return { success: false, message: "Colonnes requises introuvables" };
    }
    
    var result = {};
    // Initialiser les résultats pour tous les opérateurs
    if (operateurs && operateurs.length) {
      operateurs.forEach(function(operateur) {
        result[operateur] = 0;
      });
    }
    
    // Parcourir les données d'aléas
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var operateur = row[operateurIndex];
      var duree = parseFloat(row[dureeIndex]) || 0;
      
      // Vérifier si cet opérateur nous intéresse
      if (duree > 0 && (!operateurs || operateurs.length === 0 || operateurs.indexOf(operateur) >= 0)) {
        
        // Appliquer les filtres si nécessaires
        var inclure = true;
        
        if (filtres) {
          // Filtre de date
          if ((filtres.dateDebut || filtres.dateFin) && dateIndex >= 0) {
            var dateValue = row[dateIndex];
            var aleaDate;
            
            if (dateValue instanceof Date) {
              aleaDate = dateValue;
            } else if (typeof dateValue === 'string') {
              // Tenter de convertir une chaîne de date (format DD/MM/YYYY)
              var parts = dateValue.split('/');
              if (parts.length === 3) {
                aleaDate = new Date(parts[2], parts[1]-1, parts[0]);
              }
            }
            
            if (aleaDate) {
              if (filtres.dateDebut) {
                var dateDebut = new Date(filtres.dateDebut);
                if (aleaDate < dateDebut) inclure = false;
              }
              
              if (filtres.dateFin) {
                var dateFin = new Date(filtres.dateFin);
                dateFin.setHours(23, 59, 59, 999); // Fin de journée
                if (aleaDate > dateFin) inclure = false;
              }
            }
          }
          
          // Filtre d'équipe
          if (filtres.equipe && equipeIndex >= 0) {
            if (row[equipeIndex] !== filtres.equipe) inclure = false;
          }
          
          // Filtre de poste
          if (filtres.poste && posteIndex >= 0) {
            if (row[posteIndex] !== filtres.poste) inclure = false;
          }
        }
        
        // Si l'enregistrement passe tous les filtres, ajouter la durée
        if (inclure) {
          if (!result[operateur]) result[operateur] = 0;
          result[operateur] += duree;
        }
      }
    }
    
    return {
      success: true,
      durees: result
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDureesAleasParEquipe: " + e.toString());
    return {
      success: false, 
      message: "Erreur: " + e.toString()
    };
  }
}
 

/**
 * Fonction pour obtenir la durée des aléas par poste
 * @param {Object} filtres - Filtres à appliquer (date, équipe)
 * @return {Object} - Objet contenant les durées d'aléas par poste
 */
function obtenirDureesAleasParPoste(filtres) {
  try {
    Logger.log("Récupération des durées d'aléas par poste");
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // En-têtes
    
    // Trouver les indices des colonnes
    var posteIndex = -1;
    var dateIndex = -1;
    var dureeIndex = -1;
    var equipeIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase();
      if (header === "poste") {
        posteIndex = i;
      } else if (header === "date") {
        dateIndex = i;
      } else if (header === "durée" || header === "duree") {
        dureeIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      }
    }
    
    if (posteIndex < 0 || dureeIndex < 0) {
      return { success: false, message: "Colonnes requises introuvables" };
    }
    
    var dureeParPoste = {};
    
    // Parcourir les données d'aléas
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var poste = row[posteIndex] || "Non spécifié";
      var duree = parseFloat(row[dureeIndex]) || 0;
      
      if (duree > 0) {
        // Appliquer les filtres si nécessaires
        var inclure = true;
        
        if (filtres) {
          // Filtre de date
          if ((filtres.dateDebut || filtres.dateFin) && dateIndex >= 0) {
            var dateValue = row[dateIndex];
            var aleaDate;
            
            if (dateValue instanceof Date) {
              aleaDate = dateValue;
            } else if (typeof dateValue === 'string') {
              // Tenter de convertir une chaîne de date (format DD/MM/YYYY)
              var parts = dateValue.split('/');
              if (parts.length === 3) {
                aleaDate = new Date(parts[2], parts[1]-1, parts[0]);
              }
            }
            
            if (aleaDate) {
              if (filtres.dateDebut) {
                var dateDebut = new Date(filtres.dateDebut);
                if (aleaDate < dateDebut) inclure = false;
              }
              
              if (filtres.dateFin) {
                var dateFin = new Date(filtres.dateFin);
                dateFin.setHours(23, 59, 59, 999); // Fin de journée
                if (aleaDate > dateFin) inclure = false;
              }
            }
          }
          
          // Filtre d'équipe
          if (filtres.equipe && equipeIndex >= 0) {
            if (row[equipeIndex] !== filtres.equipe) inclure = false;
          }
        }
        
        // Si l'enregistrement passe tous les filtres, ajouter la durée
        if (inclure) {
          if (!dureeParPoste[poste]) dureeParPoste[poste] = 0;
          dureeParPoste[poste] += duree;
        }
      }
    }
    
    // Convertir en tableau de résultats
    var result = [];
    for (var poste in dureeParPoste) {
      result.push({
        poste: poste,
        duree: dureeParPoste[poste]
      });
    }
    
    // Trier par durée décroissante
    result.sort(function(a, b) {
      return b.duree - a.duree;
    });
    
    return {
      success: true,
      durees: result
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDureesAleasParPoste: " + e.toString());
    return {
      success: false, 
      message: "Erreur: " + e.toString()
    };
  }
}
/**
 * Fonction pour obtenir la durée des aléas par équipe
 * @param {Object} filtres - Filtres à appliquer (date, poste)
 * @return {Object} - Objet contenant les durées d'aléas par équipe
 */
function obtenirDureesAleasParEquipe(filtres) {
  try {
    Logger.log("Récupération des durées d'aléas par équipe");
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // En-têtes
    
    // Trouver les indices des colonnes
    var equipeIndex = -1;
    var dateIndex = -1;
    var dureeIndex = -1;
    var posteIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase();
      if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      } else if (header === "date") {
        dateIndex = i;
      } else if (header === "durée" || header === "duree") {
        dureeIndex = i;
      } else if (header === "poste") {
        posteIndex = i;
      }
    }
    
    if (equipeIndex < 0 || dureeIndex < 0) {
      return { success: false, message: "Colonnes requises introuvables" };
    }
    
    var dureeParEquipe = {};
    
    // Parcourir les données d'aléas
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var equipe = row[equipeIndex] || "Non spécifiée";
      var duree = parseFloat(row[dureeIndex]) || 0;
      
      if (duree > 0) {
        // Appliquer les filtres si nécessaires
        var inclure = true;
        
        if (filtres) {
          // Filtre de date
          if ((filtres.dateDebut || filtres.dateFin) && dateIndex >= 0) {
            var dateValue = row[dateIndex];
            var aleaDate;
            
            if (dateValue instanceof Date) {
              aleaDate = dateValue;
            } else if (typeof dateValue === 'string') {
              // Tenter de convertir une chaîne de date (format DD/MM/YYYY)
              var parts = dateValue.split('/');
              if (parts.length === 3) {
                aleaDate = new Date(parts[2], parts[1]-1, parts[0]);
              }
            }
            
            if (aleaDate) {
              if (filtres.dateDebut) {
                var dateDebut = new Date(filtres.dateDebut);
                if (aleaDate < dateDebut) inclure = false;
              }
              
              if (filtres.dateFin) {
                var dateFin = new Date(filtres.dateFin);
                dateFin.setHours(23, 59, 59, 999); // Fin de journée
                if (aleaDate > dateFin) inclure = false;
              }
            }
          }
          
          // Filtre de poste
          if (filtres.poste && posteIndex >= 0) {
            if (row[posteIndex] !== filtres.poste) inclure = false;
          }
        }
        
        // Si l'enregistrement passe tous les filtres, ajouter la durée
        if (inclure) {
          if (!dureeParEquipe[equipe]) dureeParEquipe[equipe] = 0;
          dureeParEquipe[equipe] += duree;
        }
      }
    }
    
    // Convertir en tableau de résultats
    var result = [];
    for (var equipe in dureeParEquipe) {
      result.push({
        equipe: equipe,
        duree: dureeParEquipe[equipe]
      });
    }
    
    // Trier par durée décroissante
    result.sort(function(a, b) {
      return b.duree - a.duree;
    });
    
    return {
      success: true,
      durees: result
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDureesAleasParEquipe: " + e.toString());
    return {
      success: false, 
      message: "Erreur: " + e.toString()
    };
  }
}
/**
 * Fonction pour obtenir les durées d'aléas d'un opérateur par période
 * @param {string} operateur - Nom de l'opérateur
 * @param {Object} filtres - Filtres à appliquer (date, poste, équipe)
 * @return {Object} - Objet contenant les durées d'aléas par période
 */
function obtenirDureesAleasOperateurParPeriode(operateur, filtres) {
  try {
    Logger.log("Récupération des durées d'aléas par période pour l'opérateur: " + operateur);
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // En-têtes
    
    // Trouver les indices des colonnes
    var operateurIndex = -1;
    var dateIndex = -1;
    var dureeIndex = -1;
    var posteIndex = -1;
    var equipeIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase();
      if (header === "opérateur" || header === "operateur") {
        operateurIndex = i;
      } else if (header === "date") {
        dateIndex = i;
      } else if (header === "durée" || header === "duree") {
        dureeIndex = i;
      } else if (header === "poste") {
        posteIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      }
    }
    
    if (operateurIndex < 0 || dateIndex < 0 || dureeIndex < 0) {
      return { success: false, message: "Colonnes requises introuvables" };
    }
    
    // Dictionnaire pour stocker les durées par période
    var dureeParPeriode = {};
    
    // Parcourir les données d'aléas
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var rowOperateur = row[operateurIndex];
      var duree = parseFloat(row[dureeIndex]) || 0;
      
      // Vérifier si cet opérateur nous intéresse
      if (rowOperateur === operateur && duree > 0) {
        var dateValue = row[dateIndex];
        var aleaDate;
        
        // Convertir la date
        if (dateValue instanceof Date) {
          aleaDate = dateValue;
        } else if (typeof dateValue === 'string') {
          // Tenter de convertir une chaîne de date (format DD/MM/YYYY)
          var parts = dateValue.split('/');
          if (parts.length === 3) {
            aleaDate = new Date(parts[2], parts[1]-1, parts[0]);
          }
        }
        
        if (!aleaDate || isNaN(aleaDate.getTime())) {
          continue; // Date invalide, ignorer cette ligne
        }
        
        // Appliquer les filtres si nécessaires
        var inclure = true;
        
        if (filtres) {
          // Filtre de date
          if (filtres.dateDebut) {
            var dateDebut = new Date(filtres.dateDebut);
            if (aleaDate < dateDebut) inclure = false;
          }
          
          if (filtres.dateFin) {
            var dateFin = new Date(filtres.dateFin);
            dateFin.setHours(23, 59, 59, 999); // Fin de journée
            if (aleaDate > dateFin) inclure = false;
          }
          
          // Filtre d'équipe
          if (filtres.equipe && equipeIndex >= 0) {
            if (row[equipeIndex] !== filtres.equipe) inclure = false;
          }
          
          // Filtre de poste
          if (filtres.poste && posteIndex >= 0) {
            if (row[posteIndex] !== filtres.poste) inclure = false;
          }
        }
        
        // Si l'enregistrement passe tous les filtres, ajouter la durée
        if (inclure) {
          // Déterminer la période (semaine) de cette date
          var annee = aleaDate.getFullYear();
          var startOfYear = new Date(annee, 0, 1);
          var days = Math.floor((aleaDate - startOfYear) / (24 * 60 * 60 * 1000));
          var weekNum = Math.ceil((days + startOfYear.getDay() + 1) / 7);
          
          // Formatter la période comme "YYYY-SXX"
          var periode = annee + "-S" + String(weekNum).padStart(2, '0');
          
          if (!dureeParPeriode[periode]) dureeParPeriode[periode] = 0;
          dureeParPeriode[periode] += duree;
        }
      }
    }
    
    return {
      success: true,
      durees: dureeParPeriode
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDureesAleasOperateurParPeriode: " + e.toString());
    return {
      success: false, 
      message: "Erreur: " + e.toString()
    };
  }
}

/**
 * Fonction pour obtenir les durées d'aléas par période
 * @param {string} typePeriode - Type de période ('annee', 'mois', 'semaine')
 * @param {Object} filtres - Filtres à appliquer (date, équipe, poste)
 * @return {Object} - Objet contenant les durées d'aléas par période
 */
function obtenirDureesAleasParPeriode(typePeriode, filtres) {
  try {
    Logger.log("Récupération des durées d'aléas par " + typePeriode);
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // En-têtes
    
    // Trouver les indices des colonnes
    var dateIndex = -1;
    var dureeIndex = -1;
    var posteIndex = -1;
    var equipeIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase();
      if (header === "date") {
        dateIndex = i;
      } else if (header === "durée" || header === "duree") {
        dureeIndex = i;
      } else if (header === "poste") {
        posteIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      }
    }
    
    if (dateIndex < 0 || dureeIndex < 0) {
      return { success: false, message: "Colonnes requises introuvables" };
    }
    
    // Dictionnaire pour stocker les durées par période
    var dureeParPeriode = {};
    
    // Parcourir les données d'aléas
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var duree = parseFloat(row[dureeIndex]) || 0;
      
      if (duree > 0) {
        var dateValue = row[dateIndex];
        var aleaDate;
        
        // Convertir la date
        if (dateValue instanceof Date) {
          aleaDate = dateValue;
        } else if (typeof dateValue === 'string') {
          // Tenter de convertir une chaîne de date (format DD/MM/YYYY)
          var parts = dateValue.split('/');
          if (parts.length === 3) {
            aleaDate = new Date(parts[2], parts[1]-1, parts[0]);
          }
        }
        
        if (!aleaDate || isNaN(aleaDate.getTime())) {
          continue; // Date invalide, ignorer cette ligne
        }
        
        // Appliquer les filtres si nécessaires
        var inclure = true;
        
        if (filtres) {
          // Filtre de date
          if (filtres.dateDebut) {
            var dateDebut = new Date(filtres.dateDebut);
            if (aleaDate < dateDebut) inclure = false;
          }
          
          if (filtres.dateFin) {
            var dateFin = new Date(filtres.dateFin);
            dateFin.setHours(23, 59, 59, 999); // Fin de journée
            if (aleaDate > dateFin) inclure = false;
          }
          
          // Filtre d'équipe
          if (filtres.equipe && equipeIndex >= 0) {
            if (row[equipeIndex] !== filtres.equipe) inclure = false;
          }
          
          // Filtre de poste
          if (filtres.poste && posteIndex >= 0) {
            if (row[posteIndex] !== filtres.poste) inclure = false;
          }
        }
        
        // Si l'enregistrement passe tous les filtres, déterminer la période et ajouter la durée
        if (inclure) {
          var periode;
          var annee = aleaDate.getFullYear();
          
          if (typePeriode === 'annee') {
            periode = String(annee);
          } else if (typePeriode === 'mois') {
            var mois = aleaDate.getMonth() + 1;
            periode = annee + '-' + String(mois).padStart(2, '0');
          } else if (typePeriode === 'semaine') {
            var startOfYear = new Date(annee, 0, 1);
            var days = Math.floor((aleaDate - startOfYear) / (24 * 60 * 60 * 1000));
            var weekNum = Math.ceil((days + startOfYear.getDay() + 1) / 7);
            periode = annee + '-S' + String(weekNum).padStart(2, '0');
          } else {
            continue; // Type de période non reconnu
          }
          
          if (!dureeParPeriode[periode]) dureeParPeriode[periode] = 0;
          dureeParPeriode[periode] += duree;
        }
      }
    }
    
    // Convertir en tableau pour le retour
    var result = [];
    for (var periode in dureeParPeriode) {
      result.push({
        periode: periode,
        duree: dureeParPeriode[periode]
      });
    }
    
    // Trier par période
    result.sort(function(a, b) {
      return a.periode.localeCompare(b.periode);
    });
    
    return {
      success: true,
      durees: result
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDureesAleasParPeriode: " + e.toString());
    return {
      success: false, 
      message: "Erreur: " + e.toString()
    };
  }
}

// Fonction pour obtenir les données de la feuille
// Fonction pour obtenir les données de la feuille
function obtenirDonnees() {
  try {
    Logger.log("Début de l'exécution de obtenirDonnees()");
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8'; // Remplacez par votre ID de Spreadsheet
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement'); // Assurez-vous que le nom est correct

    if (!sheet) {
      Logger.log("ERREUR: Feuille 'Historique rendement' introuvable.");
      return { erreur: "Feuille 'Historique rendement' introuvable", donnees: [], postes: [], equipes: [], operateurs: [] };
    }

    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();

    // Vérifier si la feuille contient au moins 2 lignes pour les en-têtes
    if (values.length < 2) {
       Logger.log("ERREUR: La feuille 'Historique rendement' ne contient pas d'en-têtes à la ligne 2.");
       return { erreur: "Structure de feuille incorrecte: en-têtes manquants à la ligne 2.", donnees: [], postes: [], equipes: [], operateurs: [] };
    }

    // Les en-têtes sont dans la 2ème ligne (index 1)
    var entetes = values[1];
    var donnees = [];

    // ---- Identification des Indices des Colonnes ----
    var indexOperateur = entetes.indexOf("Opérateur");
    var indexEquipe = entetes.indexOf("Equipe");
    var indexDate = entetes.indexOf("Date");
    var indexAnnee = entetes.indexOf("Année");
    var indexSemaine = entetes.indexOf("Semaine");
    var indexPosteHabituel = entetes.indexOf("Poste habituel");

    // Vérification que les colonnes de base existent
    if ([indexOperateur, indexEquipe, indexDate, indexAnnee, indexSemaine, indexPosteHabituel].includes(-1)) {
         Logger.log("ERREUR: Une ou plusieurs colonnes de base (Opérateur, Equipe, Date, Année, Semaine, Poste habituel) sont introuvables à la ligne 2.");
         return { erreur: "Structure de feuille incorrecte: colonnes de base manquantes.", donnees: [], postes: [], equipes: [], operateurs: [] };
    }

    // Colonnes TAI (Indices 15 à 21, soit P à V)
    var taiColumns = [];
    for (var i = 15; i <= 21; i++) {
        if (i < entetes.length && entetes[i]) {
            taiColumns.push({ nom: String(entetes[i]), index: i });
        }
    }

    // Colonnes présence (Indices 23 à 28, soit X à AC)
    var presenceColumns = [];
    for (var i = 23; i <= 28; i++) {
         if (i < entetes.length && entetes[i]) {
            presenceColumns.push({ nom: String(entetes[i]), index: i });
        }
    }

    // *** MODIFICATION : Identifier les colonnes d'inactivité (AD à AK => index 29 à 36) ***
    var inactiviteColumns = [];
    for (var k = 29; k <= 36; k++) { // Colonnes AD(29) à AK(36)
         if (k < entetes.length && entetes[k]) { // Vérifie si l'index existe et si l'en-tête n'est pas vide
            inactiviteColumns.push({ nom: String(entetes[k]), index: k });
        }
    }
    if (inactiviteColumns.length > 0) {
        Logger.log("Colonnes d'inactivité identifiées: " + inactiviteColumns.map(c => c.nom).join(", "));
    } else {
         Logger.log("ATTENTION: Aucune colonne d'inactivité trouvée entre AD et AK avec un nom dans la ligne 2.");
    }
    // *** FIN MODIFICATION ***

    // Indices MASQUAGE/DEMASQUAGE ... (code inchangé)
    var indexMasquage = entetes.indexOf("MASQUAGE");
    var indexDemasquage = entetes.indexOf("DEMASQUAGE");
    // ---- Fin Identification des Indices ----

    // Parcourir les données (à partir de la 3ème ligne, index 2)
    for (var i = 2; i < values.length; i++) {
      var row = values[i];

      // Vérifier que la ligne contient un opérateur
      if (row[indexOperateur]) {
        var dateStr = "";
         try {
             if (row[indexDate] instanceof Date && !isNaN(row[indexDate])) {
               dateStr = Utilities.formatDate(row[indexDate], Session.getScriptTimeZone(), "yyyy-MM-dd");
             } else if (typeof row[indexDate] === 'string' && row[indexDate].length > 0) {
               // Essayer de parser différents formats si nécessaire, ici on assume yyyy-MM-dd ou on garde tel quel
               dateStr = row[indexDate]; // Attention: peut nécessiter une validation/conversion plus poussée
             }
         } catch (e) {
              Logger.log("Erreur formatage date ligne " + (i+1) + ": " + row[indexDate] + " - Erreur: " + e);
              dateStr = ""; // Laisser vide si erreur
         }

        // Collecter TAI
        var taiTotal = 0;
        var taiParPoste = {};
        taiColumns.forEach(function(col) {
            var valeur = parseFloat(row[col.index]) || 0;
            taiParPoste[col.nom] = valeur;
            taiTotal += valeur;
        });

        // Collecter présence
        var presenceTotal = 0;
        var presenceParPoste = {};
        presenceColumns.forEach(function(col) {
             var valeur = parseFloat(row[col.index]) || 0;
             presenceParPoste[col.nom] = valeur;
             presenceTotal += valeur;
        });

        // Calculer rendement
        var rendement = presenceTotal > 0 ? taiTotal / presenceTotal : 0;

        // Créer l'objet donneeItem
        var donneeItem = {
          operateur: row[indexOperateur],
          equipe: row[indexEquipe] || "",
          poste: row[indexPosteHabituel] || "", // Poste "principal" pour info
          date: dateStr,
          annee: row[indexAnnee] || "",
          semaine: row[indexSemaine] || "",
          rendement: rendement,
          tai: taiTotal,
          presence: presenceTotal,
          taiParPoste: taiParPoste, // TAI détaillé par type de poste
          presenceParPoste: presenceParPoste, // Présence détaillée par type de poste
          // *** MODIFICATION : Ajouter l'objet inactivite ***
          inactivite: {}
          // *** FIN MODIFICATION ***
        };

         // *** MODIFICATION : Remplir les données d'inactivité ***
         inactiviteColumns.forEach(function(col) {
            donneeItem.inactivite[col.nom] = parseFloat(row[col.index]) || 0;
         });
         // *** FIN MODIFICATION ***

        // Ajouter MASQUAGE/DEMASQUAGE
        if (indexMasquage !== -1) {
            donneeItem.masquage = parseFloat(row[indexMasquage]) || 0;
        }
        if (indexDemasquage !== -1) {
            donneeItem.demasquage = parseFloat(row[indexDemasquage]) || 0;
        }

        donnees.push(donneeItem);
      }
    }

    // Extraire les listes uniques
    var postesTypes = taiColumns.map(c => c.nom); // Les types de postes sont les noms des colonnes TAI/Présence
    var equipes = [...new Set(donnees.map(item => item.equipe))].filter(e => e).sort();
    var operateurs = [...new Set(donnees.map(item => item.operateur))].filter(o => o).sort();

    Logger.log("obtenirDonnees: " + donnees.length + " enregistrements traités.");
    return {
      donnees: donnees,
      postes: postesTypes, // Liste des types de postes (issus des colonnes TAI/Présence)
      equipes: equipes,
      operateurs: operateurs,
      postesTypes: postesTypes // Maintenu pour compatibilité potentielle
    };

  } catch (e) {
    Logger.log("ERREUR GRAVE dans obtenirDonnees: " + e.toString() + "\nStack: " + e.stack);
    return {
      erreur: "Erreur serveur lors de la lecture des données: " + e.toString(),
      donnees: [], postes: [], equipes: [], operateurs: []
    };
  }
}


// Ajouter cette nouvelle fonction dans Code.gs
function obtenirDatePlusRecente() {
  try {
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (!sheet) {
      return { success: false, message: "Feuille 'Historique rendement' introuvable" };
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[1]; // Les en-têtes sont dans la 2ème ligne (index 1)
    
    // Trouver l'index de la colonne Date
    var dateIndex = headers.indexOf("Date");
    if (dateIndex === -1) {
      return { success: false, message: "Colonne 'Date' introuvable" };
    }
    
    // Parcourir les données pour trouver la date la plus récente
    var datePlusRecente = null;
    
    for (var i = 2; i < values.length; i++) {
      var row = values[i];
      var dateCell = row[dateIndex];
      
      if (dateCell instanceof Date && !isNaN(dateCell.getTime())) {
        if (!datePlusRecente || dateCell > datePlusRecente) {
          datePlusRecente = dateCell;
        }
      }
    }
    
    if (!datePlusRecente) {
      return { success: false, message: "Aucune date valide trouvée" };
    }
    
    // Formater la date au format YYYY-MM-DD pour les filtres
    var dateStr = Utilities.formatDate(datePlusRecente, Session.getScriptTimeZone(), "yyyy-MM-dd");
    
    return { 
      success: true, 
      date: dateStr
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDatePlusRecente: " + e.toString());
    return { success: false, message: e.toString() };
  }
}

// Ajouter cette fonction dans Code.gs
function obtenirListeOperateurs() {
  try {
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (!sheet) {
      return { 
        error: "Feuille 'Historique rendement' introuvable",
        operateurs: []
      };
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[1]; // Les en-têtes sont dans la 2ème ligne (index 1)
    
    // Obtenir l'index de la colonne Opérateur
    var indexOperateur = headers.indexOf("Opérateur");
    if (indexOperateur === -1) {
      return { 
        error: "Colonne 'Opérateur' introuvable",
        operateurs: []
      };
    }
    
    // Extraire tous les opérateurs (à partir de la 3ème ligne, index 2)
    var operateurs = [];
    for (var i = 2; i < values.length; i++) {
      var operateur = values[i][indexOperateur];
      if (operateur) {
        operateurs.push(operateur);
      }
    }
    
    // Éliminer les doublons et trier
    operateurs = [...new Set(operateurs)].sort();
    
    return {
      success: true,
      operateurs: operateurs
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirListeOperateurs: " + e.toString());
    return { 
      error: e.toString(),
      operateurs: []
    };
  }
}

// Fonction auxiliaire pour obtenir une liste unique de valeurs
function obtenirListeUnique(donnees, propriete) {
  var valeurs = donnees.map(function(item) {
    return item[propriete];
  });
  
  return [...new Set(valeurs)].filter(function(val) {
    // Vérifier d'abord si val est une chaîne avant d'appliquer trim()
    return val !== null && val !== undefined && 
           (typeof val !== 'string' || val.trim() !== "");
  }).sort();
}

// Fonction pour obtenir des statistiques par poste avec gestion des filtres
function obtenirDonnees() {
  try {
    Logger.log("Début de l'exécution de obtenirDonnees()");
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (!sheet) {
      return { 
        erreur: "Feuille 'Historique rendement' introuvable",
        donnees: [],
        postes: [],
        equipes: [],
        operateurs: []
      };
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    
    // Les en-têtes sont dans la 2ème ligne (index 1)
    var entetes = values[1];
    var donnees = [];
  
    // Obtenir les indices des colonnes de base
    var indexOperateur = entetes.indexOf("Opérateur");
    var indexEquipe = entetes.indexOf("Equipe");
    var indexDate = entetes.indexOf("Date");
    var indexAnnee = entetes.indexOf("Année");
    var indexSemaine = entetes.indexOf("Semaine");
    
    // Pour information seulement, nous n'utiliserons pas cette colonne pour les regroupements
    var indexPosteHabituel = entetes.indexOf("Poste habituel");
    
    // Colonnes TAI (P à V) - index 15 à 21
    var taiColumns = [];
    for (var i = 15; i <= 21; i++) {
      if (i < entetes.length && entetes[i]) {
        taiColumns.push({
          nom: String(entetes[i]),
          index: i
        });
      }
    }
    
    // Colonnes présence (X à AC) - index 23 à 28
    var presenceColumns = [];
    for (var i = 23; i <= 28; i++) {
      if (i < entetes.length && entetes[i]) {
        presenceColumns.push({
          nom: String(entetes[i]),
          index: i
        });
      }
    }
    
    Logger.log("Colonnes TAI trouvées: " + taiColumns.map(c => c.nom).join(", "));
    Logger.log("Colonnes présence trouvées: " + presenceColumns.map(c => c.nom).join(", "));
    
    // Rechercher MASQUAGE et DEMASQUAGE pour information
    var indexMasquage = -1;
    var indexDemasquage = -1;
    
    // Chercher par nom
    for (var i = 0; i < entetes.length; i++) {
      var headerName = String(entetes[i] || "");
      if (headerName === "MASQUAGE") {
        indexMasquage = i;
      } else if (headerName === "DEMASQUAGE") {
        indexDemasquage = i;
      }
    }
    
    // Parcourir les données (à partir de la 3ème ligne, index 2)
    for (var i = 2; i < values.length; i++) {
      var row = values[i];
      
      // Vérifier que la ligne contient des données valides
      if (row[indexOperateur]) {
        var dateStr = "";
        if (typeof row[indexDate] === 'object' && row[indexDate] instanceof Date) {
          dateStr = Utilities.formatDate(row[indexDate], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (typeof row[indexDate] === 'string') {
          dateStr = row[indexDate];
        }
        
        // Collecter les valeurs TAI par poste avec leur somme
        var taiTotal = 0;
        var taiParPoste = {};
        taiColumns.forEach(function(col) {
          var valeur = parseFloat(row[col.index]) || 0;
          taiParPoste[col.nom] = valeur;
          taiTotal += valeur;
        });
        
        // Collecter les valeurs de présence par poste avec leur somme
        var presenceTotal = 0;
        var presenceParPoste = {};
        presenceColumns.forEach(function(col) {
          var valeur = parseFloat(row[col.index]) || 0;
          presenceParPoste[col.nom] = valeur;
          presenceTotal += valeur;
        });
        
        // Calculer le rendement comme TAI Total / Présence Totale
        var rendement = presenceTotal > 0 ? taiTotal / presenceTotal : 0;
        
        // Ajouter l'entrée aux données
        var donneeItem = {
          operateur: row[indexOperateur],
          equipe: row[indexEquipe] || "",
          poste: row[indexPosteHabituel] || "", // Gardé pour la compatibilité, mais ne sera pas utilisé pour les regroupements
          date: dateStr,
          annee: row[indexAnnee] || "",
          semaine: row[indexSemaine] || "",
          rendement: rendement,
          tai: taiTotal,
          presence: presenceTotal,
          taiParPoste: taiParPoste,
          presenceParPoste: presenceParPoste
        };
        
        // Ajouter MASQUAGE et DEMASQUAGE si disponibles
        if (indexMasquage !== -1) {
          donneeItem.masquage = parseFloat(row[indexMasquage]) || 0;
        }
        if (indexDemasquage !== -1) {
          donneeItem.demasquage = parseFloat(row[indexDemasquage]) || 0;
        }
        
        donnees.push(donneeItem);
      }
    }
    
    // Extraire les listes uniques (maintenant basées uniquement sur les noms de colonne pour les postes)
    var postesTypes = taiColumns.map(c => c.nom);
    var equipes = [...new Set(donnees.map(item => item.equipe))].filter(e => e).sort();
    var operateurs = [...new Set(donnees.map(item => item.operateur))].filter(o => o).sort();
    
    return {
      donnees: donnees,
      postes: postesTypes, // Remplacer par les types de postes extraits des en-têtes
      equipes: equipes,
      operateurs: operateurs,
      postesTypes: postesTypes // Doublon pour compatibilité
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDonnees: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    return { 
      erreur: "Erreur lors de la lecture des données: " + e.toString(),
      donnees: [],
      postes: [],
      equipes: [],
      operateurs: []
    };
  }
}

/**
 * Fonction pour obtenir les statistiques d'inactivité
 * @param {Object} filtres - Filtres à appliquer
 * @return {Object} - Objet contenant les statistiques agrégées
 */
function obtenirStatistiquesInactivite(filtres) {
  try {
    Logger.log("Démarrage de obtenirStatistiquesInactivite() avec filtres: " + JSON.stringify(filtres));

    // 1. Obtenir directement les données du tableau
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (!sheet) {
      Logger.log("Feuille 'Historique rendement' introuvable");
      return { success: false, message: "Feuille 'Historique rendement' introuvable" };
    }
    
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[1]; // Les en-têtes sont dans la deuxième ligne (index 1)
    
    // 2. Identifier les colonnes d'inactivité (généralement colonnes AD à AK, indices 29-36)
    var inactiviteColumns = [];
    var inactiviteTypes = [];
    
    for (var i = 29; i <= 36; i++) {
      if (i < headers.length && headers[i]) {
        inactiviteColumns.push(i);
        inactiviteTypes.push(String(headers[i]));
        Logger.log("Colonne d'inactivité identifiée: " + headers[i] + " à l'index " + i);
      }
    }
    
    if (inactiviteColumns.length === 0) {
      Logger.log("Aucune colonne d'inactivité trouvée");
      return { success: false, message: "Aucune colonne d'inactivité trouvée" };
    }
    
    // 3. Identifier d'autres colonnes importantes
    var dateIndex = -1;
    var operateurIndex = -1;
    var equipeIndex = -1;
    var posteIndex = -1;
    var presenceIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase();
      if (header === "date") dateIndex = i;
      else if (header === "opérateur" || header === "operateur") operateurIndex = i;
      else if (header === "equipe" || header === "équipe") equipeIndex = i;
      else if (header === "poste habituel") posteIndex = i;
      // Pour les présences, chercher une colonne avec "présence totale" ou similaire
      else if (header.includes("présence") && (header.includes("total") || header.includes("totale"))) {
        presenceIndex = i;
        Logger.log("Colonne de présence totale trouvée à l'index " + i);
      }
    }
    
    // 4. Initialiser les structures pour les résultats
    var stats = {
      total: { duree: 0, nbEntrees: 0, presenceTotal: 0 },
      parType: {},
      parPoste: {},
      parEquipe: {},
      parOperateur: {},
      donnees: [] // Pour stocker les données détaillées pour les graphiques temporels
    };
    
    // Initialiser les compteurs pour les types d'inactivité
    inactiviteTypes.forEach(function(type) {
      stats.parType[type] = { duree: 0, nbOccurrences: 0 };
    });
    
    // 5. Agréger les données d'inactivité à partir des 3ème ligne (index 2)
    for (var rowIndex = 2; rowIndex < values.length; rowIndex++) {
      var row = values[rowIndex];
      
      // Vérifier si la ligne a un opérateur (pour être sûr que c'est une ligne valide)
      if (!row[operateurIndex]) continue;
      
      // Appliquer les filtres
      if (filtres) {
        // Filtre de date
        if ((filtres.dateDebut || filtres.dateFin) && dateIndex >= 0) {
          var dateVal = row[dateIndex];
          if (!(dateVal instanceof Date)) continue;
          
          if (filtres.dateDebut) {
            var dateDebut = new Date(filtres.dateDebut);
            if (dateVal < dateDebut) continue;
          }
          
          if (filtres.dateFin) {
            var dateFin = new Date(filtres.dateFin);
            dateFin.setHours(23, 59, 59, 999); // Fin de journée
            if (dateVal > dateFin) continue;
          }
        }
        
        // Filtre d'équipe
        if (filtres.equipe && equipeIndex >= 0 && row[equipeIndex] !== filtres.equipe) {
          continue;
        }
        
        // Filtre de poste (approximatif, car nous n'avons que le poste habituel)
        if (filtres.poste && posteIndex >= 0 && row[posteIndex] !== filtres.poste) {
          continue;
        }
      }
      
      // Compte cette ligne comme une entrée valide
      stats.total.nbEntrees++;
      
      // Ajouter la présence totale si disponible
      if (presenceIndex >= 0) {
        var presenceVal = parseFloat(row[presenceIndex]) || 0;
        stats.total.presenceTotal += presenceVal;
      }
      
      // Calculer l'inactivité totale pour cette ligne
      var rowInactiviteTotal = 0;
      var rowInactiviteParType = {};
      
      // Parcourir chaque colonne d'inactivité
      for (var i = 0; i < inactiviteColumns.length; i++) {
        var colIndex = inactiviteColumns[i];
        var type = inactiviteTypes[i];
        
        var cellValue = row[colIndex];
        var numValue = 0;
        
        // Convertir la valeur en nombre avec gestion explicite des types
        if (typeof cellValue === 'number') {
          numValue = cellValue;
        } else if (typeof cellValue === 'string') {
          // Remplacer la virgule par un point si nécessaire (format français)
          cellValue = cellValue.replace(',', '.');
          numValue = parseFloat(cellValue) || 0;
        }
        
        // Seulement compter si > 0
        if (numValue > 0) {
          stats.parType[type].duree += numValue;
          stats.parType[type].nbOccurrences++;
          rowInactiviteTotal += numValue;
          rowInactiviteParType[type] = numValue;
        }
      }
      
      // Si cette ligne a de l'inactivité, ajouter aux statistiques
      if (rowInactiviteTotal > 0) {
        // Ajouter au total général
        stats.total.duree += rowInactiviteTotal;
        
        // Ajouter aux statistiques par opérateur
        var operateur = String(row[operateurIndex]);
        if (!stats.parOperateur[operateur]) {
          stats.parOperateur[operateur] = { duree: 0, nbEntrees: 0, parType: {} };
          inactiviteTypes.forEach(function(type) {
            stats.parOperateur[operateur].parType[type] = 0;
          });
        }
        
        stats.parOperateur[operateur].duree += rowInactiviteTotal;
        stats.parOperateur[operateur].nbEntrees++;
        
        // Ajouter les détails par type
        Object.keys(rowInactiviteParType).forEach(function(type) {
          stats.parOperateur[operateur].parType[type] += rowInactiviteParType[type];
        });
        
        // Ajouter aux statistiques par poste
        if (posteIndex >= 0 && row[posteIndex]) {
          var poste = String(row[posteIndex]);
          
          if (!stats.parPoste[poste]) {
            stats.parPoste[poste] = { duree: 0, nbEntrees: 0 };
          }
          
          stats.parPoste[poste].duree += rowInactiviteTotal;
          stats.parPoste[poste].nbEntrees++;
        }
        
        // Ajouter aux statistiques par équipe
        if (equipeIndex >= 0 && row[equipeIndex]) {
          var equipe = String(row[equipeIndex]);
          
          if (!stats.parEquipe[equipe]) {
            stats.parEquipe[equipe] = { duree: 0, nbEntrees: 0, nbOperateurs: 0, operateurs: new Set() };
          }
          
          stats.parEquipe[equipe].duree += rowInactiviteTotal;
          stats.parEquipe[equipe].nbEntrees++;
          stats.parEquipe[equipe].operateurs.add(operateur);
        }
        
        // Ajouter aux données détaillées pour les graphiques temporels
        if (dateIndex >= 0 && row[dateIndex] instanceof Date) {
          var date = row[dateIndex];
          var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
          
          stats.donnees.push({
            date: dateStr,
            operateur: operateur,
            poste: posteIndex >= 0 ? row[posteIndex] : "",
            equipe: equipeIndex >= 0 ? row[equipeIndex] : "",
            inactiviteTotal: rowInactiviteTotal,
            inactiviteParType: rowInactiviteParType
          });
        }
      }
    }
    
    // Convertir les ensembles en nombres pour la sérialisation JSON
    Object.keys(stats.parEquipe).forEach(function(equipe) {
      stats.parEquipe[equipe].nbOperateurs = stats.parEquipe[equipe].operateurs.size;
      delete stats.parEquipe[equipe].operateurs;
    });
    
    // Vérification des résultats
    Logger.log("Statistiques d'inactivité calculées. Total: " + stats.total.duree + "h, Entrées: " + stats.total.nbEntrees);
    
    return { 
      success: true, 
      stats: stats,
      debug: {
        inactiviteColumns: inactiviteColumns,
        inactiviteTypes: inactiviteTypes
      }
    };

  } catch (e) {
    Logger.log("ERREUR dans obtenirStatistiquesInactivite: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    return { success: false, message: "Erreur serveur: " + e.toString() };
  }
}

// Fonction pour obtenir des statistiques par équipe
function obtenirStatistiquesParEquipe(filtres) {
  try {
    Logger.log("Démarrage de obtenirStatistiquesParEquipe() avec filtres: " + JSON.stringify(filtres));
    
    // Récupérer les données, filtrées si des filtres sont fournis
    var donnees;
    if (filtres && (filtres.dateDebut || filtres.dateFin || filtres.rendementMin !== null || filtres.rendementMax !== null || filtres.equipe || filtres.poste)) {
      donnees = obtenirDonneesFiltrees(filtres);
    } else {
      donnees = obtenirDonnees();
    }
    
    if (!donnees || donnees.erreur) {
      Logger.log("Erreur: " + (donnees ? donnees.erreur : "Aucune donnée reçue"));
      return [];
    }
    
    var resultats = {};
    
    // Regrouper par équipe
    donnees.donnees.forEach(function(item) {
      if (!item.equipe) return;
      
      if (!resultats[item.equipe]) {
        resultats[item.equipe] = {
          taiTotal: 0,
          presenceTotal: 0,
          nbEntrees: 0,
          taiParPoste: {},
          presenceParPoste: {}
        };
        
        // Initialiser les compteurs pour chaque type de poste
        var postesTypes = donnees.postesTypes || [];
        postesTypes.forEach(function(posteType) {
          resultats[item.equipe].taiParPoste[posteType] = 0;
          resultats[item.equipe].presenceParPoste[posteType] = 0;
        });
      }
      
      // Incrémenter le compteur d'entrées
      resultats[item.equipe].nbEntrees++;
      
      // Ajouter les TAI et présences
      resultats[item.equipe].taiTotal += item.tai;
      resultats[item.equipe].presenceTotal += item.presence;
      
      // Ajouter les TAI et présences par type de poste
      if (item.taiParPoste) {
        Object.keys(item.taiParPoste).forEach(function(posteType) {
          resultats[item.equipe].taiParPoste[posteType] += item.taiParPoste[posteType];
        });
      }
      
      if (item.presenceParPoste) {
        Object.keys(item.presenceParPoste).forEach(function(posteType) {
          resultats[item.equipe].presenceParPoste[posteType] += item.presenceParPoste[posteType];
        });
      }
    });
    
    // Calculer les statistiques
    var statistiques = [];
    Object.keys(resultats).forEach(function(equipe) {
      var result = resultats[equipe];
      
      // Calculer le rendement moyen (TAI total / Présence totale)
      var rendementMoyen = result.presenceTotal > 0 ? result.taiTotal / result.presenceTotal : 0;
      
      // Calculer les rendements par type de poste
      var rendementParPoste = {};
      Object.keys(result.taiParPoste).forEach(function(posteType) {
        var taiPoste = result.taiParPoste[posteType];
        var presencePoste = result.presenceParPoste[posteType];
        rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
      });
      
      statistiques.push({
        equipe: equipe,
        rendementMoyen: rendementMoyen,
        taiTotal: result.taiTotal,
        presenceTotal: result.presenceTotal,
        nbEntrees: result.nbEntrees,
        taiParPoste: result.taiParPoste,
        presenceParPoste: result.presenceParPoste,
        rendementParPoste: rendementParPoste
      });
    });
    
    // Trier par rendement décroissant
    statistiques.sort(function(a, b) {
      return b.rendementMoyen - a.rendementMoyen;
    });
    
    return statistiques;
  } catch (e) {
    Logger.log("ERREUR dans obtenirStatistiquesParEquipe: " + e.toString());
    return [];
  }
}


// Fonction pour obtenir les données temporelles de rendement
function obtenirDonneesTemporelles(filtres) {
  try {
    Logger.log("Démarrage de obtenirDonneesTemporelles() avec filtres: " + JSON.stringify(filtres));
    
    // Récupérer les données, filtrées si des filtres sont fournis
    var donnees;
    if (filtres && (filtres.dateDebut || filtres.dateFin || filtres.rendementMin !== null || filtres.rendementMax !== null || filtres.equipe || filtres.poste)) {
      donnees = obtenirDonneesFiltrees(filtres);
    } else {
      donnees = obtenirDonnees();
    }
    
    if (!donnees || donnees.erreur) {
      Logger.log("Erreur: " + (donnees ? donnees.erreur : "Aucune donnée reçue"));
      return {
        annees: [],
        mois: [],
        semaines: []
      };
    }
    
    var postesTypes = donnees.postesTypes || [];
    var resultatsParAnnee = {};
    var resultatsParMois = {};
    var resultatsParSemaine = {};
    
    // Regrouper par période
    donnees.donnees.forEach(function(item) {
      if (!item.annee || !item.date) return;
      
      // Pour l'année
      var annee = String(item.annee);
      if (!resultatsParAnnee[annee]) {
        resultatsParAnnee[annee] = {
          taiTotal: 0,
          presenceTotal: 0,
          nbEntrees: 0,
          taiParPoste: {},
          presenceParPoste: {}
        };
        
        // Initialiser les compteurs pour chaque type de poste
        postesTypes.forEach(function(posteType) {
          resultatsParAnnee[annee].taiParPoste[posteType] = 0;
          resultatsParAnnee[annee].presenceParPoste[posteType] = 0;
        });
      }
      
      // Incrémenter le compteur d'entrées
      resultatsParAnnee[annee].nbEntrees++;
      
      // Ajouter les TAI et présences
      resultatsParAnnee[annee].taiTotal += item.tai;
      resultatsParAnnee[annee].presenceTotal += item.presence;
      
      // Ajouter les TAI et présences par type de poste
      if (item.taiParPoste) {
        Object.keys(item.taiParPoste).forEach(function(posteType) {
          if (resultatsParAnnee[annee].taiParPoste.hasOwnProperty(posteType)) {
            resultatsParAnnee[annee].taiParPoste[posteType] += item.taiParPoste[posteType];
          }
        });
      }
      
      if (item.presenceParPoste) {
        Object.keys(item.presenceParPoste).forEach(function(posteType) {
          if (resultatsParAnnee[annee].presenceParPoste.hasOwnProperty(posteType)) {
            resultatsParAnnee[annee].presenceParPoste[posteType] += item.presenceParPoste[posteType];
          }
        });
      }
      
      // Pour le mois
      var dateParts = item.date.split('-');
      if (dateParts.length >= 2) {
        var mois = annee + '-' + dateParts[1];
        if (!resultatsParMois[mois]) {
          resultatsParMois[mois] = {
            taiTotal: 0,
            presenceTotal: 0,
            nbEntrees: 0,
            taiParPoste: {},
            presenceParPoste: {}
          };
          
          // Initialiser les compteurs pour chaque type de poste
          postesTypes.forEach(function(posteType) {
            resultatsParMois[mois].taiParPoste[posteType] = 0;
            resultatsParMois[mois].presenceParPoste[posteType] = 0;
          });
        }
        
        // Incrémenter le compteur d'entrées
        resultatsParMois[mois].nbEntrees++;
        
        // Ajouter les TAI et présences
        resultatsParMois[mois].taiTotal += item.tai;
        resultatsParMois[mois].presenceTotal += item.presence;
        
        // Ajouter les TAI et présences par type de poste
        if (item.taiParPoste) {
          Object.keys(item.taiParPoste).forEach(function(posteType) {
            if (resultatsParMois[mois].taiParPoste.hasOwnProperty(posteType)) {
              resultatsParMois[mois].taiParPoste[posteType] += item.taiParPoste[posteType];
            }
          });
        }
        
        if (item.presenceParPoste) {
          Object.keys(item.presenceParPoste).forEach(function(posteType) {
            if (resultatsParMois[mois].presenceParPoste.hasOwnProperty(posteType)) {
              resultatsParMois[mois].presenceParPoste[posteType] += item.presenceParPoste[posteType];
            }
          });
        }
      }
      
      // Pour la semaine
      if (item.semaine) {
        var semaine = annee + '-S' + String(item.semaine).padStart(2, '0');
        if (!resultatsParSemaine[semaine]) {
          resultatsParSemaine[semaine] = {
            taiTotal: 0,
            presenceTotal: 0,
            nbEntrees: 0,
            taiParPoste: {},
            presenceParPoste: {}
          };
          
          // Initialiser les compteurs pour chaque type de poste
          postesTypes.forEach(function(posteType) {
            resultatsParSemaine[semaine].taiParPoste[posteType] = 0;
            resultatsParSemaine[semaine].presenceParPoste[posteType] = 0;
          });
        }
        
        // Incrémenter le compteur d'entrées
        resultatsParSemaine[semaine].nbEntrees++;
        
        // Ajouter les TAI et présences
        resultatsParSemaine[semaine].taiTotal += item.tai;
        resultatsParSemaine[semaine].presenceTotal += item.presence;
        
        // Ajouter les TAI et présences par type de poste
        if (item.taiParPoste) {
          Object.keys(item.taiParPoste).forEach(function(posteType) {
            if (resultatsParSemaine[semaine].taiParPoste.hasOwnProperty(posteType)) {
              resultatsParSemaine[semaine].taiParPoste[posteType] += item.taiParPoste[posteType];
            }
          });
        }
        
        if (item.presenceParPoste) {
          Object.keys(item.presenceParPoste).forEach(function(posteType) {
            if (resultatsParSemaine[semaine].presenceParPoste.hasOwnProperty(posteType)) {
              resultatsParSemaine[semaine].presenceParPoste[posteType] += item.presenceParPoste[posteType];
            }
          });
        }
      }
    });
    
    // Calculer les statistiques par année
    var statsAnnee = [];
    Object.keys(resultatsParAnnee).sort().forEach(function(annee) {
      var result = resultatsParAnnee[annee];
      
      // Calculer le rendement (TAI total / Présence totale)
      var rendement = result.presenceTotal > 0 ? result.taiTotal / result.presenceTotal : 0;
      
      // Calculer les rendements par type de poste
      var rendementParPoste = {};
      Object.keys(result.taiParPoste).forEach(function(posteType) {
        var taiPoste = result.taiParPoste[posteType];
        var presencePoste = result.presenceParPoste[posteType];
        rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
      });
      
      statsAnnee.push({
        periode: annee,
        rendement: rendement,
        taiTotal: result.taiTotal,
        presenceTotal: result.presenceTotal,
        nbEntrees: result.nbEntrees,
        taiParPoste: result.taiParPoste,
        presenceParPoste: result.presenceParPoste,
        rendementParPoste: rendementParPoste
      });
    });
    
    // Calculer les statistiques par mois
    var statsMois = [];
    Object.keys(resultatsParMois).sort().forEach(function(mois) {
      var result = resultatsParMois[mois];
      
      // Calculer le rendement (TAI total / Présence totale)
      var rendement = result.presenceTotal > 0 ? result.taiTotal / result.presenceTotal : 0;
      
      // Calculer les rendements par type de poste
      var rendementParPoste = {};
      Object.keys(result.taiParPoste).forEach(function(posteType) {
        var taiPoste = result.taiParPoste[posteType];
        var presencePoste = result.presenceParPoste[posteType];
        rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
      });
      
      statsMois.push({
        periode: mois,
        rendement: rendement,
        taiTotal: result.taiTotal,
        presenceTotal: result.presenceTotal,
        nbEntrees: result.nbEntrees,
        taiParPoste: result.taiParPoste,
        presenceParPoste: result.presenceParPoste,
        rendementParPoste: rendementParPoste
      });
    });
    
    // Calculer les statistiques par semaine
    var statsSemaine = [];
    Object.keys(resultatsParSemaine).sort().forEach(function(semaine) {
      var result = resultatsParSemaine[semaine];
      
      // Calculer le rendement (TAI total / Présence totale)
      var rendement = result.presenceTotal > 0 ? result.taiTotal / result.presenceTotal : 0;
      
      // Calculer les rendements par type de poste
      var rendementParPoste = {};
      Object.keys(result.taiParPoste).forEach(function(posteType) {
        var taiPoste = result.taiParPoste[posteType];
        var presencePoste = result.presenceParPoste[posteType];
        rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
      });
      
      statsSemaine.push({
        periode: semaine,
        rendement: rendement,
        taiTotal: result.taiTotal,
        presenceTotal: result.presenceTotal,
        nbEntrees: result.nbEntrees,
        taiParPoste: result.taiParPoste,
        presenceParPoste: result.presenceParPoste,
        rendementParPoste: rendementParPoste
      });
    });
    
    return {
      annees: statsAnnee,
      mois: statsMois,
      semaines: statsSemaine
    };
  } catch (e) {
    Logger.log("ERREUR dans obtenirDonneesTemporelles: " + e.toString());
    return {
      annees: [],
      mois: [],
      semaines: []
    };
  }
}

function getSAParOperateur() {
  try {
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Effectif');
    
    if (!sheet) {
      Logger.log("Feuille 'Aleas' introuvable pour récupérer les SA");
      return {};
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // Les en-têtes sont en première ligne
    
    // Identifier les indices des colonnes
    var operateurIndex = -1;
    var saIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase().trim();
      if (header === "opérateur" || header === "operateur") {
        operateurIndex = i;
      } else if (header === "sa") {
        saIndex = i;
      }
    }
    
    // Vérifier que les colonnes essentielles ont été trouvées
    if (operateurIndex === -1 || saIndex === -1) {
      Logger.log("Colonnes Opérateur ou SA introuvables dans la feuille Aleas");
      return {};
    }
    
    // Créer un dictionnaire operateur -> SA
    var saParOperateur = {};
    
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      var operateur = row[operateurIndex];
      var sa = row[saIndex];
      
      if (operateur && sa) {
        saParOperateur[operateur] = sa;
      }
    }
    
    return saParOperateur;
  } catch (e) {
    Logger.log("ERREUR dans getSAParOperateur: " + e.toString());
    return {};
  }
}

// Fonction pour obtenir les performances des opérateurs
function obtenirPerformancesOperateurs(filtres) {
  try {
    Logger.log("Démarrage de obtenirPerformancesOperateurs() avec filtres: " + JSON.stringify(filtres));
    
    // Récupérer les données, filtrées si des filtres sont fournis
    var donnees;
    if (filtres && (filtres.dateDebut || filtres.dateFin || filtres.rendementMin !== null || filtres.rendementMax !== null || filtres.equipe || filtres.poste)) {
      donnees = obtenirDonneesFiltrees(filtres);
    } else {
      donnees = obtenirDonnees();
    }
    
    if (!donnees || donnees.erreur) {
      Logger.log("Erreur: " + (donnees ? donnees.erreur : "Aucune donnée reçue"));
      return [];
    }
    
    var resultats = {};
    
    // Récupérer les SA des opérateurs depuis la feuille Aleas
    var saParOperateur = getSAParOperateur();
    
    // Regrouper par opérateur
    donnees.donnees.forEach(function(item) {
      if (!item.operateur) return;
      
      if (!resultats[item.operateur]) {
        resultats[item.operateur] = {
          taiTotal: 0,
          presenceTotal: 0,
          nbEntrees: 0,
          taiParPoste: {},
          presenceParPoste: {},
          masquages: [],
          demasquages: [],
          sa: saParOperateur[item.operateur] || "" // Ajouter le SA de l'opérateur
        };
        
        // Initialiser les compteurs pour chaque type de poste
        var postesTypes = donnees.postesTypes || [];
        postesTypes.forEach(function(posteType) {
          resultats[item.operateur].taiParPoste[posteType] = 0;
          resultats[item.operateur].presenceParPoste[posteType] = 0;
        });
      }
      
      // Incrémenter le compteur d'entrées
      resultats[item.operateur].nbEntrees++;
      
      // Ajouter les TAI et présences
      resultats[item.operateur].taiTotal += item.tai;
      resultats[item.operateur].presenceTotal += item.presence;
      
      // Ajouter les valeurs de MASQUAGE et DEMASQUAGE si elles existent
      if ('masquage' in item && !isNaN(item.masquage)) {
        resultats[item.operateur].masquages.push(item.masquage);
      }
      
      if ('demasquage' in item && !isNaN(item.demasquage)) {
        resultats[item.operateur].demasquages.push(item.demasquage);
      }
      
      // Ajouter les TAI et présences par type de poste
      if (item.taiParPoste) {
        Object.keys(item.taiParPoste).forEach(function(posteType) {
          if (resultats[item.operateur].taiParPoste.hasOwnProperty(posteType)) {
            resultats[item.operateur].taiParPoste[posteType] += item.taiParPoste[posteType];
          }
        });
      }
      
      if (item.presenceParPoste) {
        Object.keys(item.presenceParPoste).forEach(function(posteType) {
          if (resultats[item.operateur].presenceParPoste.hasOwnProperty(posteType)) {
            resultats[item.operateur].presenceParPoste[posteType] += item.presenceParPoste[posteType];
          }
        });
      }
    });
    
    // Calculer les statistiques
    var statistiques = [];
    Object.keys(resultats).forEach(function(operateur) {
      var result = resultats[operateur];
      
      // Calculer le rendement moyen (TAI total / Présence totale)
      var rendementMoyen = result.presenceTotal > 0 ? result.taiTotal / result.presenceTotal : 0;
      
      // Calculer les moyennes MASQUAGE et DEMASQUAGE
      var masquageMoyen = calculerMoyenne(result.masquages);
      var demasquageMoyen = calculerMoyenne(result.demasquages);
      
      // Calculer les rendements par type de poste
      var rendementParPoste = {};
      Object.keys(result.taiParPoste).forEach(function(posteType) {
        var taiPoste = result.taiParPoste[posteType];
        var presencePoste = result.presenceParPoste[posteType];
        rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
      });
      
      // Créer l'objet de statistiques
      var statObj = {
        operateur: operateur,
        sa: result.sa, // Inclure le SA dans les statistiques
        rendementMoyen: rendementMoyen,
        taiTotal: result.taiTotal,
        presenceTotal: result.presenceTotal,
        nbEntrees: result.nbEntrees,
        taiParPoste: result.taiParPoste,
        presenceParPoste: result.presenceParPoste,
        rendementParPoste: rendementParPoste
      };
      
      // Ajouter MASQUAGE et DEMASQUAGE uniquement s'ils ont des valeurs
      if (result.masquages.length > 0) {
        statObj.masquageMoyen = masquageMoyen;
      }
      
      if (result.demasquages.length > 0) {
        statObj.demasquageMoyen = demasquageMoyen;
      }
      
      statistiques.push(statObj);
    });
    
    // Trier par rendement décroissant
    statistiques.sort(function(a, b) {
      return b.rendementMoyen - a.rendementMoyen;
    });
    
    return statistiques;
  } catch (e) {
    Logger.log("ERREUR dans obtenirPerformancesOperateurs: " + e.toString());
    return [];
  }
}



// Fonction pour obtenir les aléas d'un opérateur
function obtenirAleasOperateur(nomOperateur, filtres) {
  try {
    Logger.log("Récupération des aléas pour l'opérateur: " + nomOperateur);
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      Logger.log("Feuille 'Aleas' introuvable");
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // Les en-têtes sont en première ligne
    
    Logger.log("En-têtes trouvés: " + JSON.stringify(headers));
    
    // Identifier les indices des colonnes (recherche insensible à la casse)
    var operateurIndex = -1;
    var saIndex = -1; 
    var equipeIndex = -1;
    var posteIndex = -1;
    var dateIndex = -1;
    var typeAleaIndex = -1;
    var commentaireIndex = -1;
    var dureeIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase().trim();
      if (header === "opérateur" || header === "operateur") {
        operateurIndex = i;
      } else if (header === "sa") {
        saIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      } else if (header === "poste") {
        posteIndex = i;
      } else if (header === "date") {
        dateIndex = i;
      } else if (header === "type aléa" || header === "type alea" || header === "type") {
        typeAleaIndex = i;
      } else if (header === "commentaire" || header === "commentaires") {
        commentaireIndex = i;
      } else if (header === "durée" || header === "duree") {
        dureeIndex = i;
      }
    }
    
    Logger.log("Indices des colonnes - Opérateur: " + operateurIndex + 
               ", SA: " + saIndex + 
               ", Type Aléa: " + typeAleaIndex + 
               ", Durée: " + dureeIndex);
    
    // Vérifier que les colonnes essentielles ont été trouvées
    if (operateurIndex === -1) {
      Logger.log("Colonne 'Opérateur' introuvable dans la feuille Aleas");
      return { success: false, message: "Structure de la feuille Aleas incorrecte" };
    }
    
    // Filtrer les aléas par opérateur
    var aleas = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (row[operateurIndex] === nomOperateur) {
        var date = row[dateIndex];
        var dateFormatee = "";
        
        // Formater la date correctement
        if (date instanceof Date) {
          dateFormatee = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (typeof date === 'string') {
          dateFormatee = date;
        }
        
        var aleaItem = {
          operateur: row[operateurIndex],
          sa: saIndex >= 0 ? row[saIndex] : "",
          equipe: equipeIndex >= 0 ? row[equipeIndex] : "",
          poste: posteIndex >= 0 ? row[posteIndex] : "",
          date: dateFormatee,
          typeAlea: typeAleaIndex >= 0 ? (row[typeAleaIndex] || "") : "",
          commentaire: commentaireIndex >= 0 ? (row[commentaireIndex] || "") : "",
          duree: dureeIndex >= 0 ? (row[dureeIndex] !== null && row[dureeIndex] !== undefined ? row[dureeIndex] : "") : ""
        };
        
        Logger.log("Aléa trouvé: " + JSON.stringify(aleaItem));
        aleas.push(aleaItem);
      }
    }
    
    // Appliquer les filtres si nécessaire
    if (filtres) {
      // Filtrer par date
      if (filtres.dateDebut || filtres.dateFin) {
        aleas = aleas.filter(function(alea) {
          var dateParts = alea.date.split('/');
          if (dateParts.length !== 3) return true; // Si format invalide, inclure par défaut
          
          var aleaDateStr = dateParts[2] + '-' + dateParts[1] + '-' + dateParts[0];
          var aleaDate = new Date(aleaDateStr);
          
          if (filtres.dateDebut) {
            var dateDebut = new Date(filtres.dateDebut);
            if (aleaDate < dateDebut) return false;
          }
          
          if (filtres.dateFin) {
            var dateFin = new Date(filtres.dateFin);
            // Définir l'heure à la fin de la journée pour inclure toute la journée
            dateFin.setHours(23, 59, 59, 999);
            if (aleaDate > dateFin) return false;
          }
          
          return true;
        });
      }
      
      // Filtrer par poste
      if (filtres.poste) {
        aleas = aleas.filter(function(alea) {
          return alea.poste === filtres.poste;
        });
      }
      
      // Filtrer par équipe
      if (filtres.equipe) {
        aleas = aleas.filter(function(alea) {
          return alea.equipe === filtres.equipe;
        });
      }
    }
    
    // Trier par date (plus récent en premier)
    aleas.sort(function(a, b) {
      var dateA = convertirEnDate(a.date);
      var dateB = convertirEnDate(b.date);
      return dateB - dateA;
    });
    
    return {
      success: true,
      aleas: aleas
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirAleasOperateur: " + e.toString());
    return {
      success: false, 
      message: "Erreur lors de la récupération des aléas: " + e.toString()
    };
  }
}

// Fonction pour obtenir tous les aléas (tous opérateurs)
function obtenirTousAleas(filtres) {
  try {
    Logger.log("Récupération de tous les aléas");
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Aleas');
    
    if (!sheet) {
      Logger.log("Feuille 'Aleas' introuvable");
      return { success: false, message: "Feuille 'Aleas' introuvable" };
    }
    
    // Récupérer les données
    var range = sheet.getDataRange();
    var values = range.getValues();
    var headers = values[0]; // Les en-têtes sont en première ligne
    
    // Identifier les indices des colonnes (recherche insensible à la casse)
    var operateurIndex = -1;
    var saIndex = -1; 
    var equipeIndex = -1;
    var posteIndex = -1;
    var dateIndex = -1;
    var typeAleaIndex = -1;
    var commentaireIndex = -1;
    var dureeIndex = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "").toLowerCase().trim();
      if (header === "opérateur" || header === "operateur") {
        operateurIndex = i;
      } else if (header === "sa") {
        saIndex = i;
      } else if (header === "equipe" || header === "équipe") {
        equipeIndex = i;
      } else if (header === "poste") {
        posteIndex = i;
      } else if (header === "date") {
        dateIndex = i;
      } else if (header === "type aléa" || header === "type alea" || header === "type") {
        typeAleaIndex = i;
      } else if (header === "commentaire" || header === "commentaires") {
        commentaireIndex = i;
      } else if (header === "durée" || header === "duree") {
        dureeIndex = i;
      }
    }
    
    // Vérifier que les colonnes essentielles ont été trouvées
    if (operateurIndex === -1 || dateIndex === -1) {
      Logger.log("Colonnes essentielles introuvables dans la feuille Aleas");
      return { success: false, message: "Structure de la feuille Aleas incorrecte" };
    }
    
    // Récupérer tous les aléas
    var aleas = [];
    for (var i = 1; i < values.length; i++) {
      var row = values[i];
      if (row[operateurIndex]) { // S'assurer qu'il y a un opérateur
        var date = row[dateIndex];
        var dateFormatee = "";
        
        // Formater la date correctement
        if (date instanceof Date) {
          dateFormatee = Utilities.formatDate(date, Session.getScriptTimeZone(), "dd/MM/yyyy");
        } else if (typeof date === 'string') {
          dateFormatee = date;
        }
        
        var aleaItem = {
          operateur: row[operateurIndex],
          sa: saIndex >= 0 ? row[saIndex] : "",
          equipe: equipeIndex >= 0 ? row[equipeIndex] : "",
          poste: posteIndex >= 0 ? row[posteIndex] : "",
          date: dateFormatee,
          typeAlea: typeAleaIndex >= 0 ? (row[typeAleaIndex] || "") : "",
          commentaire: commentaireIndex >= 0 ? (row[commentaireIndex] || "") : "",
          duree: dureeIndex >= 0 ? (row[dureeIndex] !== null && row[dureeIndex] !== undefined ? row[dureeIndex] : "") : ""
        };
        
        aleas.push(aleaItem);
      }
    }
    
    // Appliquer les filtres si nécessaire
    if (filtres) {
      // Filtrer par date
      if (filtres.dateDebut || filtres.dateFin) {
        aleas = aleas.filter(function(alea) {
          var dateParts = alea.date.split('/');
          if (dateParts.length !== 3) return true;
          
          var aleaDateStr = dateParts[2] + '-' + dateParts[1] + '-' + dateParts[0];
          var aleaDate = new Date(aleaDateStr);
          
          if (filtres.dateDebut) {
            var dateDebut = new Date(filtres.dateDebut);
            if (aleaDate < dateDebut) return false;
          }
          
          if (filtres.dateFin) {
            var dateFin = new Date(filtres.dateFin);
            // Définir l'heure à la fin de la journée pour inclure toute la journée
            dateFin.setHours(23, 59, 59, 999);
            if (aleaDate > dateFin) return false;
          }
          
          return true;
        });
      }
      
      // Filtrer par poste
      if (filtres.poste) {
        aleas = aleas.filter(function(alea) {
          return alea.poste === filtres.poste;
        });
      }
      
      // Filtrer par équipe
      if (filtres.equipe) {
        aleas = aleas.filter(function(alea) {
          return alea.equipe === filtres.equipe;
        });
      }
      
      // Filtrer par opérateur si spécifié
      if (filtres.operateur) {
        aleas = aleas.filter(function(alea) {
          return alea.operateur === filtres.operateur;
        });
      }
    }
    
    // Trier par date (plus récent en premier)
    aleas.sort(function(a, b) {
      var dateA = convertirEnDate(a.date);
      var dateB = convertirEnDate(b.date);
      return dateB - dateA;
    });
    
    return {
      success: true,
      aleas: aleas
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirTousAleas: " + e.toString());
    return {
      success: false, 
      message: "Erreur lors de la récupération des aléas: " + e.toString()
    };
  }
}

// Fonction utilitaire pour convertir une date au format DD/MM/YYYY en objet Date
function convertirEnDate(dateStr) {
  if (!dateStr) return new Date(0);
  
  var parts = dateStr.split('/');
  if (parts.length !== 3) return new Date(0);
  
  // Convertir au format YYYY-MM-DD pour le constructeur Date
  return new Date(parts[2] + '-' + parts[1] + '-' + parts[0]);
}
// Fonction pour obtenir les données d'un opérateur spécifique
// Fonction pour obtenir les données d'un opérateur spécifique
function obtenirDonneesOperateur(nomOperateur, filtres) {
  try {
    Logger.log("Démarrage de obtenirDonneesOperateur() pour " + nomOperateur + " avec filtres: " + JSON.stringify(filtres));
    
    // Récupérer les données, filtrées si des filtres sont fournis
    var donnees;
    if (filtres && (filtres.dateDebut || filtres.dateFin || filtres.rendementMin !== null || filtres.rendementMax !== null || filtres.equipe || filtres.poste)) {
      donnees = obtenirDonneesFiltrees(filtres);
    } else {
      donnees = obtenirDonnees();
    }
    
    if (!donnees || donnees.erreur) {
      Logger.log("Erreur: " + (donnees ? donnees.erreur : "Aucune donnée reçue"));
      return {
        operateur: nomOperateur,
        rendements: [],
        statsPeriodes: [],
        statsPostes: [],
        statsEquipes: []
      };
    }
    
    // Récupérer le SA de l'opérateur
    var saParOperateur = getSAParOperateur();
    var saOperateur = saParOperateur[nomOperateur] || "";
    
    var postesTypes = donnees.postesTypes || [];
    
    var resultats = {
      rendements: [],
      postes: {},
      equipes: {},
      periodes: {},
      taiTotal: 0,
      presenceTotal: 0,
      taiParPoste: {},
      presenceParPoste: {}
    };
    
    // Initialiser les compteurs pour chaque type de poste
    postesTypes.forEach(function(posteType) {
      resultats.taiParPoste[posteType] = 0;
      resultats.presenceParPoste[posteType] = 0;
    });
    
    // Filtrer par opérateur
    donnees.donnees.forEach(function(item) {
      if (item.operateur !== nomOperateur) return;
      
      // Ajouter aux totaux de l'opérateur
      resultats.taiTotal += item.tai;
      resultats.presenceTotal += item.presence;
      
      // Ajouter les TAI et présences par type de poste
      if (item.taiParPoste) {
        Object.keys(item.taiParPoste).forEach(function(posteType) {
          if (resultats.taiParPoste.hasOwnProperty(posteType)) {
            resultats.taiParPoste[posteType] += item.taiParPoste[posteType];
          }
        });
      }
      
      if (item.presenceParPoste) {
        Object.keys(item.presenceParPoste).forEach(function(posteType) {
          if (resultats.presenceParPoste.hasOwnProperty(posteType)) {
            resultats.presenceParPoste[posteType] += item.presenceParPoste[posteType];
          }
        });
      }
      
      // Calculer le rendement si non défini
      var rendement = item.rendement;
      if (rendement === undefined && item.presence > 0) {
        rendement = item.tai / item.presence;
      }
      
      // Ajouter les données de rendement avec la date
      if (item.date) {
        resultats.rendements.push({
          date: item.date,
          rendement: rendement,
          tai: item.tai,
          presence: item.presence,
          poste: item.poste,
          equipe: item.equipe,
          taiParPoste: item.taiParPoste,
          presenceParPoste: item.presenceParPoste
        });
      }
      
      // Regrouper par équipe
      if (item.equipe) {
        if (!resultats.equipes[item.equipe]) {
          resultats.equipes[item.equipe] = {
            taiTotal: 0,
            presenceTotal: 0,
            nbEntrees: 0,
            taiParPoste: {},
            presenceParPoste: {}
          };
          
          // Initialiser les compteurs pour chaque type de poste
          postesTypes.forEach(function(posteType) {
            resultats.equipes[item.equipe].taiParPoste[posteType] = 0;
            resultats.equipes[item.equipe].presenceParPoste[posteType] = 0;
          });
        }
        
        // Incrémenter le compteur d'entrées
        resultats.equipes[item.equipe].nbEntrees++;
        
        // Ajouter les TAI et présences
        resultats.equipes[item.equipe].taiTotal += item.tai;
        resultats.equipes[item.equipe].presenceTotal += item.presence;
        
        // Ajouter les TAI et présences par type de poste
        if (item.taiParPoste) {
          Object.keys(item.taiParPoste).forEach(function(posteType) {
            if (resultats.equipes[item.equipe].taiParPoste.hasOwnProperty(posteType)) {
              resultats.equipes[item.equipe].taiParPoste[posteType] += item.taiParPoste[posteType];
            }
          });
        }
        
        if (item.presenceParPoste) {
          Object.keys(item.presenceParPoste).forEach(function(posteType) {
            if (resultats.equipes[item.equipe].presenceParPoste.hasOwnProperty(posteType)) {
              resultats.equipes[item.equipe].presenceParPoste[posteType] += item.presenceParPoste[posteType];
            }
          });
        }
      }
      
      // Regrouper par poste (utiliser les types de postes)
      postesTypes.forEach(function(posteType) {
        if (!resultats.postes[posteType]) {
          resultats.postes[posteType] = {
            taiTotal: 0,
            presenceTotal: 0,
            nbEntrees: 0
          };
        }
        
        // Si ce poste a du TAI ou de la présence, l'ajouter aux statistiques
        if (item.taiParPoste && item.taiParPoste[posteType] > 0) {
          resultats.postes[posteType].taiTotal += item.taiParPoste[posteType];
          resultats.postes[posteType].nbEntrees++;
        }
        
        if (item.presenceParPoste && item.presenceParPoste[posteType] > 0) {
          resultats.postes[posteType].presenceTotal += item.presenceParPoste[posteType];
        }
      });
      
      // Regrouper par période (année-semaine)
      if (item.annee && item.semaine) {
        var periode = item.annee + '-S' + String(item.semaine).padStart(2, '0');
        if (!resultats.periodes[periode]) {
          resultats.periodes[periode] = {
            taiTotal: 0,
            presenceTotal: 0,
            nbEntrees: 0,
            taiParPoste: {},
            presenceParPoste: {}
          };
          
          // Initialiser les compteurs pour chaque type de poste
          postesTypes.forEach(function(posteType) {
            resultats.periodes[periode].taiParPoste[posteType] = 0;
            resultats.periodes[periode].presenceParPoste[posteType] = 0;
          });
        }
        
        // Incrémenter le compteur d'entrées
        resultats.periodes[periode].nbEntrees++;
        
        // Ajouter les TAI et présences
        resultats.periodes[periode].taiTotal += item.tai;
        resultats.periodes[periode].presenceTotal += item.presence;
        
        // Ajouter les TAI et présences par type de poste
        if (item.taiParPoste) {
          Object.keys(item.taiParPoste).forEach(function(posteType) {
            if (resultats.periodes[periode].taiParPoste.hasOwnProperty(posteType)) {
              resultats.periodes[periode].taiParPoste[posteType] += item.taiParPoste[posteType];
            }
          });
        }
        
        if (item.presenceParPoste) {
          Object.keys(item.presenceParPoste).forEach(function(posteType) {
            if (resultats.periodes[periode].presenceParPoste.hasOwnProperty(posteType)) {
              resultats.periodes[periode].presenceParPoste[posteType] += item.presenceParPoste[posteType];
            }
          });
        }
      }
    });
    
    // Calculer les rendements par équipe
    var statsEquipes = [];
    Object.keys(resultats.equipes).forEach(function(equipe) {
      var equipeData = resultats.equipes[equipe];
      
      // Calculer le rendement
      var rendement = equipeData.presenceTotal > 0 ? equipeData.taiTotal / equipeData.presenceTotal : 0;
      
      // Calculer les rendements par type de poste
      var rendementParPoste = {};
      Object.keys(equipeData.taiParPoste).forEach(function(posteType) {
        var taiPoste = equipeData.taiParPoste[posteType];
        var presencePoste = equipeData.presenceParPoste[posteType];
        rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
      });
      
      statsEquipes.push({
        equipe: equipe,
        rendement: rendement,
        taiTotal: equipeData.taiTotal,
        presenceTotal: equipeData.presenceTotal,
        nbEntrees: equipeData.nbEntrees,
        taiParPoste: equipeData.taiParPoste,
        presenceParPoste: equipeData.presenceParPoste,
        rendementParPoste: rendementParPoste
      });
    });
    
    // Calculer les rendements par poste
    var statsPostes = [];
    Object.keys(resultats.postes).forEach(function(poste) {
      var posteData = resultats.postes[poste];
      
      // Calculer le rendement
      var rendement = posteData.presenceTotal > 0 ? posteData.taiTotal / posteData.presenceTotal : 0;
      
      statsPostes.push({
        poste: poste,
        rendement: rendement,
        taiTotal: posteData.taiTotal,
        presenceTotal: posteData.presenceTotal,
        nbEntrees: posteData.nbEntrees
      });
    });
    
    // Calculer les rendements par période
    var statsPeriodes = [];
    Object.keys(resultats.periodes).forEach(function(periode) {
      var periodeData = resultats.periodes[periode];
      
      // Calculer le rendement
      var rendement = periodeData.presenceTotal > 0 ? periodeData.taiTotal / periodeData.presenceTotal : 0;
      
      // Calculer les rendements par type de poste
      var rendementParPoste = {};
      Object.keys(periodeData.taiParPoste).forEach(function(posteType) {
        var taiPoste = periodeData.taiParPoste[posteType];
        var presencePoste = periodeData.presenceParPoste[posteType];
        rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
      });
      
      statsPeriodes.push({
        periode: periode,
        rendement: rendement,
        taiTotal: periodeData.taiTotal,
        presenceTotal: periodeData.presenceTotal,
        nbEntrees: periodeData.nbEntrees,
        taiParPoste: periodeData.taiParPoste,
        presenceParPoste: periodeData.presenceParPoste,
        rendementParPoste: rendementParPoste
      });
    });
    
    // Trier les rendements par date
    resultats.rendements.sort(function(a, b) {
      return new Date(a.date) - new Date(b.date);
    });
    
    // Calculer le rendement global de l'opérateur
    var rendementGlobal = resultats.presenceTotal > 0 ? resultats.taiTotal / resultats.presenceTotal : 0;
    
    // Calculer les rendements par type de poste
    var rendementParPoste = {};
    Object.keys(resultats.taiParPoste).forEach(function(posteType) {
      var taiPoste = resultats.taiParPoste[posteType];
      var presencePoste = resultats.presenceParPoste[posteType];
      rendementParPoste[posteType] = presencePoste > 0 ? taiPoste / presencePoste : 0;
    });
    
    return {
      operateur: nomOperateur,
      sa: saOperateur, // Ajouter le SA dans le résultat
      rendement: rendementGlobal,
      taiTotal: resultats.taiTotal,
      presenceTotal: resultats.presenceTotal,
      taiParPoste: resultats.taiParPoste,
      presenceParPoste: resultats.presenceParPoste,
      rendementParPoste: rendementParPoste,
      rendements: resultats.rendements,
      statsPeriodes: statsPeriodes.sort(function(a, b) {
        return a.periode.localeCompare(b.periode);
      }),
      statsPostes: statsPostes.sort(function(a, b) {
        return b.rendement - a.rendement;
      }),
      statsEquipes: statsEquipes.sort(function(a, b) {
        return b.rendement - a.rendement;
      })
    };
  } catch (e) {
    Logger.log("ERREUR dans obtenirDonneesOperateur: " + e.toString());
    return {
      operateur: nomOperateur,
      rendements: [],
      statsPeriodes: [],
      statsPostes: [],
      statsEquipes: []
    };
  }
}
// Fonction pour obtenir des statistiques par poste
// Fonction pour obtenir des statistiques par poste
function obtenirStatistiquesParPoste(filtres) {
  try {
    Logger.log("Démarrage de obtenirStatistiquesParPoste() avec filtres: " + JSON.stringify(filtres));
    
    // Récupérer les données, filtrées si des filtres sont fournis
    var donnees;
    if (filtres && (filtres.dateDebut || filtres.dateFin || filtres.rendementMin !== null || filtres.rendementMax !== null || filtres.equipe || filtres.poste)) {
      donnees = obtenirDonneesFiltrees(filtres);
    } else {
      donnees = obtenirDonnees();
    }
    
    if (!donnees || donnees.erreur) {
      Logger.log("Erreur: " + (donnees ? donnees.erreur : "Aucune donnée reçue"));
      return [];
    }
    
    var resultats = {};
    var postesTypes = donnees.postesTypes || [];
    
    // Initialiser les résultats pour tous les types de postes
    postesTypes.forEach(function(posteType) {
      resultats[posteType] = {
        taiTotal: 0,
        presenceTotal: 0,
        nbEntrees: 0
      };
    });
    
    // Regrouper par type de poste
    donnees.donnees.forEach(function(item) {
      // Traiter les données de TAI et présence par type de poste
      if (item.taiParPoste && item.presenceParPoste) {
        Object.keys(item.taiParPoste).forEach(function(posteType) {
          if (postesTypes.includes(posteType)) {
            // Ajouter le TAI pour ce type de poste
            var taiPoste = item.taiParPoste[posteType] || 0;
            resultats[posteType].taiTotal += taiPoste;
            
            // Ajouter la présence pour ce type de poste
            // Important: S'assurer que nous utilisons les mêmes postes pour TAI et présence
            var presencePoste = item.presenceParPoste[posteType] || 0;
            resultats[posteType].presenceTotal += presencePoste;
            
            // Incrémenter le compteur d'entrées seulement si nous avons du TAI ou de la présence
            if (taiPoste > 0 || presencePoste > 0) {
              resultats[posteType].nbEntrees++;
            }
          }
        });
      }
    });
    
    // Calculer les statistiques
    var statistiques = [];
    Object.keys(resultats).forEach(function(poste) {
      var result = resultats[poste];
      
      // Ne calculer le rendement que si nous avons des entrées
      if (result.nbEntrees > 0) {
        // Calculer le rendement moyen (TAI total / Présence totale)
        var rendementMoyen = result.presenceTotal > 0 ? result.taiTotal / result.presenceTotal : 0;
        
        statistiques.push({
          poste: poste,
          rendementMoyen: rendementMoyen,
          taiTotal: result.taiTotal,
          presenceTotal: result.presenceTotal,
          nbEntrees: result.nbEntrees,
          taiMoyen: result.taiTotal / result.nbEntrees,
          presenceMoyen: result.presenceTotal / result.nbEntrees
        });
      }
    });
    
    // Trier par rendement décroissant
    statistiques.sort(function(a, b) {
      return b.rendementMoyen - a.rendementMoyen;
    });
    
    return statistiques;
  } catch (e) {
    Logger.log("ERREUR dans obtenirStatistiquesParPoste: " + e.toString());
    return [];
  }
}
// Fonction pour tester la connexion
function testConnection() {
  try {
    return {
      success: true,
      message: "Connexion réussie",
      timestamp: new Date().toString(),
      user: Session.getEffectiveUser().getEmail()
    };
  } catch (e) {
    return {
      success: false,
      message: "Erreur lors du test de connexion: " + e.toString(),
      error: e.toString()
    };
  }
}

// Fonction pour tester l'accès au classeur
function testerAccesClasseur() {
  try {
    Logger.log("Test d'accès au classeur");
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (sheet) {
      var info = {
        success: true,
        message: "Accès au classeur réussi",
        sheetName: sheet.getName(),
        rowCount: sheet.getLastRow(),
        columnCount: sheet.getLastColumn(),
        headers: sheet.getRange(2, 1, 1, Math.min(10, sheet.getLastColumn())).getValues()[0]
      };
      
      Logger.log("Test d'accès réussi: " + JSON.stringify(info));
      return info;
    } else {
      var info = {
        success: false,
        message: "Feuille 'Historique rendement' introuvable",
        availableSheets: spreadsheet.getSheets().map(function(s) { return s.getName(); })
      };
      
      Logger.log("Test d'accès échoué: Feuille introuvable");
      return info;
    }
  } catch (e) {
    Logger.log("ERREUR lors du test d'accès: " + e.toString());
    return {
      success: false,
      message: "Erreur d'accès au classeur: " + e.toString(),
      error: e.toString(),
      stack: e.stack
    };
  }
}

// Fonction utilitaire pour calculer la moyenne
function calculerMoyenne(tableau) {
  if (!tableau || tableau.length === 0) return 0;
  
  var somme = tableau.reduce(function(acc, val) {
    return acc + val;
  }, 0);
  
  return somme / tableau.length;
}

// Fonction pour obtenir des données filtrées (optimisée)
function obtenirDonneesFiltrees(filtres) {
  try {
    Logger.log("Démarrage de obtenirDonneesFiltrees() avec filtres : " + JSON.stringify(filtres));
    // Récupérer les données complètes
    var dataComplete = obtenirDonnees();
    
    if (!dataComplete || dataComplete.erreur) {
      Logger.log("Erreur lors de la récupération des données : " + (dataComplete ? dataComplete.erreur : "Aucune donnée"));
      return dataComplete; // Retourner l'erreur si elle existe
    }
    
    // Si aucun filtre n'est actif, retourner toutes les données
    if (!filtres || (
        !filtres.dateDebut && 
        !filtres.dateFin && 
        filtres.rendementMin === null && 
        filtres.rendementMax === null && 
        !filtres.equipe && 
        !filtres.poste)) {
      Logger.log("Aucun filtre actif - retour des données complètes");
      return dataComplete;
    }
    
    // Convertir les dates des filtres en objets Date
    var dateDebut = filtres.dateDebut ? new Date(filtres.dateDebut) : null;
    var dateFin = filtres.dateFin ? new Date(filtres.dateFin) : null;
    
    Logger.log("Dates converties - début: " + (dateDebut ? dateDebut.toISOString() : "null") + 
               ", fin: " + (dateFin ? dateFin.toISOString() : "null"));
    
    // Filtrer les données
    var donneesFiltrees = dataComplete.donnees.filter(function(item) {
      // Filtre par date
      if (dateDebut || dateFin) {
        var itemDate;
        try {
          itemDate = new Date(item.date);
        } catch (e) {
          Logger.log("Erreur de conversion de date: " + item.date);
          return false; // Exclure les éléments avec des dates invalides
        }
        
        if (dateDebut && itemDate < dateDebut) {
          return false;
        }
        
        if (dateFin) {
          // Définir l'heure à la fin de la journée pour inclure toute la journée
          var dateFinJour = new Date(dateFin);
          dateFinJour.setHours(23, 59, 59, 999);
          if (itemDate > dateFinJour) {
            return false;
          }
        }
      }
      
      // Filtre par rendement
      if (filtres.rendementMin !== null && !isNaN(filtres.rendementMin)) {
        if (item.rendement < filtres.rendementMin) return false;
      }
      
      if (filtres.rendementMax !== null && !isNaN(filtres.rendementMax)) {
        if (item.rendement > filtres.rendementMax) return false;
      }
      
      // Filtre par équipe
      if (filtres.equipe && item.equipe !== filtres.equipe) {
        return false;
      }
      
      // Filtre par poste - maintenant filtrer par poste spécifique (type de poste)
      if (filtres.poste) {
        // Si le filtre correspond à un type de poste, vérifier les valeurs TAI/présence pour ce type
        if (item.taiParPoste && item.taiParPoste.hasOwnProperty(filtres.poste)) {
          // Si le TAI pour ce type de poste est 0, exclure
          if (item.taiParPoste[filtres.poste] === 0) return false;
        } else {
          // Si l'entrée n'a pas de données pour ce type de poste, exclure
          return false;
        }
      }
      
      // Si tous les filtres ont été passés, inclure cet élément
      return true;
    });
    
    Logger.log("Données filtrées : " + donneesFiltrees.length + " / " + dataComplete.donnees.length + " enregistrements");
    
    // Créer un nouvel objet de résultat avec les données filtrées
    // mais conserver les listes complètes d'équipes, postes, etc.
    var resultat = {
      donnees: donneesFiltrees,
      postes: dataComplete.postes,
      equipes: dataComplete.equipes,
      operateurs: dataComplete.operateurs,
      postesTypes: dataComplete.postesTypes,
      filtresAppliques: filtres
    };
    
    return resultat;
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDonneesFiltrees : " + e.toString());
    Logger.log("Stack trace : " + e.stack);
    return { 
      erreur: "Erreur lors du filtrage des données : " + e.toString(),
      donnees: [],
      postes: [],
      equipes: [],
      operateurs: []
    };
  }
}

// Nouvelle fonction pour obtenir des données réduites pour le tableau de bord
function obtenirDonneesReduites() {
  try {
    Logger.log("Démarrage de obtenirDonneesReduites()");
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (!sheet) {
      Logger.log("Feuille 'Historique rendement' introuvable");
      return { 
        erreur: "Feuille 'Historique rendement' introuvable",
        donnees: [],
        postes: [],
        equipes: [],
        operateurs: []
      };
    }
    
    // Obtenir les en-têtes (deuxième ligne)
    var headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
    
    // Indices des colonnes de base
    var indexOperateur = headers.indexOf("Opérateur");
    var indexEquipe = headers.indexOf("Equipe");
    var indexPosteHabituel = headers.indexOf("Poste habituel");
    
    // Colonnes TAI (P à V) - index 15 à 21
    var taiColumns = [];
    for (var i = 15; i <= 21; i++) {
      if (i < headers.length && headers[i]) {
        taiColumns.push({
          nom: String(headers[i]),
          index: i
        });
      }
    }
    
    // Colonnes présence (X à AC) - index 23 à 28
    var presenceColumns = [];
    for (var i = 23; i <= 28; i++) {
      if (i < headers.length && headers[i]) {
        presenceColumns.push({
          nom: String(headers[i]),
          index: i
        });
      }
    }
    
    // Indices pour MASQUAGE et DEMASQUAGE
    var indexMasquage = -1;
    var indexDemasquage = -1;
    
    for (var i = 0; i < headers.length; i++) {
      var headerName = String(headers[i] || "");
      if (headerName === "MASQUAGE") {
        indexMasquage = i;
      } else if (headerName === "DEMASQUAGE") {
        indexDemasquage = i;
      }
    }
    
    // Récupérer uniquement un nombre limité de lignes pour éviter les problèmes de mémoire
    var maxRows = Math.min(1000, sheet.getLastRow() - 2);
    var dataRange = sheet.getRange(3, 1, maxRows, sheet.getLastColumn());
    var values = dataRange.getValues();
    
    // Création d'un tableau pour les données
    var donnees = [];
    var operateursSet = new Set();
    var equipesSet = new Set();
    
    // Extraire les noms de postes des entêtes
    var postesTypes = taiColumns.map(c => c.nom);
    
    // Parcourir les données
    for (var i = 0; i < values.length; i++) {
      var row = values[i];
      
      // Vérifier que la ligne contient des données valides
      if (row[indexOperateur]) {
        
        // Collecter les valeurs TAI par poste avec leur somme
        var taiTotal = 0;
        var taiParPoste = {};
        taiColumns.forEach(function(col) {
          var valeur = parseFloat(row[col.index]) || 0;
          taiParPoste[col.nom] = valeur;
          taiTotal += valeur;
        });
        
        // Collecter les valeurs de présence par poste avec leur somme
        var presenceTotal = 0;
        var presenceParPoste = {};
        presenceColumns.forEach(function(col) {
          var valeur = parseFloat(row[col.index]) || 0;
          presenceParPoste[col.nom] = valeur;
          presenceTotal += valeur;
        });
        
        // Calculer le rendement comme TAI Total / Présence Totale
        var rendement = presenceTotal > 0 ? taiTotal / presenceTotal : 0;
        
        // Ajouter l'entrée aux données
        var donneeItem = {
          operateur: String(row[indexOperateur] || ""),
          equipe: String(row[indexEquipe] || ""),
          poste: String(row[indexPosteHabituel] || ""), // Gardé pour compatibilité
          rendement: rendement,
          tai: taiTotal,
          presence: presenceTotal,
          taiParPoste: taiParPoste,
          presenceParPoste: presenceParPoste
        };
        
        // Ajouter MASQUAGE et DEMASQUAGE si disponibles
        if (indexMasquage !== -1) {
          donneeItem.masquage = parseFloat(row[indexMasquage]) || 0;
        }
        if (indexDemasquage !== -1) {
          donneeItem.demasquage = parseFloat(row[indexDemasquage]) || 0;
        }
        
        donnees.push(donneeItem);
        
        // Ajouter aux ensembles uniques
        if (row[indexEquipe]) equipesSet.add(String(row[indexEquipe]));
        if (row[indexOperateur]) operateursSet.add(String(row[indexOperateur]));
      }
    }
    
    // Convertir les ensembles en tableaux triés
    var equipesList = Array.from(equipesSet).sort();
    var operateursList = Array.from(operateursSet).sort();
    
    Logger.log("Données traitées: " + donnees.length + " enregistrements valides");
    
    // Retourner l'objet
    return {
      donnees: donnees,
      postes: postesTypes, // Utiliser les types de postes des en-têtes
      equipes: equipesList,
      operateurs: operateursList,
      postesTypes: postesTypes // Doublon pour compatibilité
    };
  
  } catch (e) {
    Logger.log("ERREUR dans obtenirDonneesReduites: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    return { 
      erreur: "Erreur lors de la lecture des données: " + e.toString(),
      donnees: [],
      postes: [],
      equipes: [],
      operateurs: []
    };
  }
}

// Fonction pour obtenir des données réduites filtrées
// Fonction pour obtenir des données réduites filtrées
function obtenirDonneesReduitesFiltrees(filtres) {
  try {
    Logger.log("Démarrage de obtenirDonneesReduitesFiltrees() avec filtres : " + JSON.stringify(filtres));
    
    // Vérification supplémentaire pour les filtres
    if (!filtres) {
      Logger.log("Filtres non définis");
      return obtenirDonneesReduites();
    }
    
    // Log complet des filtres
    Logger.log("Filtres détaillés: dateDebut=" + filtres.dateDebut + 
               ", dateFin=" + filtres.dateFin + 
               ", rendementMin=" + filtres.rendementMin + 
               ", rendementMax=" + filtres.rendementMax + 
               ", equipe=" + filtres.equipe + 
               ", poste=" + filtres.poste);
    
    // Si nous avons des filtres de date, nous devons utiliser les données complètes
    if (filtres.dateDebut || filtres.dateFin) {
      Logger.log("Filtres de date détectés, utilisation des données complètes");
      return filtrerDonneesCompletes(filtres);
    }
    
    // Sinon, on peut utiliser les données réduites
    var dataReduite = obtenirDonneesReduites();
    
    if (!dataReduite || dataReduite.erreur) {
      Logger.log("Erreur lors de l'obtention des données réduites: " + 
                 (dataReduite ? dataReduite.erreur : "Aucune donnée"));
      return dataReduite;
    }
    
    // Filtrer les données réduites
    var donneesFiltrees = dataReduite.donnees.filter(function(item) {
      // Filtre par rendement
      if (filtres.rendementMin !== null && filtres.rendementMin !== undefined) {
        if (item.rendement < parseFloat(filtres.rendementMin)) return false;
      }
      
      if (filtres.rendementMax !== null && filtres.rendementMax !== undefined) {
        if (item.rendement > parseFloat(filtres.rendementMax)) return false;
      }
      
      // Filtre par équipe
      if (filtres.equipe) {
        if (item.equipe !== filtres.equipe) return false;
      }
      
      // Filtre par poste - vérifier dans les données de TAI par poste
      if (filtres.poste) {
        if (item.taiParPoste && filtres.poste in item.taiParPoste) {
          if (item.taiParPoste[filtres.poste] === 0) return false;
        } else {
          return false;
        }
      }
      
      return true;
    });
    
    Logger.log("Filtrage réussi : " + donneesFiltrees.length + " / " + dataReduite.donnees.length);
    
    return {
      donnees: donneesFiltrees,
      postes: dataReduite.postes,
      equipes: dataReduite.equipes,
      operateurs: dataReduite.operateurs,
      postesTypes: dataReduite.postesTypes,
      filtresAppliques: filtres
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDonneesReduitesFiltrees : " + e.toString());
    Logger.log("Stack trace : " + e.stack);
    return { 
      erreur: "Erreur lors du filtrage des données réduites : " + e.toString(),
      donnees: [],
      postes: [],
      equipes: [],
      operateurs: []
    };
  }
}

// Nouvelle fonction pour filtrer les données complètes (utilisée pour les filtres de date)
function filtrerDonneesCompletes(filtres) {
  var donnees = obtenirDonnees();
  
  if (!donnees || donnees.erreur) {
    return donnees;
  }
  
  // Convertir les dates des filtres
  var dateDebut = filtres.dateDebut ? new Date(filtres.dateDebut) : null;
  var dateFin = filtres.dateFin ? new Date(filtres.dateFin) : null;
  
  // Si les dates ne sont pas valides, ne pas filtrer par date
  if (dateDebut && isNaN(dateDebut.getTime())) dateDebut = null;
  if (dateFin && isNaN(dateFin.getTime())) dateFin = null;
  
  Logger.log("Filtrage avec dates : début=" + (dateDebut ? dateDebut.toISOString() : "null") + 
             ", fin=" + (dateFin ? dateFin.toISOString() : "null"));
  
  // Filtrer les données
  var donneesFiltrees = donnees.donnees.filter(function(item) {
    // Filtre par date
    if (dateDebut || dateFin) {
      try {
        var itemDate = new Date(item.date);
        
        if (dateDebut && itemDate < dateDebut) {
          return false;
        }
        
        if (dateFin) {
          // Définir l'heure à la fin de la journée pour inclure toute la journée
          var dateFinJour = new Date(dateFin);
          dateFinJour.setHours(23, 59, 59, 999);
          if (itemDate > dateFinJour) {
            return false;
          }
        }
      } catch (e) {
        Logger.log("Erreur lors de la conversion de date : " + item.date);
        return false;
      }
    }
    
    // Filtre par rendement
    if (filtres.rendementMin !== null && filtres.rendementMin !== undefined) {
      if (item.rendement < parseFloat(filtres.rendementMin)) return false;
    }
    
    if (filtres.rendementMax !== null && filtres.rendementMax !== undefined) {
      if (item.rendement > parseFloat(filtres.rendementMax)) return false;
    }
    
    // Filtre par équipe
    if (filtres.equipe) {
      if (item.equipe !== filtres.equipe) return false;
    }
    
    // Filtre par poste
    if (filtres.poste) {
      if (item.poste !== filtres.poste) return false;
    }
    
    return true;
  });
  
  Logger.log("Filtrage complet réussi : " + donneesFiltrees.length + " / " + donnees.donnees.length);
  
  return {
    donnees: donneesFiltrees,
    postes: donnees.postes,
    equipes: donnees.equipes,
    operateurs: donnees.operateurs,
    filtresAppliques: filtres
  };
}

// Récupérer les informations du serveur
function getServerInfo() {
  try {
    Logger.log("Exécution de getServerInfo()");
    
    var info = {
      timestamp: new Date().toString(),
      timezone: Session.getScriptTimeZone(),
      user: Session.getEffectiveUser().getEmail(),
      serviceVersion: 'Apps Script',
      quotaRemaining: getQuotaRemaining()
    };
    
    return info;
  } catch(e) {
    Logger.log("Erreur dans getServerInfo: " + e.toString());
    return { error: e.toString() };
  }
}

// Récupérer les journaux récents
function getRecentLogs() {
  try {
    Logger.log("Récupération des journaux récents");
    
    var logs = [];
    var currentLog = Logger.getLog();
    
    // Limiter la taille des logs
    if (currentLog.length > 50000) {
      currentLog = currentLog.substring(currentLog.length - 50000);
    }
    
    return currentLog;
  } catch(e) {
    return "Erreur lors de la récupération des logs: " + e.toString();
  }
}

// Vérifier la structure de la feuille
function checkSheetStructure() {
  try {
    Logger.log("Vérification de la structure de la feuille");
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    
    // Obtenir les informations sur toutes les feuilles
    var sheets = spreadsheet.getSheets();
    var sheetsInfo = [];
    
    sheets.forEach(function(sheet) {
      var range = sheet.getDataRange();
      var firstRow = range.offset(0, 0, 1).getValues()[0];
      
      // Essayer de lire la deuxième ligne si elle existe
      var secondRow = [];
      if (range.getNumRows() > 1) {
        secondRow = range.offset(1, 0, 1).getValues()[0];
      }
      
      sheetsInfo.push({
        name: sheet.getName(),
        rows: range.getNumRows(),
        columns: range.getNumColumns(),
        firstRowSample: firstRow.slice(0, 10), // Limiter à 10 colonnes pour ne pas surcharger
        secondRowSample: secondRow.slice(0, 10)
      });
    });
    
    var result = {
      spreadsheetName: spreadsheet.getName(),
      spreadsheetUrl: spreadsheet.getUrl(),
      sheets: sheetsInfo,
      activeSheet: spreadsheet.getActiveSheet().getName()
    };
    
    Logger.log("Structure de la feuille récupérée avec succès");
    return result;
  } catch(e) {
    Logger.log("Erreur lors de la vérification de la structure: " + e.toString());
    return { error: e.toString() };
  }
}

// Obtenir un échantillon de données
function getSampleData() {
  try {
    Logger.log("Début de getSampleData()");
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    Logger.log("Classeur ouvert avec succès");
    
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    if (!sheet) {
      Logger.log("Feuille non trouvée");
      return { success: false, error: "Feuille 'Historique rendement' introuvable" };
    }
    
    Logger.log("Feuille trouvée: " + sheet.getName());
    
    // Récupérer les en-têtes (deuxième ligne)
    var headers = [];
    if (sheet.getLastRow() >= 2) {
      headers = sheet.getRange(2, 1, 1, sheet.getLastColumn()).getValues()[0];
      // Convertir les en-têtes en chaînes de caractères pour assurer la sérialisation
      headers = headers.map(function(header) { 
        return String(header || ""); 
      });
    }
    
    // Récupérer un échantillon de données (5 lignes maximum)
    var sampleData = [];
    var maxRows = Math.min(5, sheet.getLastRow() - 2);
    if (maxRows > 0) {
      var dataRange = sheet.getRange(3, 1, maxRows, Math.min(10, sheet.getLastColumn()));
      sampleData = dataRange.getValues();
      
      // Convertir les données en format sérialisable (éviter les dates, etc.)
      sampleData = sampleData.map(function(row) {
        return row.map(function(cell) {
          if (cell instanceof Date) {
            return Utilities.formatDate(cell, Session.getScriptTimeZone(), "yyyy-MM-dd");
          }
          return cell === null || cell === undefined ? "" : String(cell);
        });
      });
    }
    
    // Vérifier la présence de colonnes clés
    var keyColumns = ["Opérateur", "Equipe", "Poste habituel", "Date", "Année", "Semaine"];
    var columnsFound = {};
    
    keyColumns.forEach(function(col) {
      columnsFound[col] = headers.indexOf(col) >= 0;
    });
    
    // Rechercher les colonnes de rendement, TAI et présence
    var rdtGlobalFound = false;
    var taiTotalFound = false;
    var presenceTotalFound = false;
    
    headers.forEach(function(header) {
      header = String(header); // S'assurer que c'est une chaîne
      if (header.includes("Rdt") && header.includes("Global")) {
        rdtGlobalFound = true;
      }
      if (header.includes("TAI") && header.includes("Total")) {
        taiTotalFound = true;
      }
      if (header.includes("Présence") && header.includes("Total")) {
        presenceTotalFound = true;
      }
    });
    
    return {
      success: true,
      sheetName: sheet.getName(),
      rowCount: sheet.getLastRow(),
      columnCount: sheet.getLastColumn(),
      headers: headers,
      sampleData: sampleData,
      columnsCheck: {
        keyColumns: columnsFound,
        rdtGlobalFound: rdtGlobalFound,
        taiTotalFound: taiTotalFound,
        presenceTotalFound: presenceTotalFound
      }
    };
  } catch(e) {
    Logger.log("Erreur dans getSampleData: " + e.toString());
    return { success: false, error: "Erreur: " + e.toString() };
  }
}

// Obtenir le quota restant pour GoogleScript
function getQuotaRemaining() {
  try {
    return UrlFetchApp.getRemainingDailyQuota();
  } catch(e) {
    return -1;
  }
}
/**
 * Fonction pour obtenir les données d'inactivité
 * @param {Object} filtres - Les filtres actifs (dateDebut, dateFin, equipe, poste)
 * @return {Object} - Un objet contenant les données d'inactivité
 */
function obtenirDonneesInactivite(filtres) {
  try {
    Logger.log("Démarrage de obtenirDonneesInactivite() avec filtres : " + JSON.stringify(filtres));
    
    var spreadsheetId = '1Ni8E2HagtluqzpJLwUbBrgZdYtIuKud1jxtpK1StFS8';
    var spreadsheet = SpreadsheetApp.openById(spreadsheetId);
    var sheet = spreadsheet.getSheetByName('Historique rendement');
    
    if (!sheet) {
      return { 
        erreur: "Feuille 'Historique rendement' introuvable"
      };
    }
    
    var dataRange = sheet.getDataRange();
    var values = dataRange.getValues();
    var headers = values[1]; // Les en-têtes sont en deuxième ligne (index 1)
    
    // Identifier les indices des colonnes importantes
    var indexOperateur = headers.indexOf("Opérateur");
    var indexEquipe = headers.indexOf("Equipe");
    var indexDate = headers.indexOf("Date");
    var indexPoste = headers.indexOf("Poste habituel");
    
    // Identifier les colonnes d'inactivité
    var inactiviteColumns = {}; // Pour stocker les index des colonnes d'inactivité
    var typesInactivite = []; // Pour stocker les types d'inactivité
    
    // Rechercher les colonnes d'inactivité (supposées être étiquetées comme "INAC_XXX")
    for (var i = 0; i < headers.length; i++) {
      var header = String(headers[i] || "");
      if (header.startsWith("INAC_") || header.startsWith("Inac_") || 
          header === "ATTENTE" || header === "PAUSE" || header === "FORMATION" || 
          header === "REUNION" || header === "AUTRE") {
        
        // Extraire le type d'inactivité (tout ce qui vient après "INAC_")
        var typeInactivite = header.startsWith("INAC_") ? header.substring(5) : header;
        inactiviteColumns[typeInactivite] = i;
        typesInactivite.push(typeInactivite);
      }
    }
    
    // Si aucune colonne d'inactivité n'est trouvée, simuler des données pour démonstration
    if (Object.keys(inactiviteColumns).length === 0) {
      Logger.log("Aucune colonne d'inactivité trouvée, création de données de simulation");
      
      // Types simulés
      typesInactivite = ["ATTENTE", "PAUSE", "FORMATION", "REUNION", "AUTRE"];
      
      // Simulation des données d'inactivité
      return simulerDonneesInactivite(values, headers, indexOperateur, indexEquipe, indexPoste, indexDate, typesInactivite, filtres);
    }
    
    Logger.log("Types d'inactivité trouvés: " + typesInactivite.join(", "));
    
    // Filtrer les données selon les critères
    var donneesFiltrees = [];
    for (var i = 2; i < values.length; i++) { // Commencer à la 3ème ligne (index 2)
      var row = values[i];
      
      // Vérifier si la ligne contient un opérateur valide
      if (!row[indexOperateur]) continue;
      
      // Si des filtres sont spécifiés, les appliquer
      if (filtres) {
        // Filtre par date
        if ((filtres.dateDebut || filtres.dateFin) && indexDate >= 0) {
          var rowDate = row[indexDate];
          
          // Convertir la date en objet Date si ce n'est pas déjà le cas
          if (!(rowDate instanceof Date)) {
            try {
              // Si la date est au format string "DD/MM/YYYY"
              if (typeof rowDate === 'string') {
                var parts = rowDate.split('/');
                if (parts.length === 3) {
                  rowDate = new Date(parts[2], parts[1] - 1, parts[0]);
                }
              }
            } catch (e) {
              Logger.log("Erreur de conversion de date: " + e);
              continue; // Ignorer cette ligne en cas d'erreur
            }
          }
          
          if (filtres.dateDebut) {
            var dateDebut = new Date(filtres.dateDebut);
            if (rowDate < dateDebut) continue;
          }
          
          if (filtres.dateFin) {
            var dateFin = new Date(filtres.dateFin);
            dateFin.setHours(23, 59, 59, 999); // Fin de journée
            if (rowDate > dateFin) continue;
          }
        }
        
        // Filtre par équipe
        if (filtres.equipe && indexEquipe >= 0) {
          if (row[indexEquipe] !== filtres.equipe) continue;
        }
        
        // Filtre par poste
        if (filtres.poste && indexPoste >= 0) {
          if (row[indexPoste] !== filtres.poste) continue;
        }
      }
      
      // Calculer les totaux d'inactivité pour cette ligne
      var inactiviteTotal = 0;
      var inactiviteParType = {};
      
      // Initialiser les valeurs pour chaque type d'inactivité
      typesInactivite.forEach(function(type) {
        var colIndex = inactiviteColumns[type];
        var value = colIndex !== undefined ? (parseFloat(row[colIndex]) || 0) : 0;
        
        inactiviteParType[type] = value;
        inactiviteTotal += value;
      });
      
      // Si des données d'inactivité existent pour cette ligne, l'ajouter aux données filtrées
      if (inactiviteTotal > 0) {
        // Formater la date pour l'affichage
        var dateStr = "";
        if (typeof row[indexDate] === 'object' && row[indexDate] instanceof Date) {
          dateStr = Utilities.formatDate(row[indexDate], Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else if (typeof row[indexDate] === 'string') {
          dateStr = row[indexDate];
        }
        
        donneesFiltrees.push({
          operateur: row[indexOperateur],
          equipe: indexEquipe >= 0 ? row[indexEquipe] || "" : "",
          poste: indexPoste >= 0 ? row[indexPoste] || "" : "",
          date: dateStr,
          inactiviteTotal: inactiviteTotal,
          inactiviteParType: inactiviteParType
        });
      }
    }
    
    // Analyser les données filtrées
    var statsByType = analyserParType(donneesFiltrees, typesInactivite);
    var statsByOperateur = analyserParOperateur(donneesFiltrees, typesInactivite);
    var statsByEquipe = analyserParEquipe(donneesFiltrees, typesInactivite);
    
    return {
      donnees: donneesFiltrees,
      typesInactivite: typesInactivite,
      statsByType: statsByType,
      statsByOperateur: statsByOperateur,
      statsByEquipe: statsByEquipe
    };
    
  } catch (e) {
    Logger.log("ERREUR dans obtenirDonneesInactivite: " + e.toString());
    Logger.log("Stack trace: " + e.stack);
    return { erreur: "Erreur lors de la récupération des données d'inactivité: " + e.toString() };
  }
}

/**
 * Fonction pour simuler des données d'inactivité
 * Utilisée lorsque les colonnes d'inactivité ne sont pas trouvées
 */
function simulerDonneesInactivite(values, headers, indexOperateur, indexEquipe, indexPoste, indexDate, typesInactivite, filtres) {
  var donnees = [];
  var operateurs = new Set();
  var equipes = new Set();
  
  // Extraire les opérateurs et équipes uniques
  for (var i = 2; i < Math.min(values.length, 100); i++) {
    var row = values[i];
    if (row[indexOperateur]) {
      operateurs.add(row[indexOperateur]);
      if (indexEquipe >= 0 && row[indexEquipe]) {
        equipes.add(row[indexEquipe]);
      }
    }
  }
  
  operateurs = Array.from(operateurs);
  equipes = Array.from(equipes);
  
  // Générer des données simulées
  for (var i = 0; i < operateurs.length; i++) {
    var operateur = operateurs[i];
    var equipe = equipes[Math.min(i % equipes.length, equipes.length - 1)];
    
    // Simuler 1 à 3 entrées d'inactivité par opérateur
    var nbEntrees = Math.floor(Math.random() * 3) + 1;
    
    for (var j = 0; j < nbEntrees; j++) {
      var inactiviteTotal = 0;
      var inactiviteParType = {};
      
      // Générer des valeurs pour chaque type d'inactivité
      typesInactivite.forEach(function(type) {
        // Probabilité de 30% d'avoir une inactivité pour chaque type
        if (Math.random() < 0.3) {
          var value = Math.random() * 1.5; // Entre 0 et 1.5 heures
          inactiviteParType[type] = Math.round(value * 100) / 100; // Arrondir à 2 décimales
          inactiviteTotal += inactiviteParType[type];
        } else {
          inactiviteParType[type] = 0;
        }
      });
      
      // Générer une date dans les 30 derniers jours
      var date = new Date();
      date.setDate(date.getDate() - Math.floor(Math.random() * 30));
      var dateStr = Utilities.formatDate(date, Session.getScriptTimeZone(), "yyyy-MM-dd");
      
      // Ajouter l'entrée aux données
      donnees.push({
        operateur: operateur,
        equipe: equipe,
        poste: "Poste " + (Math.floor(Math.random() * 5) + 1),
        date: dateStr,
        inactiviteTotal: inactiviteTotal,
        inactiviteParType: inactiviteParType
      });
    }
  }
  
  // Analyser les données simulées
  var statsByType = analyserParType(donnees, typesInactivite);
  var statsByOperateur = analyserParOperateur(donnees, typesInactivite);
  var statsByEquipe = analyserParEquipe(donnees, typesInactivite);
  
  return {
    donnees: donnees,
    typesInactivite: typesInactivite,
    statsByType: statsByType,
    statsByOperateur: statsByOperateur,
    statsByEquipe: statsByEquipe
  };
}

/**
 * Analyser les données par type d'inactivité
 */
function analyserParType(donnees, typesInactivite) {
  var result = {};
  
  // Initialiser les résultats pour chaque type
  typesInactivite.forEach(function(type) {
    result[type] = {
      total: 0,
      moyenne: 0,
      max: 0,
      maxOperateur: "",
      entrees: 0
    };
  });
  
  // Analyser les données
  donnees.forEach(function(entry) {
    typesInactivite.forEach(function(type) {
      var value = entry.inactiviteParType[type] || 0;
      
      if (value > 0) {
        result[type].total += value;
        result[type].entrees++;
        
        if (value > result[type].max) {
          result[type].max = value;
          result[type].maxOperateur = entry.operateur;
        }
      }
    });
  });
  
  // Calculer les moyennes
  typesInactivite.forEach(function(type) {
    if (result[type].entrees > 0) {
      result[type].moyenne = result[type].total / result[type].entrees;
    }
  });
  
  return result;
}

/**
 * Analyser les données par opérateur
 */
function analyserParOperateur(donnees, typesInactivite) {
  var result = {};
  
  // Regrouper par opérateur
  donnees.forEach(function(entry) {
    var operateur = entry.operateur;
    
    if (!result[operateur]) {
      result[operateur] = {
        total: 0,
        parType: {},
        entrees: 0
      };
      
      // Initialiser les compteurs pour chaque type
      typesInactivite.forEach(function(type) {
        result[operateur].parType[type] = 0;
      });
    }
    
    // Ajouter les valeurs
    result[operateur].total += entry.inactiviteTotal;
    result[operateur].entrees++;
    
    typesInactivite.forEach(function(type) {
      result[operateur].parType[type] += entry.inactiviteParType[type] || 0;
    });
  });
  
  return result;
}

/**
 * Analyser les données par équipe
 */
function analyserParEquipe(donnees, typesInactivite) {
  var result = {};
  
  // Regrouper par équipe
  donnees.forEach(function(entry) {
    var equipe = entry.equipe || "Non spécifiée";
    
    if (!result[equipe]) {
      result[equipe] = {
        total: 0,
        parType: {},
        entrees: 0,
        operateurs: new Set()
      };
      
      // Initialiser les compteurs pour chaque type
      typesInactivite.forEach(function(type) {
        result[equipe].parType[type] = 0;
      });
    }
    
    // Ajouter les valeurs
    result[equipe].total += entry.inactiviteTotal;
    result[equipe].entrees++;
    result[equipe].operateurs.add(entry.operateur);
    
    typesInactivite.forEach(function(type) {
      result[equipe].parType[type] += entry.inactiviteParType[type] || 0;
    });
  });
  
  // Convertir les sets en nombres pour la sérialisation JSON
  Object.keys(result).forEach(function(equipe) {
    result[equipe].nbOperateurs = result[equipe].operateurs.size;
    delete result[equipe].operateurs; // Supprimer le set
  });
  
  return result;
}
