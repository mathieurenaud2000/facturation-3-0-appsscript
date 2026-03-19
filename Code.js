// GÉNÉRAL

function ouvrirPopup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const checkedRows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === true) checkedRows.push(i);
  }

  if (checkedRows.length === 0) {
    SpreadsheetApp.getUi().alert("Aucune ligne cochée.");
    return;
  }

  const clients = new Set();
  const campagnes = new Set();
  const projets = new Set();
  let totalPrix = 0;
  const tachesMap = {};

  checkedRows.forEach(row => {
    const ligne = data[row];
    const client = ligne[1];
    const campagne = ligne[2];
    const projet = ligne[3];
    const prix = parseFloat(ligne[10]) || 0;
    const tache = ligne[4];
    const temps = parseFloat(ligne[8]) || 0;

    if (client) clients.add(client);
    if (campagne) campagnes.add(campagne);
    if (projet) projets.add(projet);
    totalPrix += prix;

    if (tache) {
      if (!tachesMap[tache]) tachesMap[tache] = 0;
      tachesMap[tache] += temps;
    }
  });

  const clientStr = Array.from(clients).join(' + ');
  const campagneStr = Array.from(campagnes).join(' + ');
  const projetStr = Array.from(projets).join(' + ');
  const prixStr = `${totalPrix.toFixed(2)} $`;

  const tachesFinales = Object.entries(tachesMap)
    .map(([nom, duree]) => `${nom} (${duree.toFixed(2)}h)`)
    .join(' + ');

  const template = HtmlService.createTemplateFromFile('popup1');
  template.client = clientStr;
  template.campagne = campagneStr;
  template.projet = projetStr;
  template.prix = prixStr;
  template.taches = tachesFinales;

  const html = template.evaluate()
    .setWidth(800)
    .setHeight(200);

  SpreadsheetApp.getUi().showModelessDialog(html, 'Détail d’activité');

  // 🔁 Réinitialiser les cases cochées à FALSE
  checkedRows.forEach(row => {
    sheet.getRange(row + 1, 1).setValue(false); // +1 car data est 0-indexé
  });
}

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetGestion = ss.getSheetByName("GESTION");

  Logger.log('onOpen triggered at: ' + new Date());
  if (!sheetGestion) {
    Logger.log('Erreur: GESTION manquante');
    ui.alert("Erreur : La feuille 'GESTION' est manquante.");
    return;
  }

  const isConfigured = sheetGestion.getRange("D2").getValue();
  Logger.log('GESTION!D2 value: ' + isConfigured + ', type: ' + typeof isConfigured);
  if (isConfigured === false) {
    Logger.log('Calling info() because GESTION!D2 is FALSE');
    info();
  }

  ui.createMenu('🔧 Outils perso')
    .addItem('Ouvrir la fenêtre', 'ouvrirPopup')
    .addItem('Configuration', 'info')
    .addItem('Nouvelle entrée', 'newTimeEntry')
    .addItem('Facturer', 'Facturer')
    .addItem('Feuille de temps', 'FeuilleDeTemps')
    .addItem('Facturation', 'showFacturation')
    .addItem('Supprimer', 'trash')
    .addItem('Nouveau projet', 'nouveauProjet')
    .addItem('Dossier', 'dossier')
    .addToUi();
}

function nouveauProjet() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();

  const ligneSource = 6;
  const ligneInsertion = 7;

  // Insérer 2 lignes vides après la ligne 6 (donc avant ligne 7)
  sheet.insertRowsBefore(ligneInsertion, 2);

  // Copier contenu et format de la ligne 6
  const plageSource = sheet.getRange(ligneSource, 1, 1, sheet.getLastColumn());
  const valeurs = plageSource.getValues();
  const formats = plageSource.getNumberFormats();

  // Coller dans la nouvelle ligne 8 (qui est maintenant à l’index ligneInsertion + 1)
  const plageCible = sheet.getRange(ligneInsertion + 1, 1, 1, valeurs[0].length);
  plageCible.setValues(valeurs);
  plageCible.setNumberFormats(formats);
}

//  FACTURATION

function Facturer() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerConfigSheet = facturerSpreadsheet.getSheetByName("CONFIG");
  const facturerModelSheet = facturerSpreadsheet.getSheetByName("MODÈLE");
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");
  const facturerUi = SpreadsheetApp.getUi();

  if (!facturerTimeSheet || !facturerConfigSheet || !facturerModelSheet || !facturerTrackingSheet || !facturerGestionSheet) {
    facturerUi.alert("Erreur : Une ou plusieurs feuilles nécessaires sont manquantes.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const folderId = String(facturerGestionSheet.getRange("E2").getValue() || "");
  if (!folderId) {
    facturerUi.alert("Erreur : Aucun dossier Google Drive configuré dans GESTION!E2. Veuillez configurer via le menu Configuration.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let facturerDriveFolder;
  try {
    facturerDriveFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    facturerUi.alert("Erreur : L’ID du dossier dans GESTION!E2 est invalide.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const facturerLastInvoiceNumber = Number(facturerConfigSheet.getRange("A2").getValue()) || 0;
  const facturerNewInvoiceNumber = facturerLastInvoiceNumber + 1;
  const facturerContacts = facturerConfigSheet.getRange("B2:B" + facturerConfigSheet.getLastRow()).getValues().flat().filter(String);
  const facturerActivityTypes = facturerConfigSheet.getRange("C2:C" + facturerConfigSheet.getLastRow()).getValues().flat().filter(String);

  if (facturerContacts.length === 0 || facturerActivityTypes.length === 0) {
    facturerUi.alert("Erreur : La feuille CONFIG doit contenir au moins un contact (B2:B) et un type d'activité (C2:C).", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (facturerContacts.some(c => c.includes("${c}")) || facturerActivityTypes.some(a => a.includes("${a}"))) {
    facturerUi.alert("Erreur : Les colonnes B (contacts) ou C (types d'activité) dans CONFIG contiennent des données invalides (ex. : ${c}, ${a}). Veuillez corriger.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const facturerTimeData = facturerTimeSheet.getRange("A7:Q" + facturerTimeSheet.getLastRow()).getValues();
  const facturerCheckedRows = facturerTimeData.map((row, index) => ({ row: row, index: index + 7 }))
    .filter(row => row.row[0] === true);

  if (facturerCheckedRows.length === 0) {
    facturerUi.alert("Aucune ligne cochée en colonne A, veuillez sélectionner des activités.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  let facturerErrors = [];
  let facturerAlreadyInvoiced = [];
  facturerCheckedRows.forEach(row => {
    const [facturerCheckbox, facturerClient, facturerCampaign, facturerProject, facturerActivity, facturerUnused1, facturerUnused2, facturerTimeRaw, facturerUnused3, facturerPrice, facturerUnused4, facturerUnused5, facturerUnused6, facturerInvoiced, facturerInvoiceNumber, facturerInvoiceDate] = row.row;
    const facturerRowIndex = row.index;

    let facturerTime = facturerTimeRaw;
    if (facturerTimeRaw instanceof Date) {
      facturerTime = facturerTimeRaw.getHours() + facturerTimeRaw.getMinutes() / 60;
    } else if (typeof facturerTimeRaw === "string" || isNaN(Number(facturerTimeRaw))) {
      facturerErrors.push(`Ligne ${facturerRowIndex}: temps non valide (format incorrect : ${facturerTimeRaw})`);
      return;
    }

    if (!facturerClient || !facturerCampaign || !facturerProject || !facturerActivity || facturerTimeRaw == null || !facturerPrice) {
      facturerErrors.push(`Ligne ${facturerRowIndex}: ${!facturerClient ? "client manquant" : !facturerCampaign ? "campagne manquante" : !facturerProject ? "projet manquant" : !facturerActivity ? "activité manquante" : facturerTimeRaw == null ? "temps manquant" : "prix manquant"}`);
    }
    if (facturerInvoiced && facturerInvoiceDate) {
      facturerAlreadyInvoiced.push(`Ligne ${facturerRowIndex}: déjà facturée et payée`);
    }
    if (isNaN(facturerTime) || facturerTime <= 0) {
      facturerErrors.push(`Ligne ${facturerRowIndex}: temps non valide (${facturerTime <= 0 ? "temps négatif ou nul" : facturerTimeRaw})`);
    }
  });

  if (facturerErrors.length > 0) {
    facturerUi.alert("Données manquantes ou invalides :\n" + facturerErrors.join("\n"), SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  if (facturerAlreadyInvoiced.length > 0) {
    const response = facturerUi.alert("Lignes déjà facturées :\n" + facturerAlreadyInvoiced.join("\n"), 
      "Voulez-vous continuer (Oui) ou arrêter (Non) ?", SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (response === SpreadsheetApp.getUi().Button.NO) {
      return;
    }
  }

  const facturerItems = [];
  facturerCheckedRows.forEach(row => {
    const facturerClient = row.row[1];
    const facturerCampaign = row.row[2];
    const facturerKey = `${facturerClient}:${facturerCampaign}`;
    if (!facturerItems.some(item => item.key === facturerKey)) {
      facturerItems.push({ key: facturerKey, client: facturerClient, campaign: facturerCampaign, projects: [], activities: [], totalTime: 0, totalPrice: 0 });
    }
    const facturerItem = facturerItems.find(item => item.key === facturerKey);
    if (!facturerItem.projects.includes(row.row[3])) facturerItem.projects.push(row.row[3]);
    const existingActivity = facturerItem.activities.find(a => a.activity === row.row[4]);
    const time = row.row[8] instanceof Date ? (row.row[8].getHours() + row.row[8].getMinutes() / 60) : Number(row.row[8]);
    if (existingActivity) {
      existingActivity.time += time;
    } else {
      facturerItem.activities.push({ activity: row.row[4], time: time });
    }
    facturerItem.totalTime += time;
    facturerItem.totalPrice += Number(row.row[10]);
  });

  if (facturerItems.length > 7) {
    facturerUi.alert("Erreur : Plus de 7 items client/campagne sélectionnés. Maximum autorisé : 7.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const facturerHtml = HtmlService.createHtmlOutputFromFile("popup")
    .setWidth(400)
    .setHeight(350);
  facturerHtml.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  facturerHtml.append(`<script>var contacts = ${JSON.stringify(facturerContacts)}; var activityTypes = ${JSON.stringify(facturerActivityTypes)};</script>`);
  facturerUi.showModelessDialog(facturerHtml, "Nouvelle facture");
}

function checkConfigA2() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerConfigSheet = facturerSpreadsheet.getSheetByName("CONFIG");
  if (!facturerConfigSheet) {
    throw new Error("La feuille 'CONFIG' est manquante.");
  }
  const configA2Value = facturerConfigSheet.getRange("A2").getValue();
  return { isConfigA2Empty: configA2Value === "" || configA2Value === null || isNaN(Number(configA2Value)) };
}

function submitFacturerForm(contact, activityType, invoiceNumber) {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerConfigSheet = facturerSpreadsheet.getSheetByName("CONFIG");
  const facturerModelSheet = facturerSpreadsheet.getSheetByName("MODÈLE");
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const facturerUi = SpreadsheetApp.getUi();

  if (contact === "Sélectionnez un contact" || activityType === "Sélectionnez une activité") {
    return { success: false, message: "Veuillez sélectionner un contact et une activité générale." };
  }
  if (invoiceNumber !== null && (!Number.isInteger(invoiceNumber) || invoiceNumber <= 0)) {
    return { success: false, message: "Le numéro de facture doit être un entier positif." };
  }

  const facturerLastRow = facturerConfigSheet.getLastRow();
  const facturerContactsRange = facturerConfigSheet.getRange("B2:B" + Math.max(2, facturerLastRow)).getValues();
  const facturerActivitiesRange = facturerConfigSheet.getRange("C2:C" + Math.max(2, facturerLastRow)).getValues();
  if (!facturerContactsRange.flat().includes(contact)) {
    let facturerLastContactRow = 2;
    for (let i = 0; i < facturerContactsRange.length; i++) {
      if (!facturerContactsRange[i][0]) {
        facturerLastContactRow = i + 2;
        break;
      }
      facturerLastContactRow = i + 3;
    }
    facturerConfigSheet.getRange(`B${facturerLastContactRow}`).setValue(contact.trim());
  }
  if (!facturerActivitiesRange.flat().includes(activityType)) {
    let facturerLastActivityRow = 2;
    for (let i = 0; i < facturerActivitiesRange.length; i++) {
      if (!facturerActivitiesRange[i][0]) {
        facturerLastActivityRow = i + 2;
        break;
      }
      facturerLastActivityRow = i + 3;
    }
    facturerConfigSheet.getRange(`C${facturerLastActivityRow}`).setValue(activityType.trim());
  }

  const configA2Value = facturerConfigSheet.getRange("A2").getValue();
  const facturerNewInvoiceNumber = invoiceNumber !== null ? invoiceNumber : (Number(configA2Value) + 1 || 1);
  facturerConfigSheet.getRange("A2").setValue(facturerNewInvoiceNumber);

  facturerModelSheet.showSheet();
  const facturerTempSheet = facturerModelSheet.copyTo(facturerSpreadsheet).setName(`Facture${facturerNewInvoiceNumber}`);
  facturerModelSheet.hideSheet();

  const facturerTimeData = facturerTimeSheet.getRange("A7:Q" + facturerTimeSheet.getLastRow()).getValues();
  const facturerCheckedRows = facturerTimeData.map((row, index) => ({ row: row, index: index + 7 }))
    .filter(row => row.row[0] === true);
  const facturerItems = [];
  facturerCheckedRows.forEach(row => {
    const facturerClient = row.row[1];
    const facturerCampaign = row.row[2];
    const facturerKey = `${facturerClient}:${facturerCampaign}`;
    if (!facturerItems.some(item => item.key === facturerKey)) {
      facturerItems.push({ key: facturerKey, client: facturerClient, campaign: facturerCampaign, projects: [], activities: [], totalTime: 0, totalPrice: 0 });
    }
    const facturerItem = facturerItems.find(item => item.key === facturerKey);
    if (!facturerItem.projects.includes(row.row[3])) facturerItem.projects.push(row.row[3]);
    const existingActivity = facturerItem.activities.find(a => a.activity === row.row[4]);
    const time = row.row[8] instanceof Date ? (row.row[8].getHours() + row.row[8].getMinutes() / 60) : Number(row.row[8]);
    if (existingActivity) {
      existingActivity.time += time;
    } else {
      facturerItem.activities.push({ activity: row.row[4], time: time });
    }
    facturerItem.totalTime += time;
    facturerItem.totalPrice += Number(row.row[10]);
  });

  const facturerTotalAmount = facturerItems.reduce((sum, item) => sum + item.totalPrice, 0).toFixed(2);

  facturerTempSheet.getRange("L1").setValue(`#${facturerNewInvoiceNumber}`);
  facturerTempSheet.getRange("C7").setValue(Utilities.formatDate(new Date(), "EDT", "yyyy-MM-dd"));
  facturerTempSheet.getRange("C10").setValue(contact);
  facturerTempSheet.getRange("C12").setValue([...new Set(facturerCheckedRows.map(row => row.row[1]))].join(", "));
  facturerTempSheet.getRange("C14").setValue(activityType);
  facturerTempSheet.getRange("C17").setValue(Number(facturerTotalAmount));
  facturerTempSheet.getRange("N43").setValue(Number(facturerTotalAmount));
  facturerTempSheet.getRange("C45").setValue(Number(facturerTotalAmount));

  let facturerCurrentRow = 21;
  facturerItems.forEach((item, index) => {
    facturerTempSheet.getRange(`A${facturerCurrentRow}`).setValue(index + 1).setFontSize(10).setFontFamily("Roboto").setFontColor("#000000");
    const facturerProjectsStr = item.projects.join(", ");
    facturerTempSheet.getRange(`B${facturerCurrentRow}`).setValue(`${item.client} : ${item.campaign} (${facturerProjectsStr})`)
      .setFontFamily("Roboto").setFontSize(12).setFontColor("#000000").setFontWeight("bold");
    facturerTempSheet.getRange(`M${facturerCurrentRow}`).setValue(`${item.totalTime.toFixed(2)} h`).setFontFamily("Roboto").setFontSize(12).setFontColor("#000000");
    facturerTempSheet.getRange(`O${facturerCurrentRow}`).setValue(Number(item.totalPrice))
      .setFontFamily("Roboto").setFontSize(12).setFontColor("#000000").setFontWeight("bold").setHorizontalAlignment("right");
    const facturerActivitiesStr = item.activities.map(a => `${a.activity} (${a.time.toFixed(2)}h)`).join(", ");
    facturerTempSheet.getRange(`B${facturerCurrentRow + 1}`).setValue(facturerActivitiesStr)
      .setFontFamily("Roboto").setFontSize(11).setFontColor("#999999");
    facturerCurrentRow += 3;
  });

  const folderId = String(SpreadsheetApp.getActiveSpreadsheet().getSheetByName("GESTION").getRange("E2").getValue() || "");
  let facturerDriveFolder;
  try {
    facturerDriveFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    throw new Error("Erreur : Impossible d’accéder au dossier Drive défini dans GESTION!E2.");
  }
  const facturerFileName = `Facture${facturerNewInvoiceNumber}.pdf`;
  const facturerExistingFiles = facturerDriveFolder.getFilesByName(facturerFileName);
  if (facturerExistingFiles.hasNext()) {
    const response = facturerUi.alert(`Le fichier ${facturerFileName} existe déjà dans Google Drive. Remplacer ?`, SpreadsheetApp.getUi().ButtonSet.YES_NO);
    if (response === SpreadsheetApp.getUi().Button.NO) {
      facturerSpreadsheet.deleteSheet(facturerTempSheet);
      return { success: false, message: "Processus annulé : fichier existant non remplacé." };
    }
    facturerExistingFiles.next().setTrashed(true);
  }

  facturerTimeSheet.hideSheet();
  facturerTrackingSheet.hideSheet();
  facturerSpreadsheet.getSheets().forEach(sheet => {
    if (sheet.getName() !== `Facture${facturerNewInvoiceNumber}`) {
      sheet.hideSheet();
    } else {
      sheet.showSheet();
    }
  });
  SpreadsheetApp.flush();
  Utilities.sleep(1000);

  try {
    const facturerPdfBlob = facturerSpreadsheet.getAs("application/pdf").setName(facturerFileName);
    const facturerPdfFile = facturerDriveFolder.createFile(facturerPdfBlob);
    const facturerPdfUrl = facturerPdfFile.getUrl();

    let attempts = 0;
    const maxAttempts = 5;
    let fileExists = false;
    while (attempts < maxAttempts && !fileExists) {
      try {
        const file = DriveApp.getFileById(facturerPdfFile.getId());
        if (file.getUrl()) {
          fileExists = true;
        }
      } catch (e) {
        Utilities.sleep(1000);
        attempts++;
      }
    }
    if (!fileExists) {
      facturerSpreadsheet.deleteSheet(facturerTempSheet);
      return { success: false, message: "Erreur : Impossible de vérifier la sauvegarde du PDF après plusieurs tentatives." };
    }

    facturerSpreadsheet.getSheets().forEach(sheet => {
      if (["FEUILLE DE TEMPS", "FACTURATION"].includes(sheet.getName())) {
        sheet.showSheet();
      } else {
        sheet.hideSheet();
      }
    });

    facturerCheckedRows.forEach(row => {
      const facturerRowIndex = row.index;
      facturerTimeSheet.getRange(`A${facturerRowIndex}`).setValue(false);
      facturerTimeSheet.getRange(`O${facturerRowIndex}`).setValue(true);
      facturerTimeSheet.getRange(`P${facturerRowIndex}`).setValue(facturerNewInvoiceNumber);
      facturerTimeSheet.getRange(`Q${facturerRowIndex}`).setValue(Utilities.formatDate(new Date(), "EDT", "dd MM yyyy"));
    });

    const facturerTrackingRow = facturerTrackingSheet.getLastRow() + 1 >= 6 ? facturerTrackingSheet.getLastRow() + 1 : 6;
    facturerTrackingSheet.getRange(`B${facturerTrackingRow}`).setValue(`#${facturerNewInvoiceNumber}`).setNumberFormat("@");
    facturerTrackingSheet.getRange(`C${facturerTrackingRow}`).setValue(Utilities.formatDate(new Date(), "EDT", "dd MM yyyy"));
    facturerTrackingSheet.getRange(`D${facturerTrackingRow}`).setValue(contact);
    facturerTrackingSheet.getRange(`E${facturerTrackingRow}`).setValue([...new Set(facturerCheckedRows.map(row => row.row[1]))].join(", "));
    facturerTrackingSheet.getRange(`F${facturerTrackingRow}`).setValue(facturerTotalAmount + " $");
    facturerTrackingSheet.getRange(`G${facturerTrackingRow}`).setFormula(`=HYPERLINK("${facturerPdfUrl}"; "Voir PDF")`);
    facturerTrackingSheet.getRange(`U${facturerTrackingRow}`).setValue("");

    facturerSpreadsheet.deleteSheet(facturerTempSheet);
    facturerSpreadsheet.setActiveSheet(facturerTrackingSheet);
    const selectionRange = facturerTrackingSheet.getRange(`B${facturerTrackingRow}:G${facturerTrackingRow}`);
    facturerTrackingSheet.setActiveSelection(selectionRange);
    SpreadsheetApp.flush();

    // 🔒 Supprimer toute feuille non prévue dans la configuration standard
    const feuillesAutorisees = ["FEUILLE DE TEMPS", "FACTURATION", "MODÈLE", "GESTION", "CONFIG"];
    facturerSpreadsheet.getSheets().forEach(sheet => {
      const nom = sheet.getName();
      if (!feuillesAutorisees.includes(nom)) {
        facturerSpreadsheet.deleteSheet(sheet);
      }
    });

    return { success: true, pdfUrl: facturerPdfUrl };
  } catch (e) {
    Logger.log(`Exception: ${e.message}`);
    facturerSpreadsheet.getSheets().forEach(sheet => {
      if (["FEUILLE DE TEMPS", "FACTURATION"].includes(sheet.getName())) {
        sheet.showSheet();
      } else {
        sheet.hideSheet();
      }
    });
    facturerSpreadsheet.deleteSheet(facturerTempSheet);
    return { success: false, message: `Erreur inattendue : ${e.message}` };
  }
}

// NOUVELLE ENTRÉ DE TEMPS

function newTimeEntry() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");

  if (!facturerTimeSheet || !facturerGestionSheet) {
    SpreadsheetApp.getUi().alert("Erreur : La feuille 'FEUILLE DE TEMPS' ou 'GESTION' est manquante.", SpreadsheetApp.getUi().ButtonSet.OK);
    return;
  }

  const lastRowGestion = facturerGestionSheet.getLastRow();
  const clients = facturerGestionSheet.getRange("B2:B" + Math.max(2, lastRowGestion)).getValues().flat().filter(String);
  const activities = facturerGestionSheet.getRange("A2:A" + Math.max(2, lastRowGestion)).getValues().flat().filter(String);
  let rates = ['0']; // Valeur par défaut
  try {
    rates = facturerGestionSheet.getRange("C2:C" + Math.max(2, lastRowGestion)).getValues().flat().filter(String);
    if (rates.length === 0) rates = ['0'];
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erreur : Impossible de lire les taux dans GESTION!C2:C. Valeur par défaut utilisée.", SpreadsheetApp.getUi().ButtonSet.OK);
  }

  const lastRow = facturerTimeSheet.getLastRow();
  if (lastRow < 7) return;

  const checkBoxData = facturerTimeSheet.getRange("A7:A" + lastRow).getValues();
  const checkedIndexes = [];

  for (let i = 0; i < checkBoxData.length; i++) {
    if (checkBoxData[i][0] === true) {
      checkedIndexes.push(i + 7);
    }
  }

  if (checkedIndexes.length > 1) {
    SpreadsheetApp.getUi().alert("Sélectionne une ligne max.");
    return;
  }

  const checkedRowIndex = checkedIndexes.length === 1 ? checkedIndexes[0] : -1;

  const html = HtmlService.createTemplateFromFile("popupTemps");
  html.clients = clients || [];
  html.activities = activities || [];
  html.rates = rates || ['0'];
  html.checkedRowIndex = checkedRowIndex;

  if (checkedRowIndex !== -1) {
    let sourceData;
    try {
      sourceData = facturerTimeSheet.getRange(`B${checkedRowIndex}:E${checkedRowIndex}`).getValues()[0];
    } catch (e) {
      sourceData = ["", "", "", ""];
    }
    html.clientSelected = sourceData[0] || "";
    html.campaign = sourceData[1] || "";
    html.project = sourceData[2] || "";
    html.activitySelected = sourceData[3] || "";
    html.newRow = checkedRowIndex + 1;
  } else {
    html.clientSelected = "";
    html.campaign = "";
    html.project = "";
    html.activitySelected = "";
    html.newRow = 7;
  }

  const htmlOutput = html.evaluate().setWidth(400).setHeight(350);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "Nouvelle entrée de temps");
}

function submitTimeEntryForm(client, campaign, project, activity, newRow, checkedRowIndex, newClient, newActivity, rate) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTime = ss.getSheetByName("FEUILLE DE TEMPS");
  const sheetGestion = ss.getSheetByName("GESTION");

  if (!sheetTime || !sheetGestion) {
    SpreadsheetApp.getUi().alert("Erreur : La feuille 'FEUILLE DE TEMPS' ou 'GESTION' est manquante.");
    return;
  }

  // Ajouter le nouveau client à la colonne B de GESTION
  if (newClient && newClient.trim() !== "") {
    const clients = sheetGestion.getRange("B2:B" + sheetGestion.getLastRow()).getValues().flat();
    if (!clients.map(c => c.toString().toLowerCase()).includes(newClient.trim().toLowerCase())) {
      const insertRow = sheetGestion.getLastRow() + 1;
      sheetGestion.getRange("B" + insertRow).setValue(newClient.trim());
      const range = sheetGestion.getRange("B2:B" + sheetGestion.getLastRow());
      range.sort({ column: 2, ascending: true });
    }
    client = newClient.trim();
  }

  // Ajouter la nouvelle activité à la colonne A de GESTION
  if (newActivity && newActivity.trim() !== "") {
    const activities = sheetGestion.getRange("A2:A" + sheetGestion.getLastRow()).getValues().flat();
    if (!activities.map(a => a.toString().toLowerCase()).includes(newActivity.trim().toLowerCase())) {
      const insertRow = sheetGestion.getLastRow() + 1;
      sheetGestion.getRange("A" + insertRow).setValue(newActivity.trim());
      const range = sheetGestion.getRange("A2:A" + sheetGestion.getLastRow());
      range.sort({ column: 1, ascending: true });
    }
    activity = newActivity.trim();
  }

  // Ajouter le nouveau taux horaire à la colonne C de GESTION
  if (rate && rate.trim() !== "") {
    const rates = sheetGestion.getRange("C2:C" + sheetGestion.getLastRow()).getValues().flat();
    if (!rates.map(r => r.toString().toLowerCase()).includes(rate.trim().toLowerCase())) {
      const insertRow = sheetGestion.getLastRow() + 1;
      sheetGestion.getRange("C" + insertRow).setValue(rate.trim());
      const range = sheetGestion.getRange("C2:C" + sheetGestion.getLastRow());
      range.sort({ column: 3, ascending: false });
    }
  }

  const now = new Date();
  const dateStr = Utilities.formatDate(now, "America/Guayaquil", "yyyy-MM-dd");
  const timeStr = Utilities.formatDate(now, "America/Guayaquil", "HH:mm");

  if (checkedRowIndex !== -1) {
    // Insertion après ligne cochée
    sheetTime.insertRowAfter(checkedRowIndex);
    const targetRow = checkedRowIndex + 1;
    const sourceRange = sheetTime.getRange(`A${checkedRowIndex}:Z${checkedRowIndex}`);
    const targetRange = sheetTime.getRange(`A${targetRow}:Z${targetRow}`);
    sourceRange.copyTo(targetRange, SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);

    sheetTime.getRange(`D${targetRow}:H${targetRow}`).clearContent();
    sheetTime.getRange(`I${targetRow}`).setFormula(`=IF(H${targetRow}<>""; ROUND((IF(H${targetRow}<G${targetRow}; H${targetRow}+1; H${targetRow})-G${targetRow})*96)/4; "")`);
    sheetTime.getRange(`J${targetRow}`).setFormula(`=IF(H${targetRow}<>""; J${targetRow-1}+I${targetRow}; "")`);
    sheetTime.getRange(`K${targetRow}`).setFormula(`=IF(H${targetRow}<>""; I${targetRow}*T${targetRow}; "")`);
    sheetTime.getRange(`L${targetRow}`).setFormula(`=IF(H${targetRow}<>""; L${targetRow-1}+K${targetRow}; "")`);

    sheetTime.getRange(`B${targetRow}`).setValue(client);
    sheetTime.getRange(`C${targetRow}`).setValue(campaign);
    sheetTime.getRange(`D${targetRow}`).setValue(project);
    sheetTime.getRange(`E${targetRow}`).setValue(activity);
    sheetTime.getRange(`F${targetRow}`).setValue(dateStr);
    sheetTime.getRange(`G${targetRow}`).setValue(timeStr);
    sheetTime.getRange(`T${targetRow}`).setValue(rate);
    sheetTime.getRange(`U${targetRow}`).setValue("");
    sheetTime.getRange(`A${targetRow}`).setValue(true);

    sheetTime.getRange(`A${checkedRowIndex}`).setValue(false);

    const rangeEffet = sheetTime.getRange(`A${targetRow}:Z${targetRow}`);
    rangeEffet.setBackground("#e7efe1");
    sheetTime.getRange("I3").setBackground("#6aa84f");
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    rangeEffet.setBackground("#ffffff");
    sheetTime.setActiveSelection(`H${targetRow}`);
    SpreadsheetApp.flush();

  } else {
    // Aucune case cochée, insérer à ligne 7
    sheetTime.insertRowsAfter(6, 2);
    sheetTime.getRange("A9:Z9").copyTo(sheetTime.getRange("A7:Z7"), SpreadsheetApp.CopyPasteType.PASTE_NORMAL, false);
    sheetTime.getRange("B7:H7").clearContent();
    sheetTime.getRange("P7").clearContent();
    sheetTime.getRange("Q7").clearContent();
    sheetTime.getRange("S7").clearContent();
    sheetTime.getRange("U7").setValue("");
    sheetTime.getRange("N7").setValue(false);
    sheetTime.getRange("O7").setValue(false);
    sheetTime.getRange("R7").setValue(false);

    sheetTime.getRange("I7").setFormula(`=IF(H7<>""; ROUND((IF(H7<G7; H7+1; H7)-G7)*96)/4; "")`);
    sheetTime.getRange("J7").setFormula(`=IF(H7<>""; J6+I7; "")`);
    sheetTime.getRange("K7").setFormula(`=IF(H7<>""; I7*T7; "")`);
    sheetTime.getRange("L7").setFormula(`=IF(H7<>""; L6+K7; "")`);

    sheetTime.getRange("B7").setValue(client);
    sheetTime.getRange("C7").setValue(campaign);
    sheetTime.getRange("D7").setValue(project);
    sheetTime.getRange("E7").setValue(activity);
    sheetTime.getRange("F7").setValue(dateStr);
    sheetTime.getRange("G7").setValue(timeStr);
    sheetTime.getRange("T7").setValue(rate);
    sheetTime.getRange("A7").setValue(true);

    const rangeEffet = sheetTime.getRange("A7:Z7");
    rangeEffet.setBackground("#e7efe1");
    sheetTime.getRange("I3").setBackground("#6aa84f");
    SpreadsheetApp.flush();
    Utilities.sleep(1000);
    rangeEffet.setBackground("#ffffff");
    sheetTime.setActiveSelection("H7");
    SpreadsheetApp.flush();
  }

  // Afficher FEUILLE DE TEMPS et FACTURATION, cacher les autres
  ss.getSheets().forEach(sheet => {
    if (["FEUILLE DE TEMPS", "FACTURATION"].includes(sheet.getName())) {
      sheet.showSheet();
    } else {
      sheet.hideSheet();
    }
  });
}

function FeuilleDeTemps() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("FEUILLE DE TEMPS");
  if (sheet) {
    sheet.showSheet();
    sheet.activate();
  } else {
    SpreadsheetApp.getUi().alert("La feuille 'FEUILLE DE TEMPS' est introuvable.");
  }
}

// SUPPRIMER : Supression des lignes cochées

function trash() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("A7:A").getValues();
  const ligneCocheeCount = data.filter(row => row[0] === true).length;

  const html = HtmlService.createTemplateFromFile("confirmDelete");
  html.ligneCocheeCount = ligneCocheeCount;

  const output = html.evaluate()
    .setWidth(400)
    .setHeight(220);

  SpreadsheetApp.getUi().showModalDialog(output, `Supprimer ${ligneCocheeCount} ligne${ligneCocheeCount > 1 ? 's' : ''}`);
}

function supprimerLignesCochées() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== "FEUILLE DE TEMPS") {
    SpreadsheetApp.getUi().alert("Cette fonction ne peut être utilisée que dans la feuille 'FEUILLE DE TEMPS'.");
    return;
  }

  const data = sheet.getRange("A1:A" + sheet.getLastRow()).getValues();
  const lignesASupprimer = [];

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === true) {
      lignesASupprimer.push(i + 1);
    }
  }

  if (lignesASupprimer.length === 0) {
    SpreadsheetApp.getUi().alert("Aucune ligne cochée à supprimer.");
    return;
  }

  // 🟫 Effet visuel gris temporaire
  lignesASupprimer.forEach(row => {
    const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
    range.setBackground("#dddddd");
  });

  SpreadsheetApp.flush();
  Utilities.sleep(250);

  // 🗑️ Supprimer les lignes cochées (du bas vers le haut)
  lignesASupprimer.forEach(row => {
    sheet.deleteRow(row);
  });

  // ✅ Supprimer les lignes vides consécutives à partir de la ligne 7
  let ligneActuelle = 7;
  while (ligneActuelle <= sheet.getLastRow()) {
    const ligne = sheet.getRange(ligneActuelle, 1, 1, sheet.getLastColumn()).getValues()[0];
    const estVide = ligne.every(cell => cell === "" || cell === null);
    if (estVide) {
      sheet.deleteRow(ligneActuelle);
    } else {
      break;
    }
  }
}

// AFFICHER : La feuille facturation

function showFacturation() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");

  // Vérifier si la feuille FACTURATION existe
  if (!facturerTrackingSheet) {
    SpreadsheetApp.getUi().alert("Erreur : La feuille FACTURATION n'existe pas.");
    return;
  }

  // Afficher FEUILLE DE TEMPS et FACTURATION, cacher les autres
  facturerSpreadsheet.getSheets().forEach(sheet => {
    if (["FEUILLE DE TEMPS", "FACTURATION"].includes(sheet.getName())) {
      sheet.showSheet();
    } else {
      sheet.hideSheet();
    }
  });

  // Activer la feuille FACTURATION
  facturerSpreadsheet.setActiveSheet(facturerTrackingSheet);

  // Trouver la dernière ligne non vide dans FACTURATION (à partir de B6)
  const lastRow = facturerTrackingSheet.getLastRow();
  const facturerTrackingRow = lastRow >= 6 ? lastRow : 6;

  // Sélectionner la plage B:G de la dernière ligne
  const selectionRange = facturerTrackingSheet.getRange(`B${facturerTrackingRow}:G${facturerTrackingRow}`);
  facturerTrackingSheet.setActiveSelection(selectionRange);

  // Appliquer immédiatement la sélection
  SpreadsheetApp.flush();
}

// STOP : Arrête l'entrée de temps de la ligne cochée

function checkAndSetTime() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTime = ss.getSheetByName("FEUILLE DE TEMPS");
  
  if (!sheetTime) {
    SpreadsheetApp.getUi().alert("Erreur : La feuille 'FEUILLE DE TEMPS' est manquante.");
    return;
  }

  // Obtenir les cases cochées dans la colonne A à partir de A7
  const lastRow = sheetTime.getLastRow();
  if (lastRow < 7) {
    SpreadsheetApp.getUi().alert("Erreur : Aucune donnée à partir de la ligne 7.");
    return;
  }

  const checkBoxData = sheetTime.getRange(`A7:A${lastRow}`).getValues();
  const checkedRows = checkBoxData
    .map((row, index) => ({ checked: row[0], index: index + 7 }))
    .filter(row => row.checked === true);

  // Vérifier si exactement une ligne est cochée
  if (checkedRows.length > 1) {
    SpreadsheetApp.getUi().alert("Ne cocher qu'une seule ligne.");
    return;
  }

  if (checkedRows.length === 0) {
    SpreadsheetApp.getUi().alert("Erreur : Aucune ligne cochée.");
    return;
  }

  // Une seule ligne cochée
  const rowIndex = checkedRows[0].index;

  // Vérifier la couleur de fond de I3 dans FEUILLE DE TEMPS
  const backgroundColor = sheetTime.getRange("I3").getBackground();

  if (backgroundColor === "#ffffff") {
    // Couleur blanche, ne rien faire
    return;
  } else if (backgroundColor === "#6aa84f") {
    // Couleur verte, vérifier si H de la ligne cochée est vide
    const cellH = sheetTime.getRange(`H${rowIndex}`);
    if (cellH.isBlank()) {
      // Cellule H vide, écrire l'heure actuelle
      const now = new Date();
      const timeStr = Utilities.formatDate(now, "America/Guayaquil", "HH:mm");
      cellH.setValue(timeStr);
      // Changer la couleur de I3 en blanc
      sheetTime.getRange("I3").setBackground("#ffffff");
      SpreadsheetApp.flush();
    } else {
      // Cellule H non vide, afficher popup
      SpreadsheetApp.getUi().alert(
        "Action impossible sur cette ligne",
        "La cellule H de la ligne cochée contient déjà une valeur.",
        SpreadsheetApp.getUi().ButtonSet.OK
      );
    }
  }
}

// INFO : Change les données sur MODÈLE

function info() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetModele = ss.getSheetByName("MODÈLE");
  const sheetGestion = ss.getSheetByName("GESTION");

  if (!sheetModele || !sheetGestion) {
    SpreadsheetApp.getUi().alert("Erreur : La feuille 'MODÈLE' ou 'GESTION' est manquante.");
    return;
  }

  const name = String(sheetModele.getRange("L10").getValue() || "");
  const address = String(sheetModele.getRange("L11").getValue() || "");
  const addressLines = address ? address.split("\n") : ["", "", "", ""];
  const address1 = String(addressLines[0] || "");
  const address2 = String(addressLines[1] || "");
  const address3 = String(addressLines[2] || "");
  const address4 = String(addressLines[3] || "");
  const email = String(sheetModele.getRange("L16").getValue() || "");
  const website = String(sheetModele.getRange("L17").getValue() || "");
  const folderId = String(sheetGestion.getRange("E2").getValue() || "");
  const folderUrl = folderId ? `https://drive.google.com/drive/folders/${folderId}` : "";

  const nextInvoice = String(sheetGestion.getRange("A2").getValue() || "");

  const html = HtmlService.createTemplateFromFile("popupInfo");
  html.name = name;
  html.address1 = address1;
  html.address2 = address2;
  html.address3 = address3;
  html.address4 = address4;
  html.email = email;
  html.website = website;
  html.folderUrl = folderUrl;
  html.nextInvoice = nextInvoice;

  SpreadsheetApp.getUi().showModalDialog(
    html.evaluate().setWidth(600).setHeight(450),
    "Configuration de la facture"
  );
}

function submitInfoForm(name, address1, address2, address3, address4, email, website, folderUrl, nextInvoice) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getSheetByName("CONFIG");
  if (!sheet) throw new Error("Feuille CONFIG introuvable.");

  // Mettre à jour les cellules de CONFIG
  sheet.getRange("B2").setValue(name);
  sheet.getRange("B3").setValue(address1);
  sheet.getRange("B4").setValue(address2);
  sheet.getRange("B5").setValue(address3);
  sheet.getRange("B6").setValue(address4);
  sheet.getRange("B7").setValue(email);
  sheet.getRange("B8").setValue(website);
  sheet.getRange("B9").setValue(folderUrl);

  // Comparer et mettre à jour A2 si nécessaire
  const currentInvoice = sheet.getRange("A2").getValue();
  const parsedNextInvoice = parseInt(nextInvoice, 10);

  if (!isNaN(parsedNextInvoice) && parsedNextInvoice !== currentInvoice) {
    sheet.getRange("A2").setValue(parsedNextInvoice - 1);
  }
}

// DOSSIER : Ouvre le dossier avec les PDF

function dossier() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetGestion = ss.getSheetByName("GESTION");

  if (!sheetGestion) {
    SpreadsheetApp.getUi().alert("Erreur : La feuille 'GESTION' est manquante.");
    return;
  }

  const folderId = String(sheetGestion.getRange("E2").getValue() || "");
  if (!folderId) {
    SpreadsheetApp.getUi().alert("Erreur : Aucun dossier Google Drive configuré dans GESTION!E2. Veuillez configurer via le menu Configuration.");
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const url = folder.getUrl();
    const html = `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`;
    SpreadsheetApp.getUi().showModalDialog(HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1), "Ouvrir le dossier");
  } catch (e) {
    SpreadsheetApp.getUi().alert("Erreur : L’ID du dossier dans GESTION!E2 est invalide.");
  }
}

//////////

// ERREURS : Gestion des erreurs inattendues

function onFacturerError(error) {
  return { success: false, message: `Erreur inattendue : ${error.message}` };
}



/////// TEST OUVERTURE GRANDE FENËTRE



function ouvrirPleinEcranModeless() {
  var html = HtmlService.createHtmlOutputFromFile('pleinEcran')
    .setTitle('Plein écran')
    .setWidth(1600)   // Largeur max affichable
    .setHeight(1600);  // Hauteur max affichable
  SpreadsheetApp.getUi().showModelessDialog(html, 'Plein écran');
}

function traiterMessageDepuisHTML(message) {
  Logger.log("Message reçu : " + message);
}

function ouvrirCigaleEtFourmi() {
  var html = HtmlService.createHtmlOutputFromFile('pleinEcran')
    .setTitle('La Cigale et la Fourmi')
    .setWidth(1200)
    .setHeight(800);
  SpreadsheetApp.getUi().showModelessDialog(html, 'La Cigale et la Fourmi');
}


