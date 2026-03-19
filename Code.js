// GÉNÉRAL

function openStandaloneMessageView_(message, title = "Erreur", options = {}) {
  const { width = 400, height = 220 } = options;
  const html = HtmlService.createHtmlOutputFromFile("popup")
    .setWidth(width)
    .setHeight(height);
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  html.append(`<script>var showStandaloneMessage = true; var messageViewTitle = ${JSON.stringify(title)}; var messageViewMessage = ${JSON.stringify(message)};</script>`);
  SpreadsheetApp.getUi().showModelessDialog(html, title);
}

function ouvrirPopup() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();
  const data = sheet.getDataRange().getValues();
  const checkedRows = [];

  for (let i = 1; i < data.length; i++) {
    if (data[i][0] === true) checkedRows.push(i);
  }

  if (checkedRows.length === 0) {
    openStandaloneMessageView_("Aucune ligne cochée.", "Information");
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
    openStandaloneMessageView_("Erreur : La feuille 'GESTION' est manquante.");
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
    openStandaloneMessageView_("Erreur : Une ou plusieurs feuilles nécessaires sont manquantes.");
    return;
  }

  const folderId = String(facturerGestionSheet.getRange("E2").getValue() || "");
  if (!folderId) {
    openStandaloneMessageView_("Erreur : Aucun dossier configuré.");
    return;
  }

  let facturerDriveFolder;
  try {
    facturerDriveFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    openStandaloneMessageView_("Erreur : ID de dossier invalide.");
    return;
  }

  const facturerContacts = facturerConfigSheet.getRange("B2:B" + facturerConfigSheet.getLastRow()).getValues().flat().filter(String);
  const facturerActivityTypes = facturerConfigSheet.getRange("C2:C" + facturerConfigSheet.getLastRow()).getValues().flat().filter(String);

  if (facturerContacts.length === 0 || facturerActivityTypes.length === 0) {
    openStandaloneMessageView_("Erreur : La feuille CONFIG doit contenir au moins un contact (B2:B) et un type d'activité (C2:C).");
    return;
  }

  if (facturerContacts.some(c => c.includes("${c}")) || facturerActivityTypes.some(a => a.includes("${a}"))) {
    openStandaloneMessageView_("Erreur : Les colonnes B (contacts) ou C (types d'activité) dans CONFIG contiennent des données invalides (ex. : ${c}, ${a}). Veuillez corriger.");
    return;
  }

  const facturerTimeData = facturerTimeSheet.getRange("A7:Q" + facturerTimeSheet.getLastRow()).getValues();
  const facturerCheckedRows = facturerTimeData.map((row, index) => ({ row: row, index: index + 7 }))
    .filter(row => row.row[0] === true);

  if (facturerCheckedRows.length === 0) {
    openStandaloneMessageView_("Aucune ligne cochée en colonne A, veuillez sélectionner des activités.", "Information");
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
    openStandaloneMessageView_("Données manquantes ou invalides :\n" + facturerErrors.join("\n"));
    return;
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
    openStandaloneMessageView_("Erreur : Plus de 7 items client/campagne sélectionnés. Maximum autorisé : 7.");
    return;
  }

  const invoiceNumberingSetup = checkInvoiceNumberingSetup();
  if (facturerAlreadyInvoiced.length > 0) {
    showFacturerPopup(facturerContacts, facturerActivityTypes, null, invoiceNumberingSetup.requiresInitialInvoiceSetup, {
      showStartupConfirm: true,
      confirmViewTitle: "Activités déjà facturées",
      confirmViewMessage: "Lignes déjà facturées :\n" + facturerAlreadyInvoiced.join("\n") + "\n\nVoulez-vous continuer ?",
      confirmAction: "facturer_continue",
      confirmPrimaryLabel: "Continuer",
      confirmSecondaryLabel: "Annuler"
    });
    return;
  }

  showFacturerPopup(facturerContacts, facturerActivityTypes, null, invoiceNumberingSetup.requiresInitialInvoiceSetup);
}

function checkInvoiceNumberingSetup() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  if (!facturerTrackingSheet) {
    throw new Error("La feuille 'FACTURATION' est manquante.");
  }
  const facturerTrackingLastRow = facturerTrackingSheet.getLastRow();
  const existingInvoiceValues = facturerTrackingLastRow >= 6
    ? facturerTrackingSheet.getRange(`B6:B${facturerTrackingLastRow}`).getValues().flat().map(value => String(value || "").trim()).filter(String)
    : [];
  return { requiresInitialInvoiceSetup: existingInvoiceValues.length === 0 };
}

function showFacturerPopup(facturerContacts, facturerActivityTypes, invoiceNumber, requiresInitialInvoiceSetup, popupContext = {}) {
  const facturerUi = SpreadsheetApp.getUi();
  const facturerHtml = HtmlService.createHtmlOutputFromFile("popup")
    .setWidth(400)
    .setHeight(350);
  const normalizedPopupContext = Object.assign({
    showStartupConfirm: false,
    confirmViewTitle: "",
    confirmViewMessage: "",
    confirmAction: "",
    confirmPrimaryLabel: "",
    confirmSecondaryLabel: ""
  }, popupContext || {});
  facturerHtml.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  facturerHtml.append(`<script>var contacts = ${JSON.stringify(facturerContacts)}; var activityTypes = ${JSON.stringify(facturerActivityTypes)}; var initialInvoiceNumber = ${JSON.stringify(invoiceNumber)}; var requiresInitialInvoiceSetup = ${JSON.stringify(requiresInitialInvoiceSetup)}; var showStartupConfirm = ${JSON.stringify(normalizedPopupContext.showStartupConfirm)}; var confirmViewTitle = ${JSON.stringify(normalizedPopupContext.confirmViewTitle)}; var confirmViewMessage = ${JSON.stringify(normalizedPopupContext.confirmViewMessage)}; var confirmAction = ${JSON.stringify(normalizedPopupContext.confirmAction)}; var confirmPrimaryLabel = ${JSON.stringify(normalizedPopupContext.confirmPrimaryLabel)}; var confirmSecondaryLabel = ${JSON.stringify(normalizedPopupContext.confirmSecondaryLabel)};</script>`);
  facturerUi.showModelessDialog(facturerHtml, "Nouvelle facture");
}

function validateInvoiceGeneration_() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerConfigSheet = facturerSpreadsheet.getSheetByName("CONFIG");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");
  const facturerModelSheet = facturerSpreadsheet.getSheetByName("MODÈLE");
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const isMissingValue = (value) => value === null || value === undefined || (typeof value === "string" ? value.trim() === "" : value === "");

  if (!facturerTimeSheet || !facturerConfigSheet || !facturerGestionSheet || !facturerModelSheet || !facturerTrackingSheet) {
    return { success: false, message: "Erreur : Feuilles requises manquantes." };
  }

  const facturerLastRow = facturerTimeSheet.getLastRow();
  const facturerTimeData = facturerLastRow >= 7 ? facturerTimeSheet.getRange(`A7:T${facturerLastRow}`).getValues() : [];
  const facturerCheckedRows = facturerTimeData.map((row, index) => ({ row: row, index: index + 7 }))
    .filter(row => row.row[0] === true);

  if (facturerCheckedRows.length === 0) {
    return { success: false, message: "Aucune ligne sélectionnée. Veuillez recommencer." };
  }

  const folderId = String(facturerGestionSheet.getRange("E2").getValue() || "").trim();
  if (!folderId) {
    return { success: false, message: "Dossier Drive invalide. Veuillez vérifier la configuration." };
  }

  let facturerDriveFolder;
  try {
    facturerDriveFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    return { success: false, message: "Dossier Drive invalide. Veuillez vérifier la configuration." };
  }

  const validationCategories = new Set();
  const distinctClients = new Set();

  facturerCheckedRows.forEach(({ row }) => {
    const client = String(row[1] || "").trim();
    distinctClients.add(client);

    if ([row[1], row[2], row[3], row[4], row[19]].some(isMissingValue)) {
      validationCategories.add("required_fields");
    }

    if (isMissingValue(row[7])) {
      validationCategories.add("task_running");
    }

    if (row[14] === true) {
      validationCategories.add("already_invoiced");
    }
  });

  if (distinctClients.size > 1) {
    validationCategories.add("same_client");
  }

  if (validationCategories.size > 1) {
    return { success: false, message: "Entrée non valide. Veuillez recommencer." };
  }

  if (validationCategories.size === 1) {
    const [category] = Array.from(validationCategories);
    const validationMessages = {
      same_client: "Un seul client par facture.",
      required_fields: "Des champs obligatoires sont manquants. Veuillez recommencer.",
      task_running: "Une tâche est encore en cours. Veuillez compléter la tâche puis recommencer.",
      already_invoiced: "Une activité sélectionnée a déjà été facturée. Veuillez recommencer."
    };
    return { success: false, message: validationMessages[category] };
  }

  return {
    success: true,
    facturerSpreadsheet,
    facturerTimeSheet,
    facturerConfigSheet,
    facturerModelSheet,
    facturerTrackingSheet,
    facturerDriveFolder,
    facturerCheckedRows
  };
}

function exportInvoiceSheetPdfBlob_(spreadsheetId, sheetId, fileName) {
  const exportParams = {
    format: "pdf",
    gid: sheetId,
    size: "A4",
    portrait: "true",
    fitw: "true",
    sheetnames: "false",
    printtitle: "false",
    pagenumbers: "false",
    gridlines: "false",
    fzr: "false"
  };
  const exportQuery = Object.keys(exportParams)
    .map(key => `${encodeURIComponent(key)}=${encodeURIComponent(exportParams[key])}`)
    .join("&");
  const exportUrl = `https://docs.google.com/spreadsheets/d/${spreadsheetId}/export?${exportQuery}`;
  const response = UrlFetchApp.fetch(exportUrl, {
    headers: {
      Authorization: `Bearer ${ScriptApp.getOAuthToken()}`
    },
    muteHttpExceptions: true
  });

  if (response.getResponseCode() !== 200) {
    throw new Error("PDF_EXPORT_FAILED");
  }

  return response.getBlob().setName(fileName);
}

function extractDriveFileIdFromUrl_(url) {
  const urlString = String(url || "").trim();
  const match = urlString.match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function sendInvoiceEmail(invoiceNumber, pdfUrl) {
  return sendInvoiceEmail_(invoiceNumber, pdfUrl);
}

function sendInvoiceEmail_(invoiceNumber, pdfUrl) {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const clientsSheet = facturerSpreadsheet.getSheetByName("CLIENTS");

  if (!facturerTimeSheet || !facturerTrackingSheet || !clientsSheet) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const normalizedInvoiceNumber = String(invoiceNumber || "").trim();
  const fileId = extractDriveFileIdFromUrl_(pdfUrl);
  if (!normalizedInvoiceNumber || !fileId) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  let pdfFile;
  try {
    pdfFile = DriveApp.getFileById(fileId);
  } catch (e) {
    Logger.log(`Email attachment lookup failed: ${e.message}`);
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const facturerTrackingLastRow = facturerTrackingSheet.getLastRow();
  const trackingValues = facturerTrackingLastRow >= 6
    ? facturerTrackingSheet.getRange(`A6:I${facturerTrackingLastRow}`).getValues()
    : [];
  const trackingIndex = trackingValues.findIndex(row => String(row[1] || "").trim() === normalizedInvoiceNumber);
  if (trackingIndex === -1) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const trackingRowNumber = trackingIndex + 6;
  const clientName = String(trackingValues[trackingIndex][3] || "").trim();
  if (!clientName) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const clientsLastRow = clientsSheet.getLastRow();
  const clientRows = clientsLastRow >= 2
    ? clientsSheet.getRange(`A2:I${clientsLastRow}`).getValues()
    : [];
  const clientEntry = clientRows.find(row => String(row[0] || "").trim().toLowerCase() === clientName.toLowerCase());
  if (!clientEntry) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const billingContactName = String(clientEntry[5] || "").trim();
  const billingEmail = String(clientEntry[6] || "").trim();
  const ccEmail = String(clientEntry[8] || "").trim();
  if (!billingEmail) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const facturerTimeLastRow = facturerTimeSheet.getLastRow();
  const timeValues = facturerTimeLastRow >= 7
    ? facturerTimeSheet.getRange(`A7:Q${facturerTimeLastRow}`).getValues()
    : [];
  const matchingTimeRows = timeValues
    .map((row, index) => ({ row, index: index + 7 }))
    .filter(entry => String(entry.row[15] || "").trim() === normalizedInvoiceNumber);
  if (matchingTimeRows.length === 0) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const uniquePairs = [];
  const seenPairs = new Set();
  matchingTimeRows.forEach(({ row }) => {
    const campaign = String(row[2] || "").trim();
    const project = String(row[3] || "").trim();
    const pairKey = `${campaign}:::${project}`;
    if (!campaign || !project || seenPairs.has(pairKey)) return;
    seenPairs.add(pairKey);
    uniquePairs.push({ campaign, project });
  });
  if (uniquePairs.length === 0) {
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const recipientName = billingContactName || clientName;
  const emailSubject = `Facture ${normalizedInvoiceNumber} de Mathieu Renaud`;
  const emailBody = [
    `Bonjour ${recipientName},`,
    "",
    `vous trouverez en attachement la facture ${normalizedInvoiceNumber} de Mathieu Renaud couvrant :`,
    "",
    ...uniquePairs.map(pair => `- ${pair.campaign} : ${pair.project}`),
    "",
    "Merci d'accuser la réception de cette facture et de confirmer la conformité des informations détaillées.",
    "",
    "Bien à vous",
    "Mathieu Renaud"
  ].join("\n");

  try {
    const mailOptions = {
      attachments: [pdfFile.getBlob()]
    };
    if (ccEmail) {
      mailOptions.cc = ccEmail;
    }
    MailApp.sendEmail(billingEmail, emailSubject, emailBody, mailOptions);
  } catch (e) {
    Logger.log(`Email send failed: ${e.message}`);
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const todayString = Utilities.formatDate(new Date(), "EDT", "dd MM yyyy");
  matchingTimeRows.forEach(({ index }) => {
    facturerTimeSheet.getRange(`Q${index}`).setValue(todayString);
  });
  facturerTrackingSheet.getRange(`H${trackingRowNumber}`).setValue(todayString);

  return { success: true };
}

function submitFacturerForm(contact, activityType, invoiceNumber, overwriteExistingFile) {
  const extractInvoiceNumberParts = (value) => {
    const invoiceValue = String(value || "").trim();
    if (!/^\d+$/.test(invoiceValue)) return null;
    return {
      numberText: invoiceValue,
      number: Number(invoiceValue)
    };
  };
  const formatInvoiceNumber = (number, padLength) => String(number).padStart(padLength, "0");
  const validationResult = validateInvoiceGeneration_();
  if (!validationResult.success) {
    return validationResult;
  }
  const {
    facturerSpreadsheet,
    facturerTimeSheet,
    facturerConfigSheet,
    facturerModelSheet,
    facturerTrackingSheet,
    facturerDriveFolder,
    facturerCheckedRows
  } = validationResult;

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

  const facturerTrackingLastRow = facturerTrackingSheet.getLastRow();
  const existingInvoiceValues = facturerTrackingLastRow >= 6
    ? facturerTrackingSheet.getRange(`B6:B${facturerTrackingLastRow}`).getValues().flat().map(value => String(value || "").trim()).filter(String)
    : [];
  let facturerFullInvoiceNumber;
  if (existingInvoiceValues.length === 0) {
    if (!Number.isInteger(invoiceNumber) || invoiceNumber <= 0) {
      return { success: false, message: "Le numéro de départ doit être un entier positif." };
    }
    const padLength = Math.max(3, String(invoiceNumber).length);
    facturerFullInvoiceNumber = formatInvoiceNumber(invoiceNumber, padLength);
  } else {
    const parsedInvoices = existingInvoiceValues.map(extractInvoiceNumberParts);
    if (parsedInvoices.some(parts => parts === null || Number.isNaN(parts.number))) {
      return { success: false, message: "Impossible de générer le prochain numéro : FACTURATION!B contient une valeur invalide." };
    }
    const highestInvoiceParts = parsedInvoices.reduce((currentMax, currentInvoice) => {
      return currentInvoice.number > currentMax.number ? currentInvoice : currentMax;
    });
    const facturerNextInvoiceNumber = highestInvoiceParts.number + 1;
    const padLength = Math.max(3, highestInvoiceParts.numberText.length, String(facturerNextInvoiceNumber).length);
    facturerFullInvoiceNumber = formatInvoiceNumber(facturerNextInvoiceNumber, padLength);
  }

  const facturerFileName = `${facturerFullInvoiceNumber}.pdf`;
  const facturerExistingFiles = facturerDriveFolder.getFilesByName(facturerFileName);
  let facturerExistingFile = null;
  if (facturerExistingFiles.hasNext()) {
    facturerExistingFile = facturerExistingFiles.next();
    if (!overwriteExistingFile) {
      return {
        success: false,
        requiresConfirmation: true,
        confirmAction: "replace_pdf",
        confirmTitle: "Remplacer le PDF ?",
        message: `Le fichier ${facturerFileName} existe déjà dans Google Drive. Remplacer ?`,
        confirmPrimaryLabel: "Remplacer",
        confirmSecondaryLabel: "Annuler"
      };
    }
  }

  const facturerTempSheet = facturerModelSheet.copyTo(facturerSpreadsheet).setName(facturerFullInvoiceNumber);

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
  const facturerClientName = String(facturerCheckedRows[0].row[1] || "");
  const facturerScopeSummary = facturerItems
    .map(item => `${item.campaign} (${item.projects.join(", ")})`)
    .join(" | ");

  facturerTempSheet.getRange("L1").setValue(facturerFullInvoiceNumber);
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

  SpreadsheetApp.flush();

  let facturerPdfFile = null;
  try {
    const facturerPdfBlob = exportInvoiceSheetPdfBlob_(facturerSpreadsheet.getId(), facturerTempSheet.getSheetId(), facturerFileName);
    if (facturerExistingFile) {
      facturerPdfFile = facturerDriveFolder.createFile(facturerPdfBlob);
      facturerExistingFile.setTrashed(true);
    } else {
      facturerPdfFile = facturerDriveFolder.createFile(facturerPdfBlob);
    }
    const facturerPdfUrl = facturerPdfFile.getUrl();

    facturerCheckedRows.forEach(row => {
      const facturerRowIndex = row.index;
      facturerTimeSheet.getRange(`A${facturerRowIndex}`).setValue(false);
      facturerTimeSheet.getRange(`O${facturerRowIndex}`).setValue(true);
      facturerTimeSheet.getRange(`P${facturerRowIndex}`).setValue(facturerFullInvoiceNumber);
    });

    const facturerTrackingRow = facturerTrackingSheet.getLastRow() + 1 >= 6 ? facturerTrackingSheet.getLastRow() + 1 : 6;
    facturerTrackingSheet.getRange(`A${facturerTrackingRow}:I${facturerTrackingRow}`).setValues([[
      false,
      facturerFullInvoiceNumber,
      Utilities.formatDate(new Date(), "EDT", "dd MM yyyy"),
      facturerClientName,
      facturerScopeSummary,
      facturerTotalAmount + " $",
      `=HYPERLINK("${facturerPdfUrl}"; "Voir PDF")`,
      "",
      ""
    ]]);
    facturerTrackingSheet.getRange(`B${facturerTrackingRow}`).setNumberFormat("@");

    facturerSpreadsheet.deleteSheet(facturerTempSheet);
    return { success: true, pdfUrl: facturerPdfUrl, invoiceNumber: facturerFullInvoiceNumber };
  } catch (e) {
    Logger.log(`Exception: ${e.message}`);
    if (facturerPdfFile) {
      try {
        facturerPdfFile.setTrashed(true);
      } catch (trashError) {
        Logger.log(`Cleanup exception: ${trashError.message}`);
      }
    }
    facturerSpreadsheet.deleteSheet(facturerTempSheet);
    return { success: false, message: "Erreur lors de la génération du PDF." };
  }
}

// NOUVELLE ENTRÉ DE TEMPS

function newTimeEntry() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");

  if (!facturerTimeSheet || !facturerGestionSheet) {
    openStandaloneMessageView_("Erreur : La feuille 'FEUILLE DE TEMPS' ou 'GESTION' est manquante.");
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
    openStandaloneMessageView_("Erreur : Impossible de lire les taux dans GESTION!C2:C. Valeur par défaut utilisée.", "Information");
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
    openStandaloneMessageView_("Sélectionne une ligne max.", "Information");
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
    throw new Error("Erreur : La feuille 'FEUILLE DE TEMPS' ou 'GESTION' est manquante.");
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
    openStandaloneMessageView_("La feuille 'FEUILLE DE TEMPS' est introuvable.");
  }
}

// SUPPRIMER : Supression des lignes cochées

function trash() {
  const sheet = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const data = sheet.getRange("A7:A").getValues();
  const ligneCocheeCount = data.filter(row => row[0] === true).length;

  const confirmTitle = `Supprimer ${ligneCocheeCount} ligne${ligneCocheeCount > 1 ? 's' : ''}`;
  const confirmMessage = `Confirmer la suppression de ${ligneCocheeCount} ligne${ligneCocheeCount > 1 ? 's' : ''} ?`;
  const output = HtmlService.createHtmlOutputFromFile("popup")
    .setWidth(400)
    .setHeight(220);
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  output.append(`<script>var contacts = []; var activityTypes = []; var initialInvoiceNumber = null; var requiresInitialInvoiceSetup = false; var showStartupConfirm = true; var confirmViewTitle = ${JSON.stringify(confirmTitle)}; var confirmViewMessage = ${JSON.stringify(confirmMessage)}; var confirmAction = "delete_rows"; var confirmPrimaryLabel = "Supprimer"; var confirmSecondaryLabel = "Annuler";</script>`);

  SpreadsheetApp.getUi().showModalDialog(output, confirmTitle);
}

function supprimerLignesCochées() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheet = ss.getActiveSheet();

  if (sheet.getName() !== "FEUILLE DE TEMPS") {
    throw new Error("Cette fonction ne peut être utilisée que dans la feuille 'FEUILLE DE TEMPS'.");
  }

  const data = sheet.getRange("A1:A" + sheet.getLastRow()).getValues();
  const lignesASupprimer = [];

  for (let i = data.length - 1; i >= 0; i--) {
    if (data[i][0] === true) {
      lignesASupprimer.push(i + 1);
    }
  }

  if (lignesASupprimer.length === 0) {
    throw new Error("Aucune ligne cochée à supprimer.");
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
    openStandaloneMessageView_("Erreur : La feuille FACTURATION n'existe pas.");
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
    openStandaloneMessageView_("Erreur : La feuille 'FEUILLE DE TEMPS' est manquante.");
    return;
  }

  // Obtenir les cases cochées dans la colonne A à partir de A7
  const lastRow = sheetTime.getLastRow();
  if (lastRow < 7) {
    openStandaloneMessageView_("Erreur : Aucune donnée à partir de la ligne 7.");
    return;
  }

  const checkBoxData = sheetTime.getRange(`A7:A${lastRow}`).getValues();
  const checkedRows = checkBoxData
    .map((row, index) => ({ checked: row[0], index: index + 7 }))
    .filter(row => row.checked === true);

  // Vérifier si exactement une ligne est cochée
  if (checkedRows.length > 1) {
    openStandaloneMessageView_("Ne cocher qu'une seule ligne.", "Information");
    return;
  }

  if (checkedRows.length === 0) {
    openStandaloneMessageView_("Erreur : Aucune ligne cochée.");
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
      openStandaloneMessageView_("La cellule H de la ligne cochée contient déjà une valeur.", "Action impossible sur cette ligne");
    }
  }
}

// INFO : Change les données sur MODÈLE

function info() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetModele = ss.getSheetByName("MODÈLE");
  const sheetGestion = ss.getSheetByName("GESTION");

  if (!sheetModele || !sheetGestion) {
    openStandaloneMessageView_("Erreur : La feuille 'MODÈLE' ou 'GESTION' est manquante.");
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

  const nextInvoice = String(sheetGestion.getRange("A2").getValue() || "");

  const html = HtmlService.createTemplateFromFile("popupInfo");
  html.name = name;
  html.address1 = address1;
  html.address2 = address2;
  html.address3 = address3;
  html.address4 = address4;
  html.email = email;
  html.website = website;
  html.nextInvoice = nextInvoice;

  const htmlOutput = html.evaluate().setWidth(600).setHeight(450);
  SpreadsheetApp.getUi().showModelessDialog(
    htmlOutput,
    "Configuration de la facture"
  );
}

function submitInfoForm(name, address1, address2, address3, address4, email, website, nextInvoice) {
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
}

// DOSSIER : Ouvre le dossier avec les PDF

function dossier() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetGestion = ss.getSheetByName("GESTION");

  if (!sheetGestion) {
    openStandaloneMessageView_("Erreur : La feuille 'GESTION' est manquante.");
    return;
  }

  const folderId = String(sheetGestion.getRange("E2").getValue() || "");
  if (!folderId) {
    openStandaloneMessageView_("Erreur : Aucun dossier configuré.");
    return;
  }

  try {
    const folder = DriveApp.getFolderById(folderId);
    const url = folder.getUrl();
    const html = `<script>window.open('${url}', '_blank'); google.script.host.close();</script>`;
    SpreadsheetApp.getUi().showModelessDialog(HtmlService.createHtmlOutput(html).setWidth(1).setHeight(1), "Ouvrir le dossier");
  } catch (e) {
    openStandaloneMessageView_("Erreur : ID de dossier invalide.");
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
