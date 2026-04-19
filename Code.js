// GÉNÉRAL

function openStandaloneMessageView_(message, title = "Erreur", options = {}) {
  const { width = 900, height = 450 } = options;
  const template = HtmlService.createTemplateFromFile("popup");
  template.showStandaloneMessage = true;
  template.messageViewTitle = title;
  template.messageViewMessage = message;
  template.facturationPopupContextReady = true;
  const html = template.evaluate()
    .setWidth(width)
    .setHeight(height);
  html.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  SpreadsheetApp.getUi().showModelessDialog(html, "Facturation");
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
  const facturerModelSheet = facturerSpreadsheet.getSheetByName("MODÈLE");
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");

  if (!facturerTimeSheet || !facturerModelSheet || !facturerTrackingSheet || !facturerGestionSheet) {
    openStandaloneMessageView_("Erreur : Une ou plusieurs feuilles nécessaires sont manquantes.");
    return;
  }

  if (!isInvoiceCompanyConfigured_(facturerGestionSheet)) {
    showInvoiceConfigurationDialog_({ continueFacturerAfterSave: true });
    return;
  }

  openValidatedFacturerFlow_(null);
}

function openValidatedFacturerFlow_(initialInvoiceNumber) {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerModelSheet = facturerSpreadsheet.getSheetByName("MODÈLE");
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");

  if (!facturerTimeSheet || !facturerModelSheet || !facturerTrackingSheet || !facturerGestionSheet) {
    openStandaloneMessageView_("Erreur : Une ou plusieurs feuilles nécessaires sont manquantes.");
    return { success: false, message: "Erreur : Une ou plusieurs feuilles nécessaires sont manquantes." };
  }

  const driveFolderValidation = validateDriveFolderId_(facturerGestionSheet);
  if (!driveFolderValidation.success) {
    // Étape dédiée: demander le dossier Drive et reprendre automatiquement après Enregistrer.
    showDriveFolderConfigurationDialog_({
      continueFacturerAfterSave: true,
      initialInvoiceNumber: initialInvoiceNumber
    });
    // Important: ne pas remonter "success:false" sinon la fenêtre de configuration entreprise
    // resterait ouverte et afficherait une erreur, alors qu'on veut enchaîner sur l'étape dossier.
    return { success: true, deferred: true };
  }

  const facturerTimeData = facturerTimeSheet.getRange("A7:Q" + facturerTimeSheet.getLastRow()).getValues();
  const facturerCheckedRows = facturerTimeData.map((row, index) => ({ row: row, index: index + 7 }))
    .filter(row => row.row[0] === true);

  if (facturerCheckedRows.length === 0) {
    openStandaloneMessageView_("Aucune ligne cochée en colonne A, veuillez sélectionner des activités.", "Information");
    return { success: false, message: "Aucune ligne cochée en colonne A, veuillez sélectionner des activités." };
  }

  const preliminaryValidation = validateInvoiceGeneration_();
  if (!preliminaryValidation.success) {
    const validationTitle = preliminaryValidation.message.startsWith("Attention") || preliminaryValidation.message === "Un seul client par facture."
      ? "Attention"
      : "Information";
    openStandaloneMessageView_(preliminaryValidation.message, validationTitle);
    return preliminaryValidation;
  }

  const invoiceNumberingSetup = checkInvoiceNumberingSetup();
  const requiresInitialInvoiceSetup = initialInvoiceNumber === null && invoiceNumberingSetup.requiresInitialInvoiceSetup;
  showFacturerPopup([], [], initialInvoiceNumber, requiresInitialInvoiceSetup);
  return { success: true };
}

function extractDriveFolderIdFromInput_(value) {
  // Accepte un ID collé directement ou une URL Drive contenant un ID.
  // On utilise la même règle simple que pour les liens Drive ailleurs dans le projet.
  const urlString = String(value || "").trim();
  const match = urlString.match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function validateDriveFolderId_(sheetGestion) {
  const folderId = String(sheetGestion.getRange("E2").getValue() || "").trim();
  if (!folderId) {
    return { success: false, message: "Erreur : Aucun dossier configuré." };
  }

  try {
    DriveApp.getFolderById(folderId);
  } catch (e) {
    return { success: false, message: "Erreur : ID de dossier invalide." };
  }

  return { success: true, folderId };
}

function showDriveFolderConfigurationDialog_(options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetGestion = ss.getSheetByName("GESTION");
  if (!sheetGestion) {
    openStandaloneMessageView_("Erreur : La feuille 'GESTION' est manquante.");
    return;
  }

  const html = HtmlService.createTemplateFromFile("popupFolder");
  html.initialInvoiceNumber = typeof options.initialInvoiceNumber === "undefined" ? null : options.initialInvoiceNumber;
  html.continueFacturerAfterSave = Boolean(options.continueFacturerAfterSave);

  const htmlOutput = html.evaluate().setWidth(900).setHeight(450);
  SpreadsheetApp.getUi().showModelessDialog(
    htmlOutput,
    "Dossier pour factures"
  );
}

function validateDriveFolderInput(rawValue) {
  const folderId = extractDriveFolderIdFromInput_(rawValue);
  if (!folderId) {
    return { success: false, folderId: "" };
  }

  try {
    DriveApp.getFolderById(folderId);
  } catch (e) {
    return { success: false, folderId: "" };
  }

  return { success: true, folderId };
}

function submitDriveFolderForm(rawValue, initialInvoiceNumber) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetGestion = ss.getSheetByName("GESTION");
  if (!sheetGestion) {
    return { success: false, message: "Erreur : La feuille GESTION est manquante." };
  }

  const validation = validateDriveFolderInput(rawValue);
  if (!validation.success) {
    return { success: false, message: "Erreur : ID de dossier invalide." };
  }

  // Stockage: uniquement l'ID, même si l'utilisateur colle une URL.
  sheetGestion.getRange("E2").setValue(validation.folderId);

  // Reprendre automatiquement le flux de facturation là où il s'était arrêté.
  // initialInvoiceNumber peut être null (cas classique) ou un numéro issu de la configuration entreprise.
  return openValidatedFacturerFlow_(typeof initialInvoiceNumber === "undefined" ? null : initialInvoiceNumber);
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

function normalizeInvoiceNumberInput_(invoiceNumber) {
  if (invoiceNumber === null || invoiceNumber === undefined || invoiceNumber === "") {
    return null;
  }
  if (typeof invoiceNumber === "number") {
    return Number.isInteger(invoiceNumber) ? invoiceNumber : NaN;
  }
  const invoiceNumberString = String(invoiceNumber || "").trim();
  if (!/^\d+$/.test(invoiceNumberString)) {
    return NaN;
  }
  return Number(invoiceNumberString);
}

function extractInvoiceNumberParts_(value) {
  const invoiceValue = String(value || "").trim();
  if (!/^\d+$/.test(invoiceValue)) return null;
  return {
    numberText: invoiceValue,
    number: Number(invoiceValue)
  };
}

function formatInvoiceNumber_(number, padLength) {
  return String(number).padStart(padLength, "0");
}

function normalizeInvoiceNumberForComparison_(value) {
  const trimmed = String(value || "").trim();
  if (!trimmed) return "";
  // Ne garder que les numéros, puis enlever les zéros à gauche.
  // "001" et "1" doivent être équivalents pour les comparaisons historiques.
  if (!/^\d+$/.test(trimmed)) return "";
  return String(Number(trimmed));
}

function getExistingInvoiceValues_(facturerTrackingSheet) {
  const facturerTrackingLastRow = facturerTrackingSheet.getLastRow();
  return facturerTrackingLastRow >= 6
    ? facturerTrackingSheet.getRange(`B6:B${facturerTrackingLastRow}`).getValues().flat().map(value => String(value || "").trim()).filter(String)
    : [];
}

function resolveInvoiceNumberForFacturation_(facturerTrackingSheet, invoiceNumber) {
  const normalizedInvoiceNumber = normalizeInvoiceNumberInput_(invoiceNumber);
  if (Number.isNaN(normalizedInvoiceNumber) || (normalizedInvoiceNumber !== null && normalizedInvoiceNumber <= 0)) {
    return { success: false, message: "Le numéro de facture doit être un entier positif." };
  }

  const existingInvoiceValues = getExistingInvoiceValues_(facturerTrackingSheet);
  const parsedInvoices = existingInvoiceValues.map(extractInvoiceNumberParts_);
  if (parsedInvoices.some(parts => parts === null || Number.isNaN(parts.number))) {
    return { success: false, message: "Impossible de générer le prochain numéro : FACTURATION!B contient une valeur invalide." };
  }

  if (normalizedInvoiceNumber !== null) {
    if (parsedInvoices.some(parts => parts.number === normalizedInvoiceNumber)) {
      return { success: false, message: "Ce numéro de facture existe déjà." };
    }
    const existingPadLength = parsedInvoices.reduce((maxLength, parts) => Math.max(maxLength, parts.numberText.length), 3);
    const padLength = Math.max(existingPadLength, String(normalizedInvoiceNumber).length);
    return {
      success: true,
      invoiceNumber: normalizedInvoiceNumber,
      fullInvoiceNumber: formatInvoiceNumber_(normalizedInvoiceNumber, padLength)
    };
  }

  if (existingInvoiceValues.length === 0) {
    return { success: false, message: "Le numéro de départ doit être un entier positif." };
  }

  const highestInvoiceParts = parsedInvoices.reduce((currentMax, currentInvoice) => {
    return currentInvoice.number > currentMax.number ? currentInvoice : currentMax;
  });
  const facturerNextInvoiceNumber = highestInvoiceParts.number + 1;
  const padLength = Math.max(3, highestInvoiceParts.numberText.length, String(facturerNextInvoiceNumber).length);
  return {
    success: true,
    invoiceNumber: facturerNextInvoiceNumber,
    fullInvoiceNumber: formatInvoiceNumber_(facturerNextInvoiceNumber, padLength)
  };
}

function getSuggestedNextInvoiceNumber_(facturerTrackingSheet) {
  const existingInvoiceValues = getExistingInvoiceValues_(facturerTrackingSheet);
  if (existingInvoiceValues.length === 0) {
    return "";
  }
  const parsedInvoices = existingInvoiceValues.map(extractInvoiceNumberParts_);
  if (parsedInvoices.some(parts => parts === null || Number.isNaN(parts.number))) {
    return "";
  }
  const highestInvoiceParts = parsedInvoices.reduce((currentMax, currentInvoice) => {
    return currentInvoice.number > currentMax.number ? currentInvoice : currentMax;
  });
  const nextInvoiceNumber = highestInvoiceParts.number + 1;
  const padLength = Math.max(3, highestInvoiceParts.numberText.length, String(nextInvoiceNumber).length);
  return formatInvoiceNumber_(nextInvoiceNumber, padLength);
}

function getExistingInvoiceNumberList_(facturerTrackingSheet) {
  return getExistingInvoiceValues_(facturerTrackingSheet)
    .map(extractInvoiceNumberParts_)
    .filter(parts => parts && !Number.isNaN(parts.number))
    .map(parts => parts.number);
}

function getCompanyInfoFromGestion_(sheetGestion) {
  const companyInfo = sheetGestion.getRange("F2:F8").getValues().flat();
  return {
    name: String(companyInfo[0] || ""),
    address1: String(companyInfo[1] || ""),
    address2: String(companyInfo[2] || ""),
    address3: String(companyInfo[3] || ""),
    address4: String(companyInfo[4] || ""),
    email: String(companyInfo[5] || ""),
    website: String(companyInfo[6] || "")
  };
}

function isInvoiceCompanyConfigured_(sheetGestion) {
  const companyInfo = getCompanyInfoFromGestion_(sheetGestion);
  return [companyInfo.name, companyInfo.email, companyInfo.address1, companyInfo.address2]
    .every(value => String(value || "").trim() !== "");
}

function showInvoiceConfigurationDialog_(options = {}) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetGestion = ss.getSheetByName("GESTION");
  const facturerTrackingSheet = ss.getSheetByName("FACTURATION");

  if (!sheetGestion || !facturerTrackingSheet) {
    openStandaloneMessageView_("Erreur : Une ou plusieurs feuilles nécessaires sont manquantes.");
    return;
  }

  const companyInfo = getCompanyInfoFromGestion_(sheetGestion);
  const html = HtmlService.createTemplateFromFile("popupInfo");
  html.name = companyInfo.name;
  html.address1 = companyInfo.address1;
  html.address2 = companyInfo.address2;
  html.address3 = companyInfo.address3;
  html.address4 = companyInfo.address4;
  html.email = companyInfo.email;
  html.website = companyInfo.website;
  html.nextInvoice = getSuggestedNextInvoiceNumber_(facturerTrackingSheet);
  html.existingInvoiceNumbers = getExistingInvoiceNumberList_(facturerTrackingSheet);
  html.continueFacturerAfterSave = Boolean(options.continueFacturerAfterSave);

  const htmlOutput = html.evaluate().setWidth(900).setHeight(450);
  SpreadsheetApp.getUi().showModelessDialog(
    htmlOutput,
    "Configuration de la facture"
  );
}

function continueFacturerAfterConfiguration(nextInvoice) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTrackingSheet = ss.getSheetByName("FACTURATION");
  if (!facturerTrackingSheet) {
    return { success: false, message: "Erreur : La feuille FACTURATION est manquante." };
  }

  const invoiceNumberResolution = resolveInvoiceNumberForFacturation_(facturerTrackingSheet, nextInvoice);
  if (!invoiceNumberResolution.success) {
    return invoiceNumberResolution;
  }

  return openValidatedFacturerFlow_(invoiceNumberResolution.invoiceNumber);
}

function showFacturerPopup(facturerContacts, facturerActivityTypes, invoiceNumber, requiresInitialInvoiceSetup, popupContext = {}) {
  const facturerUi = SpreadsheetApp.getUi();
  const normalizedPopupContext = Object.assign({
    showStartupConfirm: false,
    confirmViewTitle: "",
    confirmViewMessage: "",
    confirmAction: "",
    confirmPrimaryLabel: "",
    confirmSecondaryLabel: ""
  }, popupContext || {});
  const template = HtmlService.createTemplateFromFile("popup");
  template.contacts = facturerContacts;
  template.activityTypes = facturerActivityTypes;
  template.initialInvoiceNumber = invoiceNumber;
  template.requiresInitialInvoiceSetup = requiresInitialInvoiceSetup;
  template.showStartupConfirm = normalizedPopupContext.showStartupConfirm;
  template.confirmViewTitle = normalizedPopupContext.confirmViewTitle;
  template.confirmViewMessage = normalizedPopupContext.confirmViewMessage;
  template.confirmAction = normalizedPopupContext.confirmAction;
  template.confirmPrimaryLabel = normalizedPopupContext.confirmPrimaryLabel;
  template.confirmSecondaryLabel = normalizedPopupContext.confirmSecondaryLabel;
  template.facturationPopupContextReady = true;
  const facturerHtml = template.evaluate()
    .setWidth(900)
    .setHeight(450);
  facturerHtml.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  facturerUi.showModelessDialog(facturerHtml, "Facturation");
}

function validateInvoiceGeneration_() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");
  const facturerModelSheet = facturerSpreadsheet.getSheetByName("MODÈLE");
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const isMissingValue = (value) => value === null || value === undefined || (typeof value === "string" ? value.trim() === "" : value === "");

  if (!facturerTimeSheet || !facturerGestionSheet || !facturerModelSheet || !facturerTrackingSheet) {
    return { success: false, message: "Erreur : Feuilles requises manquantes." };
  }

  const facturerLastRow = facturerTimeSheet.getLastRow();
  const facturerTimeData = facturerLastRow >= 7 ? facturerTimeSheet.getRange(`A7:U${facturerLastRow}`).getValues() : [];
  const facturerCheckedRows = facturerTimeData.map((row, index) => ({ row: row, index: index + 7 }))
    .filter(row => row.row[0] === true);

  if (facturerCheckedRows.length === 0) {
    return { success: false, message: "Aucune ligne sélectionnée. Veuillez recommencer." };
  }

  const folderId = String(facturerGestionSheet.getRange("E2").getValue() || "").trim();
  if (!folderId) {
    return { success: false, message: "Erreur : Aucun dossier configuré." };
  }

  let facturerDriveFolder;
  try {
    facturerDriveFolder = DriveApp.getFolderById(folderId);
  } catch (e) {
    return { success: false, message: "Erreur : ID de dossier invalide." };
  }

  const validationCategories = new Set();
  const distinctClients = new Set();
  const toDecimalHours = (value) => {
    if (value instanceof Date) {
      return value.getHours() + value.getMinutes() / 60 + value.getSeconds() / 3600;
    }
    if (typeof value === "number") {
      return value;
    }
    if (typeof value === "string" && value.trim() !== "" && !isNaN(Number(value))) {
      return Number(value);
    }
    return NaN;
  };

  facturerCheckedRows.forEach(({ row }) => {
    const client = String(row[1] || "").trim();
    if (client) {
      distinctClients.add(client);
    }

    if ([row[1], row[2], row[3], row[4]].some(isMissingValue)) {
      validationCategories.add("required_fields");
    }

    if (isMissingValue(row[7])) {
      validationCategories.add("task_running");
    }

    const rateValue = Number(row[19]);
    if (isMissingValue(row[19]) || !isFinite(rateValue) || rateValue <= 0) {
      validationCategories.add("invalid_rate");
    }

    const timeValue = toDecimalHours(row[8]);
    if (isMissingValue(row[8]) || !isFinite(timeValue) || timeValue <= 0) {
      validationCategories.add("invalid_time");
    }

    if (row[14] === true) {
      validationCategories.add("already_invoiced");
    }

    if (row[17] === true || !isMissingValue(row[18])) {
      validationCategories.add("already_paid");
    }
  });

  if (distinctClients.size > 1) {
    validationCategories.add("same_client");
  }

  const validationPriority = [
    "same_client",
    "task_running",
    "required_fields",
    "invalid_rate",
    "already_invoiced",
    "already_paid",
    "invalid_time"
  ];
  const validationMessages = {
    same_client: "Un seul client par facture.",
    required_fields: "Attention. Données incomplètes.",
    invalid_rate: "Attention, taux invalide. Veuillez recommencer.",
    already_invoiced: "Attention, certaines entrées ont déjà été facturées.",
    already_paid: "Attention. Certaines entrées ont déjà été payées.",
    task_running: "Attention, tâche en cours.",
    invalid_time: "Attention, heure invalide."
  };

  for (const category of validationPriority) {
    if (validationCategories.has(category)) {
      return { success: false, message: validationMessages[category] };
    }
  }

  return {
    success: true,
    facturerSpreadsheet,
    facturerTimeSheet,
    facturerGestionSheet,
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

  const responseCode = response.getResponseCode();
  if (responseCode !== 200) {
    const responseBody = response.getContentText();
    Logger.log(`PDF export failed with HTTP ${responseCode}`);
    Logger.log(`PDF export response body: ${responseBody}`);
    throw new Error(`PDF_EXPORT_FAILED: HTTP ${responseCode} - ${responseBody}`);
  }

  return response.getBlob().setName(fileName);
}

function extractDriveFileIdFromUrl_(url) {
  const urlString = String(url || "").trim();
  const match = urlString.match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function normalizeString_(str) {
  return String(str || "")
    .toLowerCase()
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .replace(/\s+/g, " ")
    .trim();
}

function logSendInvoiceEmailError_(step, error) {
  Logger.log(JSON.stringify({
    context: "sendInvoiceEmail",
    step: step,
    error: error
  }));
}

function sendInvoiceEmail(invoiceNumber, pdfUrl, recipient, ccRecipient) {
  return sendInvoiceEmail_(invoiceNumber, pdfUrl, recipient, ccRecipient);
}

function cleanGeneratedInvoiceText_(value) {
  return String(value || "").replace(/\s+/g, " ").trim();
}

function extractOpenAIInvoiceResponseText_(responseData) {
  if (responseData && typeof responseData.output_text === "string") {
    return responseData.output_text.trim();
  }

  const outputItems = responseData && Array.isArray(responseData.output) ? responseData.output : [];
  for (const outputItem of outputItems) {
    const contentItems = Array.isArray(outputItem.content) ? outputItem.content : [];
    for (const contentItem of contentItems) {
      if (typeof contentItem.text === "string" && contentItem.text.trim()) {
        return contentItem.text.trim();
      }
    }
  }

  return "";
}

function parseOpenAIInvoiceText_(rawText) {
  const jsonText = String(rawText || "")
    .trim()
    .replace(/^```(?:json)?\s*/i, "")
    .replace(/\s*```$/i, "");
  if (!jsonText) {
    return null;
  }

  const parsed = JSON.parse(jsonText);
  if (!parsed || typeof parsed !== "object") {
    return null;
  }

  const serviceTitle = cleanGeneratedInvoiceText_(parsed.serviceTitle);
  const blocks = Array.isArray(parsed.blocks)
    ? parsed.blocks.map(block => ({
      blockNumber: Number(block.blockNumber),
      campaign: cleanGeneratedInvoiceText_(block.campaign),
      project: cleanGeneratedInvoiceText_(block.project),
      description: cleanGeneratedInvoiceText_(block.description)
    })).filter(block => block.description)
    : [];

  if (!serviceTitle && blocks.length === 0) {
    return null;
  }

  return { serviceTitle, blocks };
}

function generateInvoiceTextWithOpenAI(previewPayload) {
  const apiKey = PropertiesService.getScriptProperties().getProperty("OPENAI_API_KEY");
  if (!apiKey) {
    Logger.log("OpenAI invoice preview skipped: missing OPENAI_API_KEY. Local fallback used.");
    return null;
  }

  const requestPayload = {
    model: "gpt-4o-mini",
    instructions: [
      "Tu aides a preparer un apercu de facture pour un client.",
      "Retourne uniquement du JSON valide, sans Markdown ni texte hors JSON.",
      "Le champ serviceTitle doit etre un titre nominal, court, precis, idealement entre 2 et 5 mots, qui designe le mandat ou le livrable principal.",
      "Le serviceTitle ne doit jamais commencer par 'Facture pour', ne doit jamais contenir 'Services de', ne doit pas etre une phrase complete et ne doit pas repeter qu'il s'agit d'une facture.",
      "Analyse l'ensemble des blocs avant de choisir serviceTitle: si plusieurs types de projets ou livrables sont presents, produis un titre global et englobant qui represente l'ensemble, sans te limiter a un seul type.",
      "Ne liste pas les projets dans serviceTitle, ne concatene pas les livrables et produis un intitule synthetique representatif.",
      "Style vise pour serviceTitle, sans copier ces exemples: Conception de materiel promotionnel, Creation de visuels promotionnels, Conception de supports visuels.",
      "Chaque blocks[].description doit etre une seule phrase fluide, idealement entre 20 et 35 mots, professionnelle et agreable a lire.",
      "Chaque description doit commencer par une action principale comme Creation, Developpement ou Conception, puis resumer le travail en grandes phases.",
      "Regroupe les etapes en processus global: mise en place ou structuration, developpement visuel, puis ajustements ou livraison.",
      "Conserve explicitement une ou deux composantes importantes du mandat quand elles structurent le travail, comme illustrations, declinaisons, adaptations visuelles ou ajustements finaux.",
      "Ne compresse pas au point d'effacer un element significatif, mais integre ces composantes dans une phrase naturelle plutot que dans une enumeration.",
      "Utilise au besoin une structure fluide comme depuis ... jusqu'a ..., incluant ... et ..., ou avec ... et ..., sans alourdir la phrase.",
      "Ne detaille pas chaque etape, ne reprends pas les notes une a une, ne produis pas de liste separee par des virgules et n'ecris pas un inventaire.",
      "N'utilise pas de phrases commencant par Apres ou Les fichiers ont ete, et evite les textes froids, telegraphiques, promotionnels ou exageres."
    ].join(" "),
    input: JSON.stringify(previewPayload),
    text: {
      format: {
        type: "json_schema",
        name: "invoice_preview_text",
        strict: true,
        schema: {
          type: "object",
          additionalProperties: false,
          properties: {
            serviceTitle: { type: "string" },
            blocks: {
              type: "array",
              items: {
                type: "object",
                additionalProperties: false,
                properties: {
                  blockNumber: { type: "number" },
                  campaign: { type: "string" },
                  project: { type: "string" },
                  description: { type: "string" }
                },
                required: ["blockNumber", "campaign", "project", "description"]
              }
            }
          },
          required: ["serviceTitle", "blocks"]
        }
      }
    },
    temperature: 0.2,
    max_output_tokens: 1200
  };

  try {
    const response = UrlFetchApp.fetch("https://api.openai.com/v1/responses", {
      method: "post",
      contentType: "application/json",
      headers: {
        Authorization: `Bearer ${apiKey}`
      },
      payload: JSON.stringify(requestPayload),
      muteHttpExceptions: true
    });
    const responseCode = response.getResponseCode();
    if (responseCode < 200 || responseCode >= 300) {
      Logger.log(`OpenAI invoice preview failed: HTTP ${responseCode}. Local fallback used.`);
      return null;
    }

    const responseData = JSON.parse(response.getContentText());
    const outputText = extractOpenAIInvoiceResponseText_(responseData);
    const generatedText = parseOpenAIInvoiceText_(outputText);
    if (!generatedText) {
      Logger.log("OpenAI invoice preview failed: invalid or empty JSON. Local fallback used.");
      return null;
    }

    return generatedText;
  } catch (error) {
    Logger.log(`OpenAI invoice preview exception: ${error && error.message ? error.message : error}. Local fallback used.`);
    return null;
  }
}

function splitTextIntoLines(text, maxChars) {
  const normalizedText = String(text || "").replace(/\s+/g, " ").trim();
  if (!normalizedText) {
    return [""];
  }

  const words = normalizedText.split(" ");
  const lines = [];
  let currentLine = "";

  words.forEach(word => {
    if (!currentLine) {
      currentLine = word;
      return;
    }

    const nextLine = `${currentLine} ${word}`;
    if (nextLine.length <= maxChars) {
      currentLine = nextLine;
    } else {
      lines.push(currentLine);
      currentLine = word;
    }
  });

  if (currentLine) {
    lines.push(currentLine);
  }

  return lines;
}

function formatFrenchInvoiceList_(items) {
  const values = Array.isArray(items)
    ? items.map(item => String(item || "").trim()).filter(String)
    : [];
  const displayValues = values.map((value, index) => {
    if (index === 0) return value;
    return value.charAt(0).toLocaleLowerCase("fr-CA") + value.slice(1);
  });
  if (displayValues.length <= 2) {
    return displayValues.join(" et ");
  }
  return `${displayValues.slice(0, -1).join(", ")} et ${displayValues[displayValues.length - 1]}`;
}

function buildFixedInvoiceBlocks_(facturerCheckedRows) {
  const toDecimalHours = (value) => {
    if (value instanceof Date) {
      return value.getHours() + value.getMinutes() / 60 + value.getSeconds() / 3600;
    }
    return Number(value) || 0;
  };
  const cleanText = (value) => String(value || "").replace(/\s+/g, " ").trim();
  const blocks = [];

  facturerCheckedRows.forEach(({ row, index }) => {
    const campaign = cleanText(row[2]);
    const project = cleanText(row[3]);
    const blockKey = `${normalizeString_(campaign)}|||${normalizeString_(project)}`;
    let block = blocks.find(candidate => candidate.key === blockKey);
    if (!block) {
      block = {
        key: blockKey,
        campaign,
        project,
        notes: [],
        activities: [],
        rowIndexes: [],
        totalTime: 0,
        totalPrice: 0
      };
      blocks.push(block);
    }
    block.rowIndexes.push(index);

    const activityName = cleanText(row[4]);
    const time = toDecimalHours(row[8]);
    const price = Number(row[10]) || 0;
    const existingActivity = block.activities.find(activity => normalizeString_(activity.name) === normalizeString_(activityName));
    if (existingActivity) {
      existingActivity.time += time;
    } else {
      block.activities.push({ name: activityName, time });
    }
    block.totalTime += time;
    block.totalPrice += price;

    const note = cleanText(row[20]);
    if (note && !block.notes.some(existingNote => normalizeString_(existingNote) === normalizeString_(note))) {
      block.notes.push(note);
    }
  });

  return blocks.map((block, index) => ({
    blockNumber: index + 1,
    campaign: block.campaign,
    project: block.project,
    title: [block.campaign, block.project].filter(Boolean).join(" — "),
    description: buildFixedInvoiceBlockDescription_(block),
    activities: block.activities,
    rowIndexes: block.rowIndexes.slice(),
    totalTime: block.totalTime,
    totalPrice: block.totalPrice
  }));
}

function filterFacturerCheckedRowsByInvoicePreview_(facturerCheckedRows, invoicePreview) {
  const previewBlocks = invoicePreview && Array.isArray(invoicePreview.blocks) ? invoicePreview.blocks : [];
  if (!previewBlocks.length) {
    return facturerCheckedRows;
  }

  const previewRowIndexes = new Set();
  previewBlocks.forEach(block => {
    if (!Array.isArray(block.rowIndexes)) return;
    block.rowIndexes.forEach(rowIndex => {
      const numericRowIndex = Number(rowIndex);
      if (Number.isInteger(numericRowIndex)) {
        previewRowIndexes.add(numericRowIndex);
      }
    });
  });
  if (previewRowIndexes.size) {
    return facturerCheckedRows.filter(entry => previewRowIndexes.has(entry.index));
  }

  const previewBlockKeys = new Set(previewBlocks.map(block => {
    const campaign = String(block.campaign || "").trim();
    const project = String(block.project || "").trim();
    return `${normalizeString_(campaign)}|||${normalizeString_(project)}`;
  }));
  return previewBlockKeys.size
    ? facturerCheckedRows.filter(({ row }) => {
      const campaign = String(row[2] || "").trim();
      const project = String(row[3] || "").trim();
      return previewBlockKeys.has(`${normalizeString_(campaign)}|||${normalizeString_(project)}`);
    })
    : facturerCheckedRows;
}

function applyValidatedInvoicePreviewToBlocks_(blocks, invoicePreview) {
  const previewBlocks = invoicePreview && Array.isArray(invoicePreview.blocks) ? invoicePreview.blocks : [];
  if (!previewBlocks.length) {
    return blocks;
  }

  return blocks.map((block, index) => {
    const previewBlock = previewBlocks.find(candidate => Number(candidate.blockNumber) === block.blockNumber)
      || previewBlocks.find(candidate => normalizeString_(candidate.campaign) === normalizeString_(block.campaign)
        && normalizeString_(candidate.project) === normalizeString_(block.project))
      || previewBlocks[index];
    const validatedDescription = previewBlock ? cleanGeneratedInvoiceText_(previewBlock.description) : "";
    return validatedDescription
      ? Object.assign({}, block, { description: validatedDescription })
      : block;
  });
}

function buildFixedInvoiceBlockDescription_(block) {
  const notes = Array.isArray(block.notes) ? block.notes : [];
  if (notes.length) {
    return notes.join("; ");
  }

  const activityNames = Array.isArray(block.activities)
    ? block.activities.map(activity => String(activity.name || "").trim()).filter(String)
    : [];
  if (activityNames.length) {
    return `Travaux réalisés : ${activityNames.join(", ")}.`;
  }

  const context = [block.campaign, block.project].filter(Boolean).join(" — ");
  return context ? `Travaux réalisés pour ${context}.` : "Travaux réalisés.";
}

function buildFixedInvoiceLayoutRows_(blocks) {
  const rows = [];

  blocks.forEach((block, blockIndex) => {
    const titleLines = splitTextIntoLines(block.title, 75);
    const descriptionLines = splitTextIntoLines(block.description, 85);

    titleLines.forEach((line, lineIndex) => {
      rows.push({
        type: "title",
        height: 25,
        block,
        text: line,
        isFirstTitleLine: lineIndex === 0
      });
    });

    rows.push({ type: "title-space", height: 5 });

    descriptionLines.forEach(line => {
      rows.push({
        type: "description",
        height: 21,
        text: line
      });
    });

    rows.push({ type: "description-space", height: 15 });

    block.activities.forEach(activity => {
      rows.push({
        type: "activity",
        height: 21,
        activity
      });
    });

    if (blockIndex < blocks.length - 1) {
      rows.push({ type: "separator", height: 20 });
    }
  });

  return rows;
}

function writeFixedInvoiceBlocks_(sheet, blocks) {
  const startRow = 21;
  const contentRowCount = 28;
  const bufferRow = 49;
  const minimumRowHeight = 1;
  const bufferRowExtensionPx = 40;
  const targetHeight = 640; //625 605
  const layoutRows = buildFixedInvoiceLayoutRows_(blocks);
  const targetTotalHeight = targetHeight + bufferRowExtensionPx;

  if (layoutRows.length > contentRowCount) {
    return {
      success: false,
      message: "Trop d’informations pour une seule facture. Veuillez réduire le nombre de blocs, d’activités ou de descriptions."
    };
  }

  const requestedContentHeight = layoutRows.reduce((sum, layoutRow) => {
    return sum + layoutRow.height;
  }, 0);
  const preliminaryBufferHeight = targetTotalHeight - requestedContentHeight;
  if (preliminaryBufferHeight < minimumRowHeight) {
    return {
      success: false,
      message: "Trop d’informations pour une seule facture. Veuillez réduire le nombre de blocs, d’activités ou de descriptions."
    };
  }

  const workRange = sheet.getRange(startRow, 1, contentRowCount, 16);
  workRange.breakApart();
  workRange.clearContent();
  workRange.setWrap(false);
  sheet.setRowHeights(startRow, contentRowCount, minimumRowHeight);

  layoutRows.forEach((layoutRow, index) => {
    const rowNumber = startRow + index;
    sheet.setRowHeight(rowNumber, layoutRow.height);

    if (layoutRow.type === "title") {
      if (layoutRow.isFirstTitleLine) {
        sheet.getRange(rowNumber, 1)
          .setValue(layoutRow.block.blockNumber)
          .setFontFamily("Roboto")
          .setFontSize(10)
          .setFontColor("#000000")
          .setHorizontalAlignment("left")
          .setVerticalAlignment("middle");
        sheet.getRange(rowNumber, 13, 1, 2).merge()
          .setValue(`${layoutRow.block.totalTime.toFixed(2)} h`)
          .setFontFamily("Roboto")
          .setFontSize(12)
          .setFontColor("#000000")
          .setHorizontalAlignment("right")
          .setVerticalAlignment("middle");
        sheet.getRange(rowNumber, 15, 1, 2).merge()
          .setValue(Number(layoutRow.block.totalPrice.toFixed(2)))
          .setNumberFormat("0.00 $")
          .setFontFamily("Roboto")
          .setFontSize(12)
          .setFontColor("#000000")
          .setFontWeight("bold")
          .setHorizontalAlignment("right")
          .setVerticalAlignment("middle");
      }
      sheet.getRange(rowNumber, 2, 1, 11).merge()
        .setValue(layoutRow.text)
        .setFontFamily("Roboto")
        .setFontSize(12)
        .setFontColor("#000000")
        .setFontWeight("bold")
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle");
      return;
    }

    if (layoutRow.type === "description") {
      sheet.getRange(rowNumber, 2, 1, 11).merge()
        .setValue(layoutRow.text)
        .setFontFamily("Roboto")
        .setFontSize(10)
        .setFontColor("#000000")
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle");
      return;
    }

    if (layoutRow.type === "activity") {
      sheet.getRange(rowNumber, 2, 1, 3).merge()
        .setValue(layoutRow.activity.name)
        .setFontFamily("Roboto")
        .setFontSize(10)
        .setFontColor("#999999")
        .setHorizontalAlignment("left")
        .setVerticalAlignment("middle");
      sheet.getRange(rowNumber, 5)
        .setValue(`${layoutRow.activity.time.toFixed(2)} h`)
        .setFontFamily("Roboto")
        .setFontSize(10)
        .setFontColor("#999999")
        .setHorizontalAlignment("right")
        .setVerticalAlignment("middle");
    }
  });

  layoutRows.forEach((layoutRow, index) => {
    const rowNumber = startRow + index;
    sheet.setRowHeight(rowNumber, layoutRow.height);
  });

  SpreadsheetApp.flush();
  const actualContentHeight = Array.from({ length: layoutRows.length }, (_, index) => {
    return sheet.getRowHeight(startRow + index);
  }).reduce((sum, height) => sum + height, 0);
  const firstTrailingEmptyRow = startRow + layoutRows.length;
  const rowsToDelete = contentRowCount - layoutRows.length;
  const bufferHeight = targetTotalHeight - actualContentHeight;
  if (bufferHeight < minimumRowHeight) {
    return {
      success: false,
      message: "Trop d’informations pour une seule facture. Veuillez réduire le nombre de blocs, d’activités ou de descriptions."
    };
  }

  if (rowsToDelete > 0) {
    sheet.deleteRows(firstTrailingEmptyRow, rowsToDelete);
    SpreadsheetApp.flush();
  }

  const finalBufferRow = firstTrailingEmptyRow;
  sheet.setRowHeight(finalBufferRow, bufferHeight);

  return { success: true };
}

function getInvoiceRecipientOptions(invoiceNumber) {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const clientsSheet = facturerSpreadsheet.getSheetByName("CLIENTS");

  if (!facturerTrackingSheet || !clientsSheet) {
    return { success: false, message: "Erreur lors de la préparation de l’envoi." };
  }

  const normalizedInvoiceNumber = String(invoiceNumber || "").trim();
  if (!normalizedInvoiceNumber) {
    return { success: false, message: "Erreur lors de la préparation de l’envoi." };
  }

  const facturerTrackingLastRow = facturerTrackingSheet.getLastRow();
  const trackingValues = facturerTrackingLastRow >= 6
    ? facturerTrackingSheet.getRange(`A6:I${facturerTrackingLastRow}`).getValues()
    : [];
  const trackingIndex = trackingValues.findIndex(row => String(row[1] || "").trim() === normalizedInvoiceNumber);
  if (trackingIndex === -1) {
    return { success: false, message: "Erreur lors de la préparation de l’envoi." };
  }

  const clientName = String(trackingValues[trackingIndex][3] || "").trim();
  if (!clientName) {
    return { success: false, message: "Erreur lors de la préparation de l’envoi." };
  }

  const clientsLastRow = clientsSheet.getLastRow();
  const clientRows = clientsLastRow >= 2
    ? clientsSheet.getRange(`A2:G${clientsLastRow}`).getValues()
    : [];
  const clientEntry = clientRows.find(row => normalizeString_(row[0]) === normalizeString_(clientName));
  if (!clientEntry) {
    logSendInvoiceEmailError_("client_lookup", `client not found for ${clientName}`);
    return { success: false, message: "Erreur lors de la préparation de l’envoi." };
  }

  const recipientOptions = [
    { name: String(clientEntry[1] || "").trim(), email: String(clientEntry[2] || "").trim(), source: "primary" },
    { name: String(clientEntry[3] || "").trim(), email: String(clientEntry[4] || "").trim(), source: "secondary" },
    { name: String(clientEntry[5] || "").trim(), email: String(clientEntry[6] || "").trim(), source: "billing" }
  ]
    .filter(option => option.email !== "")
    .map(option => ({
      name: option.name,
      email: option.email,
      source: option.source,
      label: `${option.name || clientName} - ${option.email}`
    }));

  if (recipientOptions.length === 0) {
    return { success: false, message: "Aucun contact avec courriel n'est configuré pour ce client." };
  }

  const defaultRecipient = recipientOptions.find(option => option.source === "billing")
    || recipientOptions.find(option => option.source === "primary");
  if (!defaultRecipient) {
    return { success: false, message: "Aucun contact avec courriel n'est configuré pour ce client." };
  }

  return {
    success: true,
    recipientOptions,
    defaultRecipientEmail: defaultRecipient.email
  };
}

function sendInvoiceEmail_(invoiceNumber, pdfUrl, recipient, ccRecipient) {
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
    logSendInvoiceEmailError_("file_lookup", e.message);
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

  const clientName = String(trackingValues[trackingIndex][3] || "").trim();
  if (!clientName) {
    logSendInvoiceEmailError_("tracking_client_lookup", "missing client name");
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const clientsLastRow = clientsSheet.getLastRow();
  const clientRows = clientsLastRow >= 2
    ? clientsSheet.getRange(`A2:G${clientsLastRow}`).getValues()
    : [];
  const clientEntry = clientRows.find(row => normalizeString_(row[0]) === normalizeString_(clientName));
  if (!clientEntry) {
    logSendInvoiceEmailError_("client_lookup", `client not found for ${clientName}`);
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const recipientName = String(recipient && recipient.name || "").trim();
  const recipientEmail = String(recipient && recipient.email || "").trim();
  const ccEmail = String(ccRecipient && ccRecipient.email || "").trim();
  if (!recipientEmail) {
    logSendInvoiceEmailError_("recipient_lookup", `missing recipient email for ${clientName}`);
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
    logSendInvoiceEmailError_("time_rows_lookup", `no rows found for invoice ${normalizedInvoiceNumber}`);
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
    logSendInvoiceEmailError_("scope_build", `no campaign/project pairs for invoice ${normalizedInvoiceNumber}`);
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  const greetingName = recipientName || clientName;
  const emailSubject = `Facture ${normalizedInvoiceNumber} de Mathieu Renaud`;
  const emailBody = [
    `Bonjour ${greetingName},`,
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
    if (ccEmail && ccEmail !== recipientEmail) {
      mailOptions.cc = ccEmail;
    }
    MailApp.sendEmail(recipientEmail, emailSubject, emailBody, mailOptions);
  } catch (e) {
    logSendInvoiceEmailError_("email_send", e.message);
    return { success: false, message: "Erreur lors de l’envoi du courriel." };
  }

  return { success: true };
}

function paye() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");

  if (!facturerTrackingSheet) {
    openStandaloneMessageView_("Erreur : La feuille FACTURATION n'existe pas.");
    return;
  }

  const lastRow = facturerTrackingSheet.getLastRow();
  const trackingValues = lastRow >= 6
    ? facturerTrackingSheet.getRange(`B6:H${lastRow}`).getValues()
    : [];
  const unpaidInvoices = trackingValues
    .map((row, index) => ({
      invoiceNumber: String(row[0] || "").trim(),
      paymentDate: row[6],
      rowNumber: index + 6
    }))
    .filter(entry => entry.invoiceNumber !== "" && (entry.paymentDate === "" || entry.paymentDate === null));

  if (unpaidInvoices.length === 0) {
    openStandaloneMessageView_("Aucune facture non payée existe.", "Attention");
    return;
  }

  const html = HtmlService.createTemplateFromFile("popupPayment");
  html.invoices = unpaidInvoices;

  const htmlOutput = html.evaluate().setWidth(400).setHeight(360);
  htmlOutput.addMetaTag('viewport', 'width=device-width, initial-scale=1');
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "Enregistrer un paiement");
}

function registerInvoicePayments(selectedInvoiceNumbers) {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTrackingSheet = facturerSpreadsheet.getSheetByName("FACTURATION");
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");

  if (!facturerTrackingSheet || !facturerTimeSheet) {
    throw new Error("Erreur : La feuille 'FACTURATION' ou 'FEUILLE DE TEMPS' est manquante.");
  }

  const normalizedInvoiceNumbers = [...new Set(
    (Array.isArray(selectedInvoiceNumbers) ? selectedInvoiceNumbers : [])
      .map(value => String(value || "").trim())
      .filter(String)
  )];
  const normalizedSelectedInvoiceKeys = new Set(
    normalizedInvoiceNumbers.map(normalizeInvoiceNumberForComparison_).filter(String)
  );

  if (normalizedInvoiceNumbers.length === 0) {
    throw new Error("Aucune facture sélectionnée.");
  }

  const lastTrackingRow = facturerTrackingSheet.getLastRow();
  const trackingValues = lastTrackingRow >= 6
    ? facturerTrackingSheet.getRange(`B6:H${lastTrackingRow}`).getValues()
    : [];
  const selectedTrackingRows = trackingValues
    .map((row, index) => ({
      invoiceNumber: String(row[0] || "").trim(),
      paymentDate: row[6],
      rowNumber: index + 6
    }))
    .filter(entry => normalizedSelectedInvoiceKeys.has(normalizeInvoiceNumberForComparison_(entry.invoiceNumber)));

  if (selectedTrackingRows.some(entry => entry.paymentDate !== "" && entry.paymentDate !== null)) {
    throw new Error("Cette facture est déjà marquée comme payée.");
  }

  const today = new Date();
  selectedTrackingRows.forEach(({ rowNumber }) => {
    facturerTrackingSheet.getRange(`H${rowNumber}`).setValue(today).setNumberFormat("d mmmm yyyy");
  });

  const lastTimeRow = facturerTimeSheet.getLastRow();
  if (lastTimeRow >= 7) {
    const timeRows = facturerTimeSheet.getRange(`P7:S${lastTimeRow}`).getValues();
    timeRows.forEach((row, index) => {
      const invoiceNumber = String(row[0] || "").trim();
      const normalizedTimeInvoice = normalizeInvoiceNumberForComparison_(invoiceNumber);
      if (normalizedSelectedInvoiceKeys.has(normalizedTimeInvoice)) {
        const rowNumber = index + 7;
        facturerTimeSheet.getRange(`R${rowNumber}`).setValue(true);
        facturerTimeSheet.getRange(`S${rowNumber}`).setValue(today).setNumberFormat("d mmmm yyyy");
        facturerTimeSheet.getRange(`A${rowNumber}:U${rowNumber}`).setFontColor("#999999");
      }
    });
  }
}

function submitFacturerForm(contact, activityType, invoiceNumber, overwriteExistingFile, invoicePreview) {
  const validationResult = validateInvoiceGeneration_();
  if (!validationResult.success) {
    return validationResult;
  }
  const {
    facturerSpreadsheet,
    facturerTimeSheet,
    facturerModelSheet,
    facturerTrackingSheet,
    facturerDriveFolder,
    facturerCheckedRows
  } = validationResult;

  if (contact === "Sélectionnez un contact" || activityType === "Sélectionnez une activité") {
    return { success: false, message: "Veuillez sélectionner un contact et une activité générale." };
  }

  const invoiceNumberResolution = resolveInvoiceNumberForFacturation_(facturerTrackingSheet, invoiceNumber);
  if (!invoiceNumberResolution.success) {
    return invoiceNumberResolution;
  }
  const facturerFullInvoiceNumber = invoiceNumberResolution.fullInvoiceNumber;

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

  const facturerRowsToInvoice = filterFacturerCheckedRowsByInvoicePreview_(facturerCheckedRows, invoicePreview);
  if (!facturerRowsToInvoice.length) {
    return { success: false, message: "Aucune activité n’est sélectionnée" };
  }

  const wasModelSheetHidden = facturerModelSheet.isSheetHidden();
  if (wasModelSheetHidden) {
    facturerModelSheet.showSheet();
  }
  let facturerTempSheet;
  try {
    facturerTempSheet = facturerModelSheet.copyTo(facturerSpreadsheet).setName(facturerFullInvoiceNumber);
  } finally {
    if (wasModelSheetHidden) {
      facturerModelSheet.hideSheet();
    }
  }
  Logger.log(`Temp invoice sheet created: name="${facturerTempSheet.getName()}", sheetId=${facturerTempSheet.getSheetId()}, index=${facturerTempSheet.getIndex()}`);

  const facturerInvoiceBlocks = applyValidatedInvoicePreviewToBlocks_(buildFixedInvoiceBlocks_(facturerRowsToInvoice), invoicePreview);
  const validatedServiceTitle = cleanGeneratedInvoiceText_(invoicePreview && invoicePreview.serviceTitle);
  const facturerServiceTitle = validatedServiceTitle || activityType;
  const facturerTotalAmount = facturerInvoiceBlocks.reduce((sum, block) => sum + block.totalPrice, 0).toFixed(2);
  const facturerClientName = String(facturerRowsToInvoice[0].row[1] || "");
  const facturerToday = new Date();
  const facturerCampaignSummary = [];
  const facturerSeenCampaignKeys = new Set();
  facturerRowsToInvoice.forEach(row => {
    const facturerCampaignName = String(row.row[2] || "").trim();
    const facturerCampaignKey = normalizeString_(facturerCampaignName);
    if (!facturerCampaignName || facturerSeenCampaignKeys.has(facturerCampaignKey)) {
      return;
    }
    facturerSeenCampaignKeys.add(facturerCampaignKey);
    facturerCampaignSummary.push(facturerCampaignName);
  });
  const facturerScopeSummary = facturerCampaignSummary.join(", ");

  // Forcer texte pour préserver les zéros à gauche ("001") dans le PDF.
  facturerTempSheet.getRange("L1").setNumberFormat("@").setValue(facturerFullInvoiceNumber);
  facturerTempSheet.getRange("C7").setValue(facturerToday).setNumberFormat("d mmmm yyyy");
  facturerTempSheet.getRange("C10").setValue(contact);
  facturerTempSheet.getRange("C12").setValue(formatFrenchInvoiceList_(facturerCampaignSummary));
  facturerTempSheet.getRange("C14").setValue(facturerServiceTitle);
  facturerTempSheet.getRange("C17").setValue(Number(facturerTotalAmount));
  facturerTempSheet.getRange("N51").setValue(Number(facturerTotalAmount));
  const fixedBlockWriteResult = writeFixedInvoiceBlocks_(facturerTempSheet, facturerInvoiceBlocks);
  if (!fixedBlockWriteResult.success) {
    facturerSpreadsheet.deleteSheet(facturerTempSheet);
    return fixedBlockWriteResult;
  }

  SpreadsheetApp.flush();
  let facturerPdfFile = null;
  try {
    const facturerSheetNamesBeforeExport = facturerSpreadsheet.getSheets().map(sheet => sheet.getName());
    const facturerTempSheetStillExists = facturerSpreadsheet.getSheets().some(sheet => sheet.getSheetId() === facturerTempSheet.getSheetId());
    Logger.log(`Temp invoice sheet expected before export: name="${facturerTempSheet.getName()}", sheetId=${facturerTempSheet.getSheetId()}`);
    Logger.log(`Spreadsheet sheets before export: ${JSON.stringify(facturerSheetNamesBeforeExport)}`);
    Logger.log(`Temp invoice sheet exists before export: ${facturerTempSheetStillExists}`);
    const facturerPdfBlob = exportInvoiceSheetPdfBlob_(facturerSpreadsheet.getId(), facturerTempSheet.getSheetId(), facturerFileName);
    if (facturerExistingFile) {
      facturerPdfFile = facturerDriveFolder.createFile(facturerPdfBlob);
      facturerExistingFile.setTrashed(true);
    } else {
      facturerPdfFile = facturerDriveFolder.createFile(facturerPdfBlob);
    }
    const facturerPdfUrl = facturerPdfFile.getUrl();

    facturerRowsToInvoice.forEach(row => {
      const facturerRowIndex = row.index;
      facturerTimeSheet.getRange(`A${facturerRowIndex}`).setValue(false);
      facturerTimeSheet.getRange(`O${facturerRowIndex}`).setValue(true);
      // Forcer texte pour préserver les zéros à gauche dans FEUILLE DE TEMPS.
      facturerTimeSheet.getRange(`P${facturerRowIndex}`).setNumberFormat("@").setValue(facturerFullInvoiceNumber);
      facturerTimeSheet.getRange(`Q${facturerRowIndex}`).setValue(facturerToday).setNumberFormat("d mmmm yyyy");
    });

    const facturerTrackingRow = facturerTrackingSheet.getLastRow() + 1 >= 6 ? facturerTrackingSheet.getLastRow() + 1 : 6;
    // IMPORTANT: fixer le format texte avant d'écrire pour éviter la coercition "001" -> 1.
    facturerTrackingSheet.getRange(`B${facturerTrackingRow}`).setNumberFormat("@");
    facturerTrackingSheet.getRange(`A${facturerTrackingRow}:I${facturerTrackingRow}`).setValues([[
      false,
      facturerFullInvoiceNumber,
      facturerToday,
      facturerClientName,
      facturerScopeSummary,
      Number(facturerTotalAmount),
      `=HYPERLINK("${facturerPdfUrl}"; "Voir PDF")`,
      "",
      ""
    ]]);
    // Case à cocher réelle en colonne A (pas seulement la valeur FALSE).
    facturerTrackingSheet.getRange(`A${facturerTrackingRow}`).insertCheckboxes().setValue(false);
    facturerTrackingSheet.getRange(`C${facturerTrackingRow}`).setNumberFormat("d mmmm yyyy");

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

function prepareInvoicePreview(invoiceNumber) {
  const validationResult = validateInvoiceGeneration_();
  if (!validationResult.success) {
    return validationResult;
  }
  const {
    facturerTimeSheet,
    facturerGestionSheet,
    facturerTrackingSheet
  } = validationResult;

  const invoiceNumberResolution = resolveInvoiceNumberForFacturation_(facturerTrackingSheet, invoiceNumber);
  if (!invoiceNumberResolution.success) {
    return invoiceNumberResolution;
  }
  const facturerFullInvoiceNumber = invoiceNumberResolution.fullInvoiceNumber;

  const facturerLastRow = facturerTimeSheet.getLastRow();
  const facturerTimeData = facturerLastRow >= 7 ? facturerTimeSheet.getRange(`A7:U${facturerLastRow}`).getValues() : [];
  const facturerCheckedRows = facturerTimeData.map((row, index) => ({ row: row, index: index + 7 }))
    .filter(row => row.row[0] === true);
  const toPreviewTime = (value) => value instanceof Date
    ? value.getHours() + value.getMinutes() / 60
    : Number(value);
  const cleanPreviewText = (value) => String(value || "").replace(/\s+/g, " ").trim();
  const joinPreviewList = (values) => {
    if (values.length <= 1) return values.join("");
    if (values.length === 2) return `${values[0]} et ${values[1]}`;
    return `${values.slice(0, -1).join(", ")} et ${values[values.length - 1]}`;
  };
  const buildPreviewServiceTitle = (blocks) => {
    const activityTotals = [];
    blocks.forEach(block => {
      block.activities.forEach(activity => {
        const activityName = cleanPreviewText(activity.name);
        if (!activityName) return;
        const activityKey = normalizeString_(activityName);
        const existingActivity = activityTotals.find(item => item.key === activityKey);
        if (existingActivity) {
          existingActivity.time += activity.time;
        } else {
          activityTotals.push({ key: activityKey, name: activityName, time: activity.time });
        }
      });
    });
    const orderedActivities = activityTotals
      .sort((a, b) => b.time - a.time || a.name.localeCompare(b.name))
      .map(activity => activity.name);
    if (orderedActivities.length === 1) {
      return `Service : ${orderedActivities[0]}`;
    }
    if (orderedActivities.length > 1) {
      const titleActivities = orderedActivities.length > 2
        ? `${joinPreviewList(orderedActivities.slice(0, 2))} et autres activités`
        : joinPreviewList(orderedActivities);
      return `Services : ${titleActivities}`;
    }

    const projectNames = [...new Set(blocks.map(block => cleanPreviewText(block.project)).filter(String))];
    if (projectNames.length) {
      return `Services pour ${joinPreviewList(projectNames)}`;
    }
    return "Services facturés";
  };
  const buildPreviewDescription = (block) => {
    const cleanedNotes = [];
    block.notes.forEach(note => {
      const cleanNote = cleanPreviewText(note).replace(/[.;\s]+$/g, "");
      if (!cleanNote) return;
      const noteKey = normalizeString_(cleanNote);
      if (!cleanedNotes.some(existingNote => normalizeString_(existingNote) === noteKey)) {
        cleanedNotes.push(cleanNote);
      }
    });
    if (cleanedNotes.length) {
      return `Travaux réalisés : ${cleanedNotes.join("; ")}.`;
    }

    const activitiesSummary = block.activities
      .map(activity => `${activity.name} (${activity.time.toFixed(2)} h)`)
      .join(", ");
    if (activitiesSummary) {
      return `Activités réalisées : ${activitiesSummary}.`;
    }

    const blockContext = cleanPreviewText(block.project) || cleanPreviewText(block.campaign);
    return blockContext ? `Travaux réalisés pour ${blockContext}.` : "Travaux réalisés.";
  };
  const previewBlocks = [];
  facturerCheckedRows.forEach(({ row, index }) => {
    const campaign = String(row[2] || "").trim();
    const project = String(row[3] || "").trim();
    const blockKey = `${campaign}|||${project}`;
    if (!previewBlocks.some(block => block.key === blockKey)) {
      previewBlocks.push({
        key: blockKey,
        campaign,
        project,
        notes: [],
        activities: [],
        rowIndexes: [],
        totalTime: 0,
        totalPrice: 0
      });
    }
    const previewBlock = previewBlocks.find(block => block.key === blockKey);
    previewBlock.rowIndexes.push(index);
    const activityName = String(row[4] || "").trim();
    const time = toPreviewTime(row[8]);
    const price = Number(row[10]);
    const existingActivity = previewBlock.activities.find(activity => activity.name === activityName);
    if (existingActivity) {
      existingActivity.time += time;
    } else {
      previewBlock.activities.push({ name: activityName, time });
    }
    previewBlock.totalTime += time;
    previewBlock.totalPrice += price;

    const note = String(row[20] || "").trim();
    if (note && !previewBlock.notes.includes(note)) {
      previewBlock.notes.push(note);
    }
  });

  const facturerClientName = String(facturerCheckedRows[0].row[1] || "");
  const facturerToday = new Date();
  const projects = [...new Set(facturerCheckedRows.map(({ row }) => String(row[3] || "").trim()).filter(String))];
  const campaigns = [...new Set(previewBlocks.map(block => String(block.campaign || "").trim()).filter(String))];
  const totalAmount = previewBlocks.reduce((sum, block) => sum + block.totalPrice, 0);
  const shouldUseOpenAI = facturerGestionSheet.getRange("G2").getValue() === true;
  let serviceTitle = buildPreviewServiceTitle(previewBlocks);
  const blocks = previewBlocks.map((block, index) => {
    return {
      blockNumber: index + 1,
      campaign: block.campaign,
      project: block.project,
      description: buildPreviewDescription(block),
      activities: block.activities.map(activity => ({
        name: activity.name,
        time: Number(activity.time.toFixed(2))
      })),
      rowIndexes: block.rowIndexes.slice(),
      totalTime: Number(block.totalTime.toFixed(2)),
      totalPrice: Number(block.totalPrice.toFixed(2))
    };
  });
  const generatedInvoiceText = shouldUseOpenAI
    ? generateInvoiceTextWithOpenAI({
      client: facturerClientName,
      campaigns,
      projects,
      blocks: previewBlocks.map((block, index) => ({
        blockNumber: index + 1,
        campaign: block.campaign,
        project: block.project,
        activities: block.activities.map(activity => ({
          name: activity.name,
          time: Number(activity.time.toFixed(2))
        })),
        notes: block.notes
      }))
    })
    : null;
  if (generatedInvoiceText) {
    const generatedServiceTitle = cleanGeneratedInvoiceText_(generatedInvoiceText.serviceTitle);
    if (generatedServiceTitle) {
      serviceTitle = generatedServiceTitle;
    }

    blocks.forEach((block, index) => {
      const generatedBlocks = Array.isArray(generatedInvoiceText.blocks) ? generatedInvoiceText.blocks : [];
      const generatedBlock = generatedBlocks.find(candidate => Number(candidate.blockNumber) === block.blockNumber)
        || generatedBlocks.find(candidate => normalizeString_(candidate.campaign) === normalizeString_(block.campaign)
          && normalizeString_(candidate.project) === normalizeString_(block.project))
        || generatedBlocks[index];
      const generatedDescription = generatedBlock ? cleanGeneratedInvoiceText_(generatedBlock.description) : "";
      if (generatedDescription) {
        block.description = generatedDescription;
      }
    });
  }

  return {
    success: true,
    preview: {
      invoiceNumber: facturerFullInvoiceNumber,
      invoiceDate: Utilities.formatDate(facturerToday, Session.getScriptTimeZone(), "yyyy-MM-dd"),
      client: facturerClientName,
      projects,
      serviceTitle,
      totalAmount: Number(totalAmount.toFixed(2)),
      blocks
    }
  };
}

// NOUVELLE ENTRÉ DE TEMPS

function newTimeEntry() {
  const facturerSpreadsheet = SpreadsheetApp.getActiveSpreadsheet();
  const facturerTimeSheet = facturerSpreadsheet.getSheetByName("FEUILLE DE TEMPS");
  const facturerGestionSheet = facturerSpreadsheet.getSheetByName("GESTION");
  const facturerClientsSheet = facturerSpreadsheet.getSheetByName("CLIENTS");

  if (!facturerTimeSheet || !facturerGestionSheet || !facturerClientsSheet) {
    openStandaloneMessageView_("Erreur : La feuille 'FEUILLE DE TEMPS', 'GESTION' ou 'CLIENTS' est manquante.");
    return;
  }

  const lastRowGestion = facturerGestionSheet.getLastRow();
  const lastRow = Math.max(7, facturerTimeSheet.getLastRow());
  const clients = [...new Set(
    facturerTimeSheet.getRange("B7:B" + lastRow).getValues().flat().map(value => String(value || "").trim()).filter(String)
  )];
  const activities = facturerGestionSheet.getRange("A2:A" + Math.max(2, lastRowGestion)).getValues().flat().filter(String);
  let rates = ['0']; // Valeur par défaut
  try {
    rates = facturerGestionSheet.getRange("C2:C" + Math.max(2, lastRowGestion)).getValues().flat().filter(String);
    if (rates.length === 0) rates = ['0'];
  } catch (e) {
    openStandaloneMessageView_("Erreur : Impossible de lire les taux dans GESTION!C2:C. Valeur par défaut utilisée.", "Information");
  }

  const clientScopedRows = facturerTimeSheet.getRange("B7:D" + lastRow).getValues();
  const campaignOptionsByClient = {};
  const projectOptionsByClientAndCampaign = {};
  clientScopedRows.forEach(row => {
    const client = String(row[0] || "").trim();
    const campaign = String(row[1] || "").trim();
    const project = String(row[2] || "").trim();
    if (!client) return;
    if (!campaignOptionsByClient[client]) campaignOptionsByClient[client] = [];
    if (campaign && !campaignOptionsByClient[client].includes(campaign)) {
      campaignOptionsByClient[client].push(campaign);
    }
    const projectKey = `${client}|||${campaign}`;
    if (!projectOptionsByClientAndCampaign[projectKey]) projectOptionsByClientAndCampaign[projectKey] = [];
    if (campaign && project && !projectOptionsByClientAndCampaign[projectKey].includes(project)) {
      projectOptionsByClientAndCampaign[projectKey].push(project);
    }
  });

  const taskStateData = facturerTimeSheet.getRange("A7:H" + lastRow).getValues();
  const checkedRows = taskStateData
    .map((row, index) => ({ checked: row[0], startTime: row[6], endTime: row[7], index: index + 7 }))
    .filter(row => row.checked === true);
  const checkedIndexes = checkedRows.map(row => row.index);

  if (checkedIndexes.length > 1) {
    openStandaloneMessageView_(`${checkedIndexes.length} lignes sont présentement cochées. Veuillez recommencer.`, "Attention");
    return;
  }

  const activeTaskRows = taskStateData
    .map((row, index) => ({ startTime: row[6], endTime: row[7], index: index + 7 }))
    .filter(row => row.startTime !== "" && row.startTime !== null && (row.endTime === "" || row.endTime === null));
  if (activeTaskRows.length > 0) {
    openStandaloneMessageView_("Une tâche est présentement en cours.", "Attention");
    return;
  }

  const checkedRowIndex = checkedIndexes.length === 1 ? checkedIndexes[0] : -1;

  const html = HtmlService.createTemplateFromFile("popupTemps");
  html.clients = clients || [];
  html.campaignOptionsByClient = campaignOptionsByClient || {};
  html.projectOptionsByClientAndCampaign = projectOptionsByClientAndCampaign || {};
  html.activities = activities || [];
  html.rates = rates || ['0'];
  html.checkedRowIndex = checkedRowIndex;

  if (checkedRowIndex !== -1) {
    let sourceData;
    let sourceRate = "";
    try {
      sourceData = facturerTimeSheet.getRange(`B${checkedRowIndex}:E${checkedRowIndex}`).getValues()[0];
      sourceRate = facturerTimeSheet.getRange(`T${checkedRowIndex}`).getValue();
    } catch (e) {
      sourceData = ["", "", "", ""];
      sourceRate = "";
    }
    html.clientSelected = sourceData[0] || "";
    html.campaign = sourceData[1] || "";
    html.project = sourceData[2] || "";
    html.activitySelected = sourceData[3] || "";
    html.rateSelected = sourceRate === null || typeof sourceRate === "undefined" ? "" : sourceRate;
    html.newRow = checkedRowIndex + 1;
  } else {
    html.clientSelected = "";
    html.campaign = "";
    html.project = "";
    html.activitySelected = "";
    html.rateSelected = "";
    html.newRow = 7;
  }

  const htmlOutput = html.evaluate().setWidth(690).setHeight(300);
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, "Nouvelle entrée de temps");
}

function submitTimeEntryForm(client, campaign, project, activity, newRow, checkedRowIndex, newClient, newActivity, rate, note, options) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetTime = ss.getSheetByName("FEUILLE DE TEMPS");
  const sheetGestion = ss.getSheetByName("GESTION");
  const sheetClients = ss.getSheetByName("CLIENTS");
  const normalizedOptions = options || {};
  const newClientDetails = normalizedOptions.newClientDetails || null;
  const isCustomCampaign = normalizedOptions.isCustomCampaign === true;
  const isCustomProject = normalizedOptions.isCustomProject === true;

  if (!sheetTime || !sheetGestion || !sheetClients) {
    throw new Error("Erreur : La feuille 'FEUILLE DE TEMPS', 'GESTION' ou 'CLIENTS' est manquante.");
  }

  if (!newClientDetails) {
    const lastRowTime = sheetTime.getLastRow();
    const clientScopedRows = lastRowTime >= 7
      ? sheetTime.getRange("B7:D" + lastRowTime).getValues()
      : [];
    const allowedCampaigns = [...new Set(
      clientScopedRows
        .filter(row => String(row[0] || "").trim() === String(client || "").trim())
        .map(row => String(row[1] || "").trim())
        .filter(String)
    )];
    const allowedProjects = [...new Set(
      clientScopedRows
        .filter(row => String(row[0] || "").trim() === String(client || "").trim())
        .map(row => String(row[2] || "").trim())
        .filter(String)
    )];

    if ((!isCustomCampaign && !allowedCampaigns.includes(String(campaign || "").trim())) || (!isCustomProject && !allowedProjects.includes(String(project || "").trim()))) {
      throw new Error("Attention. Données invalides.");
    }
  }

  if (newClientDetails && newClientDetails.companyName) {
    const companyName = String(newClientDetails.companyName || "").trim();
    const primaryContactName = String(newClientDetails.primaryContactName || "").trim();
    const primaryContactEmail = String(newClientDetails.primaryContactEmail || "").trim();
    const secondaryContactName = String(newClientDetails.secondaryContactName || "").trim();
    const secondaryContactEmail = String(newClientDetails.secondaryContactEmail || "").trim();
    const billingContactName = String(newClientDetails.billingContactName || "").trim();
    const billingEmail = String(newClientDetails.billingEmail || "").trim();

    const lastRowClients = sheetClients.getLastRow();
    const clientRows = lastRowClients >= 2
      ? sheetClients.getRange(`A2:G${lastRowClients}`).getValues()
      : [];
    const existingClientIndex = clientRows.findIndex(row => normalizeString_(row[0]) === normalizeString_(companyName));
    if (existingClientIndex === -1) {
      const insertRow = Math.max(2, lastRowClients + 1);
      sheetClients.getRange(`A${insertRow}:G${insertRow}`).setValues([[
        companyName,
        primaryContactName,
        primaryContactEmail,
        secondaryContactName,
        secondaryContactEmail,
        billingContactName,
        billingEmail
      ]]);
    }
    client = companyName;
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
  const currentTimeValue = (now.getHours() * 60 + now.getMinutes()) / (24 * 60);
  const buildTimeFormula = (rowNumber) => `=IF(H${rowNumber}<>""; ROUND((IF(H${rowNumber}<G${rowNumber}; H${rowNumber}+1; H${rowNumber})-G${rowNumber})*96)/4; "")`;
  const buildHoursCumulativeFormula = (rowNumber) => `=IF(H${rowNumber}<>""; SUM($I$7:I${rowNumber}); "")`;
  const buildAmountFormula = (rowNumber) => `=IF(H${rowNumber}<>""; I${rowNumber}*T${rowNumber}; "")`;
  const buildAmountCumulativeFormula = (rowNumber) => `=IF(H${rowNumber}<>""; SUM($K$7:K${rowNumber}); "")`;
  const applyTimeColumnsDisplayFormat = () => {
    const lastTimeRow = Math.max(7, sheetTime.getLastRow());
    sheetTime.getRange(`G7:H${lastTimeRow}`).setNumberFormat("HH:mm");
  };
  const applyTimeEntryRowFormatting = (rowNumber) => {
    sheetTime.getRange(`A${rowNumber}:U${rowNumber}`)
      .setFontFamily("Roboto")
      .setFontSize(10)
      .setFontColor("#000000")
      .setBackground("#ffffff")
      .setVerticalAlignment("middle")
      .setBorder(false, false, false, false, false, false);

    sheetTime.getRange(`A${rowNumber}`).setHorizontalAlignment("center");
    sheetTime.getRange(`B${rowNumber}:D${rowNumber}`).setHorizontalAlignment("left");
    sheetTime.getRange(`E${rowNumber}`).setHorizontalAlignment("left").setFontWeight("bold");
    sheetTime.getRange(`F${rowNumber}`).setHorizontalAlignment("right").setNumberFormat("d mmmm yyyy");
    sheetTime.getRange(`G${rowNumber}`).setHorizontalAlignment("center").setNumberFormat("HH:mm");
    sheetTime.getRange(`H${rowNumber}`).setHorizontalAlignment("left").setNumberFormat("HH:mm");
    sheetTime.getRange(`I${rowNumber}:J${rowNumber}`).setHorizontalAlignment("center").setNumberFormat("0.00");
    sheetTime.getRange(`K${rowNumber}`).setHorizontalAlignment("right").setNumberFormat("0.00 $").setFontWeight("bold");
    sheetTime.getRange(`L${rowNumber}`).setHorizontalAlignment("right").setNumberFormat("0.00 $");
    sheetTime.getRange(`N${rowNumber}:P${rowNumber}`).setHorizontalAlignment("center");
    sheetTime.getRange(`Q${rowNumber}`).setHorizontalAlignment("center").setNumberFormat("dd MM yyyy");
    sheetTime.getRange(`R${rowNumber}`).setHorizontalAlignment("center");
    sheetTime.getRange(`S${rowNumber}`).setHorizontalAlignment("center").setNumberFormat("dd MM yyyy");
    sheetTime.getRange(`T${rowNumber}`).setHorizontalAlignment("center").setNumberFormat("0").setFontColor("#9a9a9a");
    sheetTime.getRange(`U${rowNumber}`).setHorizontalAlignment("left");

    ["E", "M", "N", "Q", "S"].forEach(columnLetter => {
      sheetTime.getRange(`${columnLetter}${rowNumber}`).setBorder(
        false,
        false,
        false,
        true,
        false,
        false,
        "#acacac",
        SpreadsheetApp.BorderStyle.DOTTED
      );
    });
  };

  if (checkedRowIndex !== -1) {
    // Insertion après ligne cochée
    sheetTime.insertRowAfter(checkedRowIndex);
    const targetRow = checkedRowIndex + 1;
    sheetTime.getRange(`A${targetRow}:U${targetRow}`).clearContent();
    sheetTime.getRange(`A${targetRow}`).insertCheckboxes();
    sheetTime.getRange(`N${targetRow}`).insertCheckboxes();
    sheetTime.getRange(`O${targetRow}`).insertCheckboxes();
    sheetTime.getRange(`R${targetRow}`).insertCheckboxes();
    sheetTime.getRange(`A${targetRow}`).setValue(true);
    sheetTime.getRange(`B${targetRow}`).setValue(client);
    sheetTime.getRange(`C${targetRow}`).setValue(campaign);
    sheetTime.getRange(`D${targetRow}`).setValue(project);
    sheetTime.getRange(`E${targetRow}`).setValue(activity);
    sheetTime.getRange(`F${targetRow}`).setValue(now).setNumberFormat("d mmmm yyyy");
    sheetTime.getRange(`G${targetRow}`).setValue(currentTimeValue).setNumberFormat("HH:mm");
    sheetTime.getRange(`H${targetRow}`).clearContent();
    sheetTime.getRange(`I${targetRow}`).setFormula(buildTimeFormula(targetRow));
    sheetTime.getRange(`J${targetRow}`).setFormula(buildHoursCumulativeFormula(targetRow));
    sheetTime.getRange(`K${targetRow}`).setFormula(buildAmountFormula(targetRow));
    sheetTime.getRange(`L${targetRow}`).setFormula(buildAmountCumulativeFormula(targetRow));
    sheetTime.getRange(`N${targetRow}`).setValue(false);
    sheetTime.getRange(`O${targetRow}`).setValue(false);
    sheetTime.getRange(`P${targetRow}`).clearContent();
    sheetTime.getRange(`Q${targetRow}`).clearContent();
    sheetTime.getRange(`R${targetRow}`).setValue(false);
    sheetTime.getRange(`S${targetRow}`).clearContent();
    sheetTime.getRange(`T${targetRow}`).setValue(rate);
    sheetTime.getRange(`U${targetRow}`).setValue(note);
    applyTimeEntryRowFormatting(targetRow);
    applyTimeColumnsDisplayFormat();

    sheetTime.getRange(`A${checkedRowIndex}`).setValue(false);

    const rangeEffet = sheetTime.getRange(`A${targetRow}:Z${targetRow}`);
    rangeEffet.setBackground("#f1f6ee");
    sheetTime.getRange("I3").setBackground("#6aa84f");
    sheetTime.setActiveSelection(`H${targetRow}`);
    SpreadsheetApp.flush();

  } else {
    // Aucune case cochée, insérer à ligne 7
    sheetTime.insertRowsAfter(6, 2);
    sheetTime.getRange("A7:U7").clearContent();
    sheetTime.getRange("A8:U8").clearContent();
    sheetTime.getRange("A7").insertCheckboxes();
    sheetTime.getRange("N7").insertCheckboxes();
    sheetTime.getRange("O7").insertCheckboxes();
    sheetTime.getRange("R7").insertCheckboxes();
    sheetTime.getRange("H7").clearContent();
    sheetTime.getRange("B7").setValue(client);
    sheetTime.getRange("C7").setValue(campaign);
    sheetTime.getRange("D7").setValue(project);
    sheetTime.getRange("E7").setValue(activity);
    sheetTime.getRange("F7").setValue(now).setNumberFormat("d mmmm yyyy");
    sheetTime.getRange("G7").setValue(currentTimeValue).setNumberFormat("HH:mm");
    sheetTime.getRange("I7").setFormula(buildTimeFormula(7));
    sheetTime.getRange("J7").setFormula(buildHoursCumulativeFormula(7));
    sheetTime.getRange("K7").setFormula(buildAmountFormula(7));
    sheetTime.getRange("L7").setFormula(buildAmountCumulativeFormula(7));
    sheetTime.getRange("N7").setValue(false);
    sheetTime.getRange("O7").setValue(false);
    sheetTime.getRange("P7").clearContent();
    sheetTime.getRange("Q7").clearContent();
    sheetTime.getRange("R7").setValue(false);
    sheetTime.getRange("S7").clearContent();
    sheetTime.getRange("T7").setValue(rate);
    sheetTime.getRange("U7").setValue(note);
    sheetTime.getRange("A7").setValue(true);
    applyTimeEntryRowFormatting(7);
    applyTimeColumnsDisplayFormat();

    const rangeEffet = sheetTime.getRange("A7:Z7");
    rangeEffet.setBackground("#f1f6ee");
    sheetTime.getRange("I3").setBackground("#6aa84f");
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
  const lastRow = sheet.getLastRow();
  const data = lastRow >= 7 ? sheet.getRange(`A7:S${lastRow}`).getValues() : [];
  const checkedRows = data
    .map((row, index) => ({ row, index: index + 7 }))
    .filter(entry => entry.row[0] === true);
  const ligneCocheeCount = checkedRows.length;

  if (ligneCocheeCount === 0) {
    openStandaloneMessageView_("Aucune ligne cochée à supprimer.");
    return;
  }

  const hasInvoicedOrPaidRow = checkedRows.some(({ row }) => row[14] === true || row[17] === true || row[18] !== "" && row[18] !== null);
  const confirmTitle = hasInvoicedOrPaidRow ? "Attention" : `Supprimer ${ligneCocheeCount} entrée${ligneCocheeCount > 1 ? 's' : ''}`;
  const confirmMessage = hasInvoicedOrPaidRow
    ? "Attention, au moins une entrée sélectionnée a déjà été facturée ou payée."
    : ligneCocheeCount === 1
      ? "Voulez-vous vraiment supprimer cette entrée ?"
      : `Voulez-vous vraiment supprimer ces ${ligneCocheeCount} entrées ?`;
  const confirmSecondaryLabel = hasInvoicedOrPaidRow ? "Fermer" : "Annuler";
  const template = HtmlService.createTemplateFromFile("popup");
  template.contacts = [];
  template.activityTypes = [];
  template.initialInvoiceNumber = null;
  template.requiresInitialInvoiceSetup = false;
  template.showStartupConfirm = true;
  template.confirmViewTitle = confirmTitle;
  template.confirmViewMessage = confirmMessage;
  template.confirmAction = "delete_rows";
  template.confirmPrimaryLabel = "Supprimer";
  template.confirmSecondaryLabel = confirmSecondaryLabel;
  template.facturationPopupContextReady = true;
  const output = template.evaluate()
    .setWidth(400)
    .setHeight(78);
  output.addMetaTag('viewport', 'width=device-width, initial-scale=1');

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

  // ⬜️ Effet visuel gris temporaire
  lignesASupprimer.forEach(row => {
    const range = sheet.getRange(row, 1, 1, sheet.getLastColumn());
    range.setBackground("#cccccc").setFontColor("#000000");
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

  const applyTimeColumnsDisplayFormat = () => {
    const lastTimeRow = Math.max(7, sheetTime.getLastRow());
    sheetTime.getRange(`G7:H${lastTimeRow}`).setNumberFormat("HH:mm");
  };

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
    openStandaloneMessageView_("Ne sélectionner qu'une tâche active.", "Information");
    return;
  }

  if (checkedRows.length === 0) {
    openStandaloneMessageView_("Sélectionner une tâche active.");
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
      const currentTimeValue = (now.getHours() * 60 + now.getMinutes()) / (24 * 60);
      cellH.setValue(currentTimeValue).setNumberFormat("HH:mm");
      applyTimeColumnsDisplayFormat();
      sheetTime.getRange(`A${rowIndex}:Z${rowIndex}`).setBackground("#ffffff");
      // Changer la couleur de I3 en blanc
      sheetTime.getRange("I3").setBackground("#ffffff");
      SpreadsheetApp.flush();
    } else {
      // Cellule H non vide, afficher popup
      openStandaloneMessageView_("La tâche sélectionnée a déjà été fermée.", "Action impossible sur cette ligne");
    }
  }
}

// INFO : Change les coordonnées de l'entreprise

function info() {
  showInvoiceConfigurationDialog_({ continueFacturerAfterSave: false });
}

function submitInfoForm(name, address1, address2, address3, address4, email, website, nextInvoice) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const sheetGestion = ss.getSheetByName("GESTION");
  if (!sheetGestion) throw new Error("Feuille GESTION introuvable.");

  const requiredFields = [name, email, address1, address2];
  if (requiredFields.some(value => String(value || "").trim() === "")) {
    return { success: false, message: "Veuillez remplir tous les champs obligatoires." };
  }

  sheetGestion.getRange("F2:F8").setValues([
    [name],
    [address1],
    [address2],
    [address3],
    [address4],
    [email],
    [website]
  ]);
  return { success: true };
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
