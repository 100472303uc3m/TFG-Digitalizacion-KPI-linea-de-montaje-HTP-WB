// Extract sheet Id and store it in global variable to send as URL parameter
const sheetId = SpreadsheetApp.getActiveSpreadsheet().getId();

// Create and add custom menu on sheet when opened
function onOpen() {
  var ui = SpreadsheetApp.getUi();
  // Or DocumentApp, SlidesApp or FormApp.
  ui.createMenu('Open in App')
    .addSubMenu(ui.createMenu('Workload')
      .addItem('100', 'openWorkload100Link')
      .addItem('101', 'openWorkload101Link')
      .addItem('102', 'openWorkload102Link')
      .addItem('A320', 'openWorkloadA320Link')
      .addItem('A330', 'openWorkloadA330Link')
      .addItem('A350', 'openWorkloadA350Link')
      .addItem('S19', 'openWorkloadS19Link'))
    .addSubMenu(ui.createMenu('Competitividad')
      .addItem('100', 'openCompetitividad100Link')
      .addItem('101', 'openCompetitividad101Link')
      .addItem('102', 'openCompetitividad102Link')
      .addItem('103', 'openCompetitividad103Link'))
    .addSeparator()
    .addItem('Configuración de enlaces', 'openDialog')
    .addToUi();
};

// Functions to open pages in web page
function openWorkload100Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?program=first&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openWorkload101Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?program=second&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openWorkload102Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?program=third&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openWorkloadA320Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?program=fourth&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openWorkloadA330Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?program=fifth&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openWorkloadA350Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?program=sixth&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openWorkloadS19Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?program=seventh&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openCompetitividad100Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?mode=Competitividad&program=first&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openCompetitividad101Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?mode=Competitividad&program=second&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openCompetitividad102Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?mode=Competitividad&program=third&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

function openCompetitividad103Link() {
  const html = HtmlService.createHtmlOutput('<script>window.open("https://script.google.com/a/macros/alumnos.uc3m.es/s/AKfycbyelu2v32m2utS6b9CKVVnLvh0mtygA5BX_YhJmPHsO/dev?mode=Competitividad&program=fourth&sheetId=' + sheetId + '", "_blank");google.script.host.close();</script>');
  SpreadsheetApp.getUi().showModalDialog(html, "Opening...");
};

// Function that opens dialog
function openDialog() {
  var html = HtmlService.createHtmlOutputFromFile('Dialog')
      .setWidth(700)
      .setHeight(310);
  SpreadsheetApp.getUi().showModalDialog(html, 'Configuración de enlaces');
};

// Function that reads links from link sheet
function readLinks() {
  // Acccess sheet with links
  var ss = SpreadsheetApp.openById('1rHltJDUtCu6zGz-jziYVCfU5ZxQofNAHeOdsP9KYLGo');
  
  // Extract links from the sheet
  var links = ss.getRange('Enlaces!D6:D9').getValues();

  // Return links - If no links are found it will return an empty array
  return links;
};

// Function that edits links in link sheet
function editLinks(newLinks) {
  // Acccess sheet with links
  var ss = SpreadsheetApp.openById('1rHltJDUtCu6zGz-jziYVCfU5ZxQofNAHeOdsP9KYLGo');

  // Write new links into link sheet
  ss.getRange('Enlaces!D6:D9').setValues(newLinks);
};

