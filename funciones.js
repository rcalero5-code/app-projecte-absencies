const full = SpreadsheetApp.openById('1FzOhlhpPKt8nESIGZRkVAJF4ErkjYuYOcleWpskB_60');

function doGet() {
    return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Absències').setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3135/3135715.png');
}

function obtenirDadesHTML(nom) {
    return HtmlService.createHtmlOutputFromFile(nom).getContent();
}

function obtenirAlumnes(aula) {
    const fulla = full.getSheetByName(aula);
    const alumnes = fulla.getDataRange().getValues();
    return alumnes;
}

function obtenirAules() {
    return full.getSheets().map(f => f.getName());
}