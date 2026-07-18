const full = SpreadsheetApp.openById(
  "1FzOhlhpPKt8nESIGZRkVAJF4ErkjYuYOcleWpskB_60",
);

const ID_FULL_ABSENCIES = SpreadsheetApp.openById(
  "19Dupkk8Y_3oEjF3XvAfDrhCgOyWiH5F4kZX3mnVcmmA",
);

const CAPCALERA = [
  [
    "COGNOM",
    "SEGON COGNOM",
    "NOM",
    "AULA",
    "ABSÈNCIA",
    "JUSTIFICACIÓ",
    "TELÈFONS",
    "OBSERVACIONS",
    "MISSATGE",
  ],
];

function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Absències")
    .setFaviconUrl("https://cdn-icons-png.flaticon.com/512/3135/3135715.png");
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
  return full.getSheets().map((f) => f.getName());
}

function enviarAbsencies(absencies) {
  const full = ID_FULL_ABSENCIES;

  let dia = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "d");

  let fulla = full.getSheetByName(dia);

  // Comprovem si existeix dia sinó la creem
  if (!fulla) {
    fulla = crearFullaDia(full, dia);
    crearCapcalera(fulla);
    configurarFulla(fulla);
  }
  absencies.forEach((alumne) => {
    const filaNova = fulla.getLastRow() + 1;

    fulla
      .getRange(filaNova, 1, 1, 9)
      .setValues([
        [
          alumne.cognom,
          alumne.segonCognom,
          alumne.nom,
          alumne.aula,
          alumne.absencia,
          alumne.justificacio,
          `${alumne.telefon1} / ${alumne.telefon2}`,
          alumne.comentari,
          false,
        ],
      ]);

    fulla.getRange(filaNova, 9).insertCheckboxes();
  });
}

function crearFullaDia(full, dia) {
  return full.insertSheet(dia);
}

function crearCapcalera(fulla) {
  // Configurar la capçalera
  fulla
    .getRange(1, 1, 1, CAPCALERA[0].length)
    .setValues(CAPCALERA)
    .setBorder(true, true, true, true, true, true)
    .setBackground("#ececec")
    .setFontColor("#32CD32")
    .setHorizontalAlignment("center")
    .setVerticalAlignment("middle")
    .setFontWeight("bold")
    .setFontFamily("Montserrat")
    .setWrap(true); // Justificar el text
}

function configurarFulla(fulla) {
  // Afegir una regla de formatació condicional si la casella de la columna I és TRUE
  const rangeToFormat = fulla.getRange(2, 1, fulla.getMaxRows() - 1, 9); // Tota la taula de dades, excepte capçalera

  let rule = SpreadsheetApp.newConditionalFormatRule()
    .whenFormulaSatisfied("=$I2=TRUE")
    .setBackground("#D6EAF8") // Color blau clar (pots canviar-ho)
    .setRanges([rangeToFormat])
    .build();

  let rules = fulla.getConditionalFormatRules();
  rules.push(rule);
  fulla.setConditionalFormatRules(rules);

  // Ajustar les mides de les columnes
  const amplades = [100, 100, 100, 50, 30, 30, 300, 300, 80];

  amplades.forEach((ample, i) => {
    fulla.setColumnWidth(i + 1, ample);
  });
}
