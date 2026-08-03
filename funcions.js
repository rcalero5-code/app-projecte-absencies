const FULL = SpreadsheetApp.openById(
  "1FzOhlhpPKt8nESIGZRkVAJF4ErkjYuYOcleWpskB_60",
);

const ID_FULL_ABSENCIES = SpreadsheetApp.openById(
  "19Dupkk8Y_3oEjF3XvAfDrhCgOyWiH5F4kZX3mnVcmmA",
);

const FULL_ABSENCIES = SpreadsheetApp.openById(
  "1teCIFUY-HBG7Xv3n0Cb3uJFX-pSoilPihXsKKGPlY00",
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

function obtenirDadesHTML(fitxer) {
  return HtmlService.createHtmlOutputFromFile(fitxer).getContent();
}

function obtenirAlumnes(aula) {
  Logger.log("Aula rebuda: " + aula);

  const fulla = SpreadsheetApp.openById(
    "1FzOhlhpPKt8nESIGZRkVAJF4ErkjYuYOcleWpskB_60",
  ).getSheetByName(aula);

  if (!fulla) {
    throw new Error("No existeix la fulla: [" + aula + "]");
  }

  return fulla.getDataRange().getValues();
}

function obtenirAules() {
  return FULL.getSheets().map((f) => f.getName());
}

function obtenirAlumnesAbsents(aula) {
  const full = FULL_ABSENCIES.getSheetByName(aula);

  if (!full) {
    throw new Error("No existeix la fulla " + aula);
  }

  const dades = full.getDataRange().getValues();

  // Dia d'avui
  const dia = new Date().getDate();

  // Buscar la columna del dia (fila 3)
  let columnaDia = -1;

  for (let c = 0; c < dades[2].length; c++) {
    if (Number(dades[2][c]) === dia) {
      columnaDia = c;
      break;
    }
  }

  if (columnaDia === -1) {
    throw new Error("No s'ha trobat la columna del dia " + dia);
  }

  alumnes = [];

  // Les dades comencen a la fila 4
  for (let i = 3; i < dades.length; i++) {
    const absencia = dades[i][columnaDia];

    if (absencia !== "") {
      alumnes.push({
        fila: i + 1,
        cognom1: dades[i][0],
        cognom2: dades[i][1],
        nom: dades[i][2],
      });
    }
  }

  return alumnes;
}

function obtenirTutors(aula, fila) {
  const fullAlumnes = FULL.getSheetByName(aula);

  if (!fullAlumnes) {
    throw new Error("No existeix la fulla " + aula);
  }

  const filaLlista = Number(fila) - 2;

  Logger.log("Fila FULL_ABSENCIES: " + fila);
  Logger.log("Fila LLISTA: " + filaLlista);

  const dades = fullAlumnes
    .getRange(filaLlista, 1, 1, fullAlumnes.getLastColumn())
    .getValues()[0];

  Logger.log(dades);

  return {
    tutor1: dades[5],
    tutor2: dades[6],
  };
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
          alumne.cognom1,
          alumne.cognom2,
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

  enviarDadesAbsencies(absencies);
}

function enviarDadesAbsencies(absencies) {
  absencies.forEach((alumne) => {
    let nom = alumne.nom;
    let cognom1 = alumne.cognom1;
    let cognom2 = alumne.cognom2;
    let aula = alumne.aula;
    let absencia = alumne.absencia;
    let justificacio = alumne.justificacio;
    let comentari = alumne.comentari;
    let dia = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "d",
    );
    let mesNombre = Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "M",
    );

    let data = dia + "/" + mesNombre;
    console.log("Data: ", data);

    // Obrir la fulla de càlcul alumnes
    let spreadsheetAlumnes = FULL_ABSENCIES;

    if (spreadsheetAlumnes !== null) {
      let sheetAlumnes = spreadsheetAlumnes.getSheetByName(aula);

      if (!sheetAlumnes) {
        console.log("No existeix la fulla " + aula);
        return;
      }
      // Trobar la fila corresponent a l'alumne
      const dades = sheetAlumnes.getDataRange().getValues();

      let filaAlumne = -1;

      for (let i = 3; i < dades.length; i++) {
        if (
          dades[i][0] === cognom1 &&
          dades[i][1] === cognom2 &&
          dades[i][2] === nom
        ) {
          filaAlumne = i + 1;
          break;
        }
      }

      if (filaAlumne > 0) {
        // Buscar la columna del dia
        const dies = sheetAlumnes
          .getRange(3, 1, 1, sheetAlumnes.getLastColumn())
          .getValues()[0];

        let columnaDia = -1;

        for (let c = 0; c < dies.length; c++) {
          if (Number(dies[c]) === Number(dia)) {
            columnaDia = c + 1;
            break;
          }
        }

        if (columnaDia > 0) {
          // Escriure els minuts d'absència
          sheetAlumnes.getRange(filaAlumne, columnaDia).setValue(380);

          // Afegir comentari si existeix
          if (comentari) {
            sheetAlumnes.getRange(filaAlumne, columnaDia).setNote(comentari);
          }
        } else {
          console.log("No s'ha trobat la columna del dia.");
        }
      } else {
        console.log("No s'ha trobat l'alumne.");
      }
    }
  });
}

function desarEntrada(dades) {
  Logger.log(JSON.stringify(dades));

  // De moment només comprovem que arriba
  return true;
}

function obtenirAbsenciaColumna(absencia, justificacio) {
  return justificacio === "si" ? "J" : ""; //Operador Ternari
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

function generarDocumentEntrada(dades) {
  const ID_PLANTILLA = "1a3dnY5J1vBT51bx1frg4iFgfIcYGkj1y9iCwgjHYttA";
  const ID_CARPETA = "1x3iY3aUgd5z2cykWIw2gca57Fj1jf-_-";

  // Copiar plantilla
  const carpeta = obtenirCarpetaAula(ID_CARPETA, dades.aula);

  const nomDocument = `${dades.cognom1} ${dades.cognom2}, ${dades.nom} - ${dades.data} ${dades.hora}`;

  const copia = DriveApp.getFileById(ID_PLANTILLA).makeCopy(
    nomDocument,
    carpeta,
  );

  const doc = DocumentApp.openById(copia.getId());

  const body = doc.getBody();

  dades.registre = obtenirSeguentRegistreEntrada();

  Logger.log("data = " + dades.data);
  Logger.log("hora = " + dades.hora);
  Logger.log("nom = " + dades.nom);
  Logger.log("cognom1 = " + dades.cognom1);
  Logger.log("cognom2 = " + dades.cognom2);
  Logger.log("alumne = " + dades.alumne);

  body.replaceText("{{REGISTRE}}", dades.registre);
  body.replaceText("{{DATA}}", dades.data);
  body.replaceText("{{HORA}}", dades.hora);
  body.replaceText("{{ALUMNE}}", dades.alumne);
  body.replaceText("{{AULA}}", dades.aula);
  body.replaceText("{{TUTOR}}", dades.tutor);
  body.replaceText("{{OBSERVACIONS}}", dades.observacions || "");

  inserirSignatura(body, dades.signatura);

  if (dades.justificant) {
    inserirJustificant(body, dades.justificant, carpeta);
  }

  inserirLogotip(body);

  doc.saveAndClose();

  const pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);

  Logger.log(
    `${dades.data}_${dades.hora}_${dades.aula}_${dades.cognom1}_${dades.nom}`,
  );

  const nom = `${dades.data}_${dades.hora}_${dades.aula}_${dades.cognom1}_${dades.nom}_Entrada.pdf`;

  const fitxer = carpeta.createFile(pdf).setName(nom);

  DriveApp.getFileById(doc.getId()).setTrashed(true);

  return fitxer.getUrl();
}

function inserirSignatura(body, base64) {
  const etiqueta = body.findText("{{SIGNATURA}}");

  if (!etiqueta) return;

  const element = etiqueta.getElement();

  const parraf = element.getParent().asParagraph();

  element.asText().setText("");

  const bytes = Utilities.base64Decode(base64.split(",")[1]);

  const blob = Utilities.newBlob(bytes, "image/png", "signatura.png");

  parraf.appendInlineImage(blob).setWidth(180).setHeight(70);
}

function inserirJustificant(body, justificant, carpeta) {
  const tipus = justificant.tipus;
  const bytes = Utilities.base64Decode(justificant.dades.split(",")[1]);

  body.appendParagraph("");
  body
    .appendParagraph("JUSTIFICANT")
    .setHeading(DocumentApp.ParagraphHeading.HEADING2);

  // ---------- IMATGES ----------
  if (tipus.startsWith("image/")) {
    const blob = Utilities.newBlob(bytes, tipus, justificant.nom);

    body.appendImage(blob).setWidth(450);

    return;
  }

  // ---------- PDF ----------
  if (tipus === "application/pdf") {
    const pdf = Utilities.newBlob(bytes, MimeType.PDF, justificant.nom);

    const fitxer = carpeta.createFile(pdf);

    const p = body.appendParagraph("Justificant adjunt:");

    p.appendText(" ");
    p.appendText(fitxer.getName()).setLinkUrl(fitxer.getUrl());

    return;
  }

  body.appendParagraph("Tipus de fitxer no suportat.");
}

function inserirLogotip(body) {
  const etiqueta = body.findText("{{LOGOTIP}}");

  if (!etiqueta) return;

  const element = etiqueta.getElement();

  const parraf = element.getParent().asParagraph();

  element.asText().setText("");

  const blob = DriveApp.getFileById("ID_LOGOTIP").getBlob();

  parraf.appendInlineImage(blob).setWidth(90).setHeight(90);
}

function obtenirCarpetaAula(idCarpetaPare, aula) {
  const carpetaPare = DriveApp.getFolderById(idCarpetaPare);

  const carpetes = carpetaPare.getFoldersByName(aula);

  if (carpetes.hasNext()) {
    return carpetes.next();
  }

  return carpetaPare.createFolder(aula);
}

function obtenirSeguentRegistreEntrada() {
  const CURS = "2627";

  const props = PropertiesService.getScriptProperties();

  let comptador = Number(props.getProperty("REG_ENT_" + CURS)) || 0;

  comptador++;

  props.setProperty("REG_ENT_" + CURS, comptador);

  return `REG_ENT_${CURS}_${String(comptador).padStart(3, "0")}`;
}

function obtenirAlumne(aula, fila) {
  const full = FULL.getSheetByName(aula);

  const dades = full.getRange(fila - 2, 1, 1, 7).getValues()[0];

  return {
    cognom1: dades[0],
    cognom2: dades[1],
    nom: dades[2],
    telefon1: dades[3],
    telefon2: dades[4],
    tutor1: dades[5],
    tutor2: dades[6],
  };
}
