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

function obtenirAlumnes(aula) {
  Logger.log("Aula rebuda: [" + aula + "]");
  Logger.log("Spreadsheet: " + FULL.getName());

  const fulla = FULL.getSheetByName(aula);

  Logger.log(fulla);

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

  const alumnes = [];

  // Les dades comencen a la fila 4
  for (let i = 3; i < dades.length; i++) {
    const absencia = dades[i][columnaDia];

    if (absencia !== "") {
      alumnes.push({
        fila: i + 1,
        cognom1: dades[i][0],
        cognom2: dades[i][1],
        nom: dades[i][2],
        minuts: Number(dades[i][columnaDia]) || 380
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

function actualitzarMinutsAbsencia(aula, fila) {

  const sheet = FULL_ABSENCIES.getSheetByName(aula);

  if (!sheet) {
    throw new Error("No existeix la fulla " + aula);
  }

  const dia = new Date().getDate();

  const dies = sheet
    .getRange(3, 1, 1, sheet.getLastColumn())
    .getValues()[0];

  let columnaDia = -1;

  for (let c = 0; c < dies.length; c++) {
    if (Number(dies[c]) === dia) {
      columnaDia = c + 1;
      break;
    }
  }

  if (columnaDia === -1) {
    throw new Error("No s'ha trobat la columna del dia");
  }

  const HORA_INICI = 8 * 60 + 15;
  const HORA_FINAL = 14 * 60 + 35;

  const ara = new Date();

  const horaActual = ara.getHours() * 60 + ara.getMinutes();

  const minuts = Math.max(
    0,
    Math.min(horaActual - HORA_INICI, HORA_FINAL - HORA_INICI)
  );

  sheet.getRange(fila, columnaDia).setValue(minuts);

  return minuts;
}

function desarEntrada(dades) {
  Logger.log("Processant registre d'entrada: " + JSON.stringify(dades));

  try {
    // Generem el document PDF i obtenim la seva URL de Drive
    const urlPdf = generarDocumentEntrada(dades);
    return { status: "success", url: urlPdf };
  } catch (error) {
    Logger.log("Error a desarEntrada: " + error.message);
    throw new Error(
      "No s'ha pogut generar el registre d'entrada: " + error.message,
    );
  }
}

function obtenirAbsenciaColumna(absencia, justificacio) {
  return justificacio === "si" ? "J" : ""; //Operador Ternari
}

function obtenirAlumne(aula, fila) {
  Logger.log("===== OBTENIR ALUMNE =====");
  Logger.log("Aula: [" + aula + "]");
  Logger.log("Fila: " + fila);

  const full = FULL.getSheetByName(aula);

  Logger.log("Full: " + full);

  if (!full) {
    throw new Error("No existeix la fulla [" + aula + "]");
  }

  const dades = full.getRange(fila - 2, 1, 1, 7).getValues()[0];

  Logger.log(dades);

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

/**
 * Retorna els alumnes de l'aula combinats amb el seu estat d'absència d'avui
 */
function obtenirDadesAulaAmbProgres(aula) {
  const alumnes = obtenirAlumnes(aula); // La teva funció existent
  const absents = obtenirAlumnesAbsents(aula); // Els absents d'avui
  
  // Creem un conjunt (Set) de files d'alumnes absents per a una cerca ràpida
  const filesAbsents = new Set(absents.map(a => Number(a.fila)));

  // Retornem la llista indicant quins alumnes ja tenen l'absència registrada
  return alumnes.slice(1).map((row, index) => {
    const numFila = index + 2; // +2 per compensar capçaleres/índex
    return {
      fila: numFila,
      cognom1: row[0],
      cognom2: row[1],
      nom: row[2],
      telefon1: row[3],
      telefon2: row[4],
      isAbsent: filesAbsents.has(numFila)
    };
  });
}

/**
 * Retorna la llista d'alumnes d'una aula fusionada amb l'estat d'absència d'avui.
 */
function obtenirAlumnesAmbProgres(aula) {
  const fullaBase = FULL.getSheetByName(aula);
  if (!fullaBase) throw crearNotificacioError(`No existeix l'aula ${aula}`, "ERROR");
  
  const dadesAlumnes = fullaBase.getDataRange().getValues();
  const fullaAbs = FULL_ABSENCIES.getSheetByName(aula);
  
  const diaAvui = new Date().getDate();
  const minutsTotalsJornada = 380; // De 8:15 a 14:35h
  const mapaAbsencies = {};

  // Llegim les absències d'avui si la fulla existeix
  if (fullaAbs) {
    const dadesAbs = fullaAbs.getDataRange().getValues();
    const diesFila = dadesAbs[2]; // La fila 3 conté els dies del mes
    
    let colDia = -1;
    for (let c = 0; c < diesFila.length; c++) {
      const val = diesFila[c];
      // Permet detectar tant si la cel·la és un número com un objecte Data de Sheets
      let diaVal = (val instanceof Date) ? val.getDate() : Number(val);

      if (diaVal === diaAvui) {
        colDia = c;
        break;
      }
    }

    if (colDia !== -1) {
      for (let i = 3; i < dadesAbs.length; i++) {
        // .trim() i .toLowerCase() per evitar fallades si hi ha espais amagats
        const c1 = String(dadesAbs[i][0] || "").trim().toLowerCase();
        const c2 = String(dadesAbs[i][1] || "").trim().toLowerCase();
        const nom = String(dadesAbs[i][2] || "").trim().toLowerCase();
        
        const clau = `${c1}_${c2}_${nom}`;
        const val = dadesAbs[i][colDia];
        if (val !== "" && val !== null && val !== undefined) {
          mapaAbsencies[clau] = Number(val) || 380;
        }
      }
    }
  }

  // Fusionem la llista d'alumnes amb les absències registrades
  return dadesAlumnes.slice(1).map((row, index) => {
    const c1 = String(row[0] || "").trim().toLowerCase();
    const c2 = String(row[1] || "").trim().toLowerCase();
    const nom = String(row[2] || "").trim().toLowerCase();
    const clau = `${c1}_${c2}_${nom}`;

    const minuts = mapaAbsencies[clau] || 0;
    const percentatge = Math.min(100, Math.round((minuts / minutsTotalsJornada) * 100));

    return {
      fila: index + 1,
      cognom1: row[0],
      cognom2: row[1],
      nom: row[2],
      telefon1: row[3],
      telefon2: row[4],
      isAbsent: minuts > 0,
      minutsAbsencia: minuts,
      percentatge: percentatge
    };
  });
}
