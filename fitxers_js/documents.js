function generarDocumentEntrada(dades) {
  const ID_PLANTILLA = "1a3dnY5J1vBT51bx1frg4iFgfIcYGkj1y9iCwgjHYttA";
  const ID_CARPETA = "1x3iY3aUgd5z2cykWIw2gca57Fj1jf-_-";

  // 1. Obrim la còpia de la plantilla
  const carpeta = obtenirCarpetaAula(ID_CARPETA, dades.aula);
  const nomDocument = `${dades.cognom1} ${dades.cognom2}, ${dades.nom} - ${dades.data} ${dades.hora}`;

  const copia = DriveApp.getFileById(ID_PLANTILLA).makeCopy(
    nomDocument,
    carpeta,
  );

  const doc = DocumentApp.openById(copia.getId());
  const body = doc.getBody();

  dades.registre = obtenirSeguentRegistreEntrada();

  // 2. Substituïm el text de les etiquetes normals
  body.replaceText("{{REGISTRE}}", dades.registre);
  body.replaceText("{{DATA}}", dades.data);
  body.replaceText("{{HORA}}", dades.hora);
  body.replaceText("{{ALUMNE}}", dades.alumne);
  body.replaceText("{{AULA}}", dades.aula);
  body.replaceText("{{TUTOR}}", dades.tutor);
  body.replaceText("{{OBSERVACIONS}}", dades.observacions || "");

  // 3. Inserim el logotip a la capçalera o cos
  inserirLogotip(doc);

  // 4. Inserim la signatura (mida adaptada)
  if (dades.signatura) {
    const trobat = body.findText("{{SIGNATURA}}") || body.findText("SIGNATURA");
    if (trobat) {
      const parrafSig = trobat.getElement().getParent().asParagraph();
      body.replaceText("{{SIGNATURA}}", "");
      const base64Sig = dades.signatura.includes(",")
        ? dades.signatura.split(",")[1]
        : dades.signatura;

      const blobSig = Utilities.newBlob(
        Utilities.base64Decode(base64Sig),
        "image/png",
        "signatura.png",
      );
      parrafSig.appendInlineImage(blobSig).setWidth(110).setHeight(45);
    }
  }

  // 5. Inserim el justificant (mida adaptada)
  if (dades.justificant && dades.justificant.dades) {
    Logger.log("📷 Justificant rebut: " + dades.justificant.nom);
    const trobat =
      body.findText("{{JUSTIFICANT}}") || body.findText("JUSTIFICANT");
    if (trobat) {
      Logger.log("✅ Etiqueta {{JUSTIFICANT}} localitzada al document.");
      const parrafJust = trobat.getElement().getParent().asParagraph();
      body.replaceText("{{JUSTIFICANT}}", "");

      // Extraiem les dades Base64
      const base64Data = dades.justificant.dades.includes(",")
        ? dades.justificant.dades.split(",")[1]
        : dades.justificant.dades;

      const blobJust = Utilities.newBlob(
        Utilities.base64Decode(dades.justificant.dades.split(",")[1]),
        dades.justificant.tipus,
        dades.justificant.nom,
      );
      parrafJust.appendInlineImage(blobJust).setWidth(180);
    } else {
      Logger.log(
        "⚠️ ERROR: No s'ha trobat el text {{JUSTIFICANT}} a la plantilla de Google Docs.",
      );
    }
  } else {
    Logger.log(
      "ℹ️ No s'ha adjuntat cap cap fitxer de justificant a la petició.",
    );
  }

  // 6. Desem i convertim a PDF
  doc.saveAndClose();

  const pdf = DriveApp.getFileById(doc.getId()).getAs(MimeType.PDF);
  const nom = `${dades.data}_${dades.hora}_${dades.aula}_${dades.cognom1}_${dades.nom}_Entrada.pdf`;
  const fitxer = carpeta.createFile(pdf).setName(nom);

  // Esborrar la còpia temporal del document de text
  DriveApp.getFileById(doc.getId()).setTrashed(true);

  // Actualitzar minuts d'absència
  if (dades.aula && dades.fila) {
    actualitzarMinutsAbsencia(dades.aula, dades.fila);
  }

  return fitxer.getUrl();
}

function inserirSignatura(body, base64) {
  Logger.log("Longitud base64: " + base64.length);

  const bytes = Utilities.base64Decode(base64.split(",")[1]);

  Logger.log("Bytes: " + bytes.length);

  const blob = Utilities.newBlob(bytes, "image/png", "signatura.png");

  Logger.log("Blob mida: " + blob.getBytes().length);

  const etiqueta = body.findText("{{SIGNATURA}}");

  Logger.log("Etiqueta trobada: " + (etiqueta != null));

  if (!etiqueta) return;

  const element = etiqueta.getElement();

  const parraf = element.getParent().asParagraph();

  element.asText().setText("");

  parraf.appendInlineImage(blob).setWidth(200);

  Logger.log("Imatge inserida");
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

    body.appendImage(blob).setWidth(200);

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

function inserirLogotip(doc) {
  // Buscar primer a la capçalera; si no n'hi ha, busca al cos
  const header = doc.getHeader();
  const seccio = header || doc.getBody();

  const etiqueta = seccio.findText("{{LOGOTIP}}");

  if (!etiqueta) return;

  const element = etiqueta.getElement();

  const parraf = element.getParent().asParagraph();

  element.asText().setText("");

  // Descarreguem la imatge i la convertim en Blob
  if (typeof URL_LOGOTIP !== "undefined" && URL_LOGOTIP) {
    const blob = UrlFetchApp.fetch(URL_LOGOTIP).getBlob();

    parraf.appendInlineImage(blob).setWidth(90).setHeight(90);
  }
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
