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

  parraf.appendInlineImage(blob).setWidth(180).setHeight(70);

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

