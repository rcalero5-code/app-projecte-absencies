function doGet() {
  return HtmlService.createTemplateFromFile("index")
    .evaluate()
    .setTitle("Absències")
    .setFaviconUrl("https://cdn-icons-png.flaticon.com/512/3135/3135715.png");
}

function obtenirDadesHTML(fitxer) {
  return HtmlService.createHtmlOutputFromFile(fitxer).getContent();
}