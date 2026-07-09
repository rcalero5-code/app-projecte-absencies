function doGet() {
    return HtmlService.createTemplateFromFile('index').evaluate().setTitle('Absències').setFaviconUrl('https://cdn-icons-png.flaticon.com/512/3135/3135715.png');
}

function obtenerDatosHTML(nombre) {
    return HtmlService.createHtmlOutputFromFile(nombre).getContent();
}