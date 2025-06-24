const lookURL = (soughtValue, fil, col, fil2, col2) => {
  const sheet = SpreadsheetApp.getActive().getSheetByName('Enlaces de acceso');
  let comparativeValue;
  while ((comparativeValue = sheet.getRange(fil, col).getValue()) !== "") {
    if (soughtValue === comparativeValue) {
      return sheet.getRange(fil2, col2).getValue();
    }
    fil++;
    fil2++;
  }
  return null;
};

const emailOnFormSubmit = (e) => {
  const userResponse = {
    timestamp: e.values[0],
    email: e.values[1],
    service: e.values[2],
    title: e.values[3],
    datetime: e.values[4] + e.values[7] + e.values[10] + e.values[13] + e.values[16] + e.values[19] + e.values[22] + e.values[25],
    fullname: e.values[5] + e.values[8] + e.values[11] + e.values[14] + e.values[17] + e.values[20] + e.values[23] + e.values[26],
    businessUnit: e.values[6] + e.values[9] + e.values[12] + e.values[15] + e.values[18] + e.values[21] + e.values[24] + e.values[27],
    dni: e.values[28] + e.values[32] + e.values[36] + e.values[39],
    studiesModality: e.values[29] + e.values[33] + e.values[37],
    academicProgram: e.values[30] + e.values[34] + e.values[38],
    userType: e.values[31] + e.values[35],
    campus: e.values[40],
    phone: e.values[41],
    dpp: e.values[42],
  };

  const FULL_TITLE = (userResponse.service == "Capacitación") ? userResponse.title + " / " + userResponse.datetime : "Inducción / " + userResponse.datetime;
  const MEET_URL = lookURL(FULL_TITLE, 1, 1, 1, 2);
  const DESCRIPTION = `<p>Título: <strong>${(userResponse.service == "Capacitación") ? userResponse.title : 'Inducción'}</strong></p>
<p>Fecha: <strong>${userResponse.datetime.split(", ")[0] + ", " + userResponse.datetime.split(", ")[1]}</strong></p>
<p>Hora: <strong>${userResponse.datetime.split(", ")[2]}</strong></p>
<p>Tipo de sesión: <strong>Virtual</strong></p>
<p>Enlace de acceso: <strong>${MEET_URL}</strong></p>
`;

  const RESPONSES_SHEET = SpreadsheetApp.getActive().getSheetByName('respuestas');
  RESPONSES_SHEET.appendRow([
    userResponse.timestamp,
    userResponse.email,
    userResponse.service,
    FULL_TITLE,
    userResponse.fullname,
    userResponse.dni,
    userResponse.userType,
    userResponse.campus,
    userResponse.academicProgram,
    userResponse.studiesModality,
    userResponse.businessUnit,
    userResponse.phone,
    userResponse.dpp])

  // emailBody es para aquellos dispositivos que no pueden renderizar HTML, es texto plano.
  const emailBody = `Hola: ${userResponse.fullname}, ¡Confirmamos tu inscripción!
Cualquier consulta por favor no dudes en escribirnos a: capacitacioneshub@continental.edu.pe o contactarte a nuestros números telefónicos según el campus o sede más cercano.`;

  // emailTemplate obtiene la plantilla HTML con formato.
  const emailTemplate = HtmlService.createTemplateFromFile("emailTemplate");
  emailTemplate.fullname = userResponse.fullname;
  emailTemplate.description = DESCRIPTION;
  const htmlBody = emailTemplate.evaluate().getContent();

  const advancedOpts = {
    name: "Hub de Información | Apoyo a la Investigación",
    htmlBody: htmlBody,
    // cc: "capacitacioneshub@continental.edu.pe"
    cc: "fromeror@continental.edu.pe"
  };

  MailApp.sendEmail(userResponse.email, "✅ Confirmación de inscripción", emailBody, advancedOpts);
}

const hasScript = () => { SpreadsheetApp.getUi().alert("\nName: AI | Solicitud de capacitaciones 2025") }

const onOpen = () => { SpreadsheetApp.getUi().createMenu("🟢").addItem("Ver información del script", "hasScript").addToUi() }

