/**
 * Find Bounced Emails - AI
 *
 * REFERENCIAS:
 * How to Generate a Report of Bounced Email Addresses in Gmail
 * https://www.labnol.org/internet/gmail-bounced-email-report/29209/
 * 
 * Guide to Function Calling with Gemini and Google Apps Script
 * https://medium.com/google-cloud/guide-to-function-calling-with-gemini-and-google-apps-script-0e058d472f45
 * 
 */

const API_KEY = '__INSERTE_AQUI_SU_LLAVE__';

function findBouncedEmails() {
  // Proceso
  let entry = '<strong>CORREO: ##1##</strong><br/>FECHA: ##0##<br/>ERROR: ##2##<br/><br/>##3##';  
  const { messages = [] } = Gmail.Users.Messages.list( 'me', {  q: 'from:mailer-daemon, newer_than:3d', maxResults: 15 });
  let ids = [];
  let response = '';
  let counter = 0;
  let emails = '';
  for ( let m=0; m<messages.length; m+= 1 ) {
    const bounceData = _parseMessage( messages[ m ].id );
    if ( bounceData ) {
      let res = entry;
      for ( let indx=0; indx<bounceData.length; indx++ ) res = res.replace( `##${indx}##`, bounceData[ indx ] );
      response += `${res}<br/><br/>`;
      ids.push( messages[ m ].id );
      emails += `${bounceData[ 1 ]}, `;  // Se acumula el correo
      counter++;
    };
  };//for
  if ( ids.length > 0 ) {
    // Prepara los datos para el correo de notificación
    let alert = {
      email: 'jorge.e.forero@gmail.com',
      subject: `Para Revisar: Tenemos rebote de correo`,
      text: `Por favor revisa los siguientes correos que presentan rebote:<br/><br/>${response} `,
    };
    // Envio el correo de notificaición
    GmailApp.sendEmail( alert.email, alert.subject, '', { htmlBody: alert.text} );
    // Aplica el label y archiva los emails identificados como rebotes
    applyGmailLabel( ids, 'Rebotes' );
    archiveBounces( 'Rebotes' );
  };
  // Registra el resultado de la operación
  console.log( `rebotes: ${counter} :: correos: ${emails}` ); 
};

/**
 * parseMessage
 * Obtiene la información de error del correo que falló con la traducción y el consejo a aplicar
 */
function _parseMessage( messageId ) {
  const message = GmailApp.getMessageById( messageId );
  const body = message.getPlainBody();
  const [, failAction] = body.match(/^Action:\s*(.+)/m) || [];
  // Si el envío fallo, se procesa
  if ( failAction === 'failed' ) {
    // El encabezadop del correo contiene la dirección de correo
    const email = message.getHeader( 'X-Failed-Recipients' );
    // Se obtiene el código de diagnostico
    const [, , bounceReason] = body.match(/^Diagnostic-Code:\s*(.+)\s*;\s*(.+)/m) || [];
    let reason = bounceReason.replace(/\s*(Please|Learn|See).+$/, '');
    let prompt = `I have received a bounced email with this error: ${reason}. Give two short advices to solve it and how to solve it`;
    // Llamado a Gemini
    let advice = _assistMeGemini( prompt );
    // Remueve * encontrados en la respuesta. Se reemplazan los saltos de línea por <br/> para el correo
    advice = advice.replace( /\**/g, '' );
    advice = advice.replace( /\n/g, '<br/>' );
    // Obtiene la traducción del error y del consejo
    let translatedReason = LanguageApp.translate( reason, 'en', 'es');
    let transatedAdvice = LanguageApp.translate( advice, 'en', 'es');
    return [ message.getDate(), email, translatedReason, transatedAdvice ];
  };
  return false;
};

/**
 * assistMeGemini
 * Obtenine una respuesta de Gemini a partir de el prompt dado
 * 
 * @param {string} Prompt - Texto de pregunta a resolver
 * @return {string} response - Texto de respuesta desde Gemini
 */
function _assistMeGemini( Prompt ) {
  // Url del servicio generate Content de Gemini
  const url = `https://generativelanguage.googleapis.com/v1beta/models/gemini-pro:generateContent?key=${API_KEY}`;
  const payload = { contents: [{ parts: [{ text: Prompt }] }] };
  const res = UrlFetchApp.fetch( url, { payload: JSON.stringify(payload), contentType: "application/json" });
  const obj = JSON.parse(res.getContentText());
  let response = 'No tenemos respuesta a este problema';
  if ( obj.candidates && obj.candidates.length > 0 && obj.candidates[0].content.parts.length > 0 ) {
     response = obj.candidates[0].content.parts[0].text;
  };
  return( response );
};

/**
 * applyGmailLabel
 * Aplica un Label a un conjunto de mensajes
 * 
 * @param {array} messageIds - arreglo con los ids de los mensajes a aplicar el label
 * @param {string} labelName - nombre del label a aplicar
 */
function applyGmailLabel( messageIds, labelName ) {
  const labelId = getGmailLabelId( labelName );
  Gmail.Users.Messages.batchModify( { addLabelIds: [ labelId ], ids: messageIds }, 'me' );
};

/**
 * archiveBounces
 * Una vez marcados los correos que han rebotado, se archivan.  Aplica para los rebotes que no hayan 
 * sido marcadados como leidos.
 * 
 * @param {string} Label - Nombre del label a archivar
 * @return {void} - mensajes archivados
 */
function archiveBounces( Label ) {
  const getUserLabelByName = GmailApp.getUserLabelByName( Label );
  const threads = getUserLabelByName.getThreads();
  const hasThreads = Array.isArray( threads ) && threads.length > 0;
  if ( hasThreads ) {
    threads.forEach(thread => {
    if ( thread.isUnread() ) {
        GmailApp.moveThreadToArchive( thread );
      };
    });
  };
};
