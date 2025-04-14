function notificarOnFormSubmit(e) {
    var lock = LockService.getScriptLock();
    lock.waitLock(30000);  

    try {
        
        var hojaControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control");
        if (!hojaControl) throw new Error("La hoja 'Control' no existe.");
        
        var estadoProceso = hojaControl.getRange("A2").getValue();
        if (estadoProceso === "Procesando") {
            throw new Error("Otro proceso está en curso. Inténtalo nuevamente.");
        }

        hojaControl.getRange("A2").setValue("Procesando");
        SpreadsheetApp.flush();  

        var hojaOrigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas Transportes");
        if (!hojaOrigen) throw new Error("La hoja 'Respuestas Transportes' no existe.");

        var ultimaFila = hojaOrigen.getLastRow();
        if (ultimaFila < 1) throw new Error("No hay datos en la hoja 'Respuestas Transportes'.");

        var ticket = generarTicketUnico();
        copiarDatos(ultimaFila, ticket);

        var datos = hojaOrigen.getRange(ultimaFila, 1, 1, hojaOrigen.getLastColumn()).getValues()[0];
        var Email_Address = datos[1];  
        var NOMBRE_CONSULTORA = datos[3];  

        if (!Email_Address) throw new Error("La dirección de correo está vacía en la última fila.");
        if (!NOMBRE_CONSULTORA) NOMBRE_CONSULTORA = "Consultora";

        var ccEmailAddress = "anthonnymarcelooficial@gmail.com";
        var asunto = `Registro número ${ticket}`;
        var email_plano = `
            Hola ${NOMBRE_CONSULTORA}, te has registrado correctamente.\n\n
            Tu número de registro es: ${ticket}.\n\n
            Gracias por confiar en nosotros.\n\n
            Atentamente,\nEquipo Yobel.
        `;
        var email_html = `
            <p>Estimado <strong>${NOMBRE_CONSULTORA}</strong>, su solicitud ha sido registrada correctamente.</p>
            <p>Use el siguiente "<span style="color: red;"><strong>${ticket}</strong></span>" para su seguimiento:</p>
            <p>Responderemos en la brevedad de lo posible</p>
            <p>Gracias por confiar en el equipo Yobel.</p>
            <p><img src="https://media.licdn.com/dms/image/v2/C4E0BAQGNDseOJJX_XQ/company-logo_200_200/company-logo_200_200/0/1644592844922/yobel_scm_logo?e=2147483647&v=beta&t=HT3DZCAgqV941Akymdn3FQaZw9ybkuBH5Ayo302N4eA" height="100" /></p>
            <p>En caso este correo llegue a tu bandeja de entrada, fuera de tu jornada diaria laboral o desconexión digital, deberás revisarlo al día hábil siguiente.</p>
            <p>Este mensaje de correo electrónico contiene información estrictamente confidencial no susceptible de ser ditribuida. Si usted no es el destinatario de este mensaje, por favor no publicarlo, copiarlo o tomar cualquier otro tipo de acción sobre esta transmisión. Si recibió este mensaje por error, por favor notifíquenoslo y elimínelo lo antes posible.</p>
        `;
        var opcionesAvanzadas = { 
            name: "Yobel SCM", 
            htmlBody: email_html, 
            cc: ccEmailAddress 
        };

        MailApp.sendEmail(Email_Address, asunto, email_plano, opcionesAvanzadas);

    } catch (error) {
        Logger.log("Error en el proceso: " + error.message);
        throw error;

    } finally {
        if (hojaControl) {
            hojaControl.getRange("A2").setValue("");
            SpreadsheetApp.flush(); 
        }
        lock.releaseLock(); 
    }
}

function generarTicketUnico() {
    var hojaControl = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Control");
    if (!hojaControl) throw new Error("La hoja 'Control' no existe.");

    var celdaTicket = hojaControl.getRange("A1");
    var ultimoTicket = celdaTicket.getValue() || 0;
    var nuevoTicket = ultimoTicket + 1;

    celdaTicket.setValue(nuevoTicket);
    SpreadsheetApp.flush(); 

    return `Ticket${Utilities.formatString("%05d", nuevoTicket)}`;
}

function copiarDatos(filaOrigen, ticket) {
    var hojaDestino = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("YOBEL");
    var hojaOrigen = SpreadsheetApp.getActiveSpreadsheet().getSheetByName("Respuestas Transportes");

    if (!hojaDestino) throw new Error("La hoja 'YOBEL' no existe.");
    if (!hojaOrigen) throw new Error("La hoja 'Respuestas Transportes' no existe.");

    var ultimaFilaDestino = hojaDestino.getLastRow() + 1;
    var datos = hojaOrigen.getRange(filaOrigen, 1, 1, hojaOrigen.getLastColumn()).getValues();

    hojaDestino.getRange(ultimaFilaDestino, 1, 1, datos[0].length).setValues(datos);
    hojaDestino.getRange(ultimaFilaDestino, 12).setValue(ticket);
    hojaDestino.getRange(ultimaFilaDestino, 13).setValue("Pendiente");

    SpreadsheetApp.flush();

    var ticketGuardado = hojaDestino.getRange(ultimaFilaDestino, 12).getValue();
    if (ticketGuardado !== ticket) {
        throw new Error(`Error al guardar el ticket en la fila ${ultimaFilaDestino}.`);
    }
}
