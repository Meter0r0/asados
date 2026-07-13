/**
 * Aplicación Web de Preventa Pública de Álbumes - ORC
 */

const SPREADSHEET_PV_ID = "1U_RVxTZMXZlHN1pbc-H0b020iDz4cvmMXdmkEWADIIY";
const FOLDER_COMPROBANTES_ID = "1ILXcVaUxMSzz16fxiFgnwek7T_jgifkK";

function doGet(e) {
  // Intentar mover el script a la carpeta solicitada en Drive si no está allí
  try {
    const fileId = "1Vx7xTo9TH3R3wyAKqFXiuGiLxBRJJVtCJkXIv6j6Mt3ecM-lHgbjxHIU";
    const folderId = "18deU2PFXDCVsHKYoJs200Yui_lDIC7RY";
    const file = DriveApp.getFileById(fileId);
    const folder = DriveApp.getFolderById(folderId);
    
    const parents = file.getParents();
    let alreadyInFolder = false;
    const parentList = [];
    while (parents.hasNext()) {
      const p = parents.next();
      parentList.push(p);
      if (p.getId() === folderId) {
        alreadyInFolder = true;
      }
    }
    
    if (!alreadyInFolder) {
      folder.addFile(file);
      parentList.forEach(p => {
        p.removeFile(file);
      });
    }
  } catch (err) {
    Logger.log("Error al mover archivo: " + err.toString());
  }

  return HtmlService.createHtmlOutputFromFile('Preventa')
    .setTitle('Reserva de Álbumes - ORC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.DEFAULT)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Retorna la URL de publicación de este script (útil si se necesita internamente)
 */
function getWebAppUrl() {
  return ScriptApp.getService().getUrl();
}

/**
 * Registra un comensal/interesado en la preventa de álbumes.
 */
function guardarRegistroPreventa(nombre, categoria, telefono, cantidad, observaciones, fileObj) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    let sheet = ss.getSheetByName("Preventa");
    const headers = ["Timestamp", "ID", "Nombre de la Jugadora", "Categoría", "Teléfono de Contacto", "Cantidad de Álbumes", "Observaciones", "Comprobante de Transferencia", "Estado"];
    
    if (!sheet) {
      sheet = ss.insertSheet("Preventa");
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    } else {
      // Aseguramos que la primera fila siempre tenga las cabeceras deseadas
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const timestamp = new Date();
    
    // Generar un ID único con formato PR-XXXX (4 dígitos aleatorios)
    let newId;
    let isUnique = false;
    const values = sheet.getDataRange().getValues();
    const existingIds = values.map(row => String(row[1]).trim());
    
    while (!isUnique) {
      const randomDigits = Math.floor(1000 + Math.random() * 9000); // 1000 a 9999
      newId = "PR-" + randomDigits;
      if (existingIds.indexOf(newId) === -1) {
        isUnique = true;
      }
    }

    const estado = "Pendiente";

    // Subir comprobante a Google Drive
    let fileUrl = "";
    if (fileObj && fileObj.data) {
      const folder = DriveApp.getFolderById(FOLDER_COMPROBANTES_ID);
      const cleanName = nombre.replace(/[^a-zA-Z0-9]/g, "_");
      const ext = fileObj.fileName.substring(fileObj.fileName.lastIndexOf('.'));
      const newFileName = "Comprobante_" + newId + "_" + cleanName + ext;
      
      const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, newFileName);
      const file = folder.createFile(blob);
      fileUrl = file.getUrl();
    } else {
      throw new Error("El comprobante de transferencia es obligatorio.");
    }

    sheet.appendRow([
      timestamp,
      newId,
      nombre,
      categoria,
      telefono,
      Number(cantidad) || 1,
      observaciones,
      fileUrl,
      estado
    ]);

    return { success: true, id: newId };
  } catch (e) {
    Logger.log("Error en guardarRegistroPreventa: " + e.toString());
    throw new Error("No se pudo registrar la reserva: " + e.message);
  }
}

function testDrive() {
  const folder = DriveApp.getFolderById(FOLDER_COMPROBANTES_ID);
  const testFile = folder.createFile("test_conexion.txt", "Prueba de conexión exitosa");
  Logger.log("Archivo de prueba creado exitosamente: " + testFile.getUrl());
}

/**
 * Recibe peticiones HTTP POST para guardar registros de preventa
 * desde orígenes externos (como GitHub Pages) evitando problemas de CORS y sandbox.
 */
function doPost(e) {
  try {
    const data = JSON.parse(e.postData.contents);
    const result = guardarRegistroPreventa(
      data.nombre,
      data.categoria,
      data.telefono,
      data.cantidad,
      data.observaciones,
      data.fileObj
    );
    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const errorResult = { success: false, error: err.toString() };
    return ContentService.createTextOutput(JSON.stringify(errorResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
