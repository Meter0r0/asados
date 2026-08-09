/**
 * Aplicación Web de Preventa Pública de Álbumes - ORC
 */

const SPREADSHEET_PV_ID = "1U_RVxTZMXZlHN1pbc-H0b020iDz4cvmMXdmkEWADIIY";
const FOLDER_COMPROBANTES_ID = "1ILXcVaUxMSzz16fxiFgnwek7T_jgifkK";
const FOLDER_FIGURITAS_ORIGEN_ID = "1e9ZbfZMweSS2kWP3cuN22rWgE0kUM5EE";
const FOLDER_FIGURITAS_DESTINO_ID = "17IdbCEHjXylvG7V_wsJvysRgTX7pwOB8";

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

/**
 * Valida que un número de figurita sea un entero entre 0 y 600,
 * o uno de los códigos especiales orc1..orc5.
 */
function esNumeroFiguritaValido(valor) {
  const v = String(valor || '').trim().toLowerCase();
  if (/^orc[1-5]$/.test(v)) return true;
  if (/^\d{1,3}$/.test(v)) {
    const n = parseInt(v, 10);
    return n >= 0 && n <= 600;
  }
  return false;
}

/**
 * Busca en la carpeta de origen un archivo cuyo nombre (sin extensión)
 * coincida con el número de figurita, y crea una copia en la carpeta de destino.
 * Devuelve la URL de la copia, o "" si no se encontró el archivo (best-effort, no bloqueante).
 */
function buscarYCopiarFigurita(numeroRaw, nombreJugadora, sourceFolder, destFolder, pedidoId) {
  try {
    const numero = String(numeroRaw || '').trim().toLowerCase();
    const files = sourceFolder.getFiles();
    let match = null;

    while (files.hasNext()) {
      const f = files.next();
      const name = f.getName();
      const dot = name.lastIndexOf('.');
      const baseName = (dot > -1 ? name.substring(0, dot) : name).trim().toLowerCase();
      if (baseName === numero) {
        match = f;
        break;
      }
    }

    if (!match) {
      Logger.log("FiguGigante: no se encontró archivo para el número '" + numeroRaw + "'");
      return "";
    }

    const originalName = match.getName();
    const dot = originalName.lastIndexOf('.');
    const ext = dot > -1 ? originalName.substring(dot) : "";
    const cleanNombre = String(nombreJugadora || '').replace(/[^a-zA-Z0-9]/g, "_");
    const newName = "FiguGigante_" + pedidoId + "_" + numeroRaw + "_" + cleanNombre + ext;

    const copy = match.makeCopy(newName, destFolder);
    return copy.getUrl();
  } catch (err) {
    Logger.log("Error copiando figurita '" + numeroRaw + "': " + err.toString());
    return "";
  }
}

/**
 * Registra un pedido de FiguGigante (una o más figuritas ampliadas)
 * en su propia pestaña de la planilla, y copia las fotos correspondientes
 * desde la carpeta de origen a la carpeta de destino para impresión.
 */
function guardarRegistroFiguGigante(nombre, categoria, telefono, figuritas, observaciones, fileObj) {
  try {
    if (!Array.isArray(figuritas) || figuritas.length === 0) {
      throw new Error("Debes agregar al menos una FiguGigante al pedido.");
    }
    figuritas.forEach(item => {
      if (!esNumeroFiguritaValido(item.numero)) {
        throw new Error("Número de figurita inválido: " + item.numero);
      }
      if (!item.nombreJugadora || !String(item.nombreJugadora).trim()) {
        throw new Error("Falta el nombre de la jugadora para la figurita " + item.numero);
      }
    });

    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    let sheet = ss.getSheetByName("FiguGigante");
    const headers = ["Timestamp", "ID Pedido", "Nombre Contacto", "Categoría", "Teléfono", "Número Figurita", "Nombre en la Foto", "Observaciones", "Comprobante", "Archivo Copiado", "Estado"];

    if (!sheet) {
      sheet = ss.insertSheet("FiguGigante");
      sheet.appendRow(headers);
      sheet.getRange(1, 1, 1, headers.length).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    } else {
      sheet.getRange(1, 1, 1, headers.length).setValues([headers]);
    }

    const timestamp = new Date();

    // Generar un ID único con formato FG-XXXX (4 dígitos aleatorios)
    let newId;
    let isUnique = false;
    const values = sheet.getDataRange().getValues();
    const existingIds = values.map(row => String(row[1]).trim());

    while (!isUnique) {
      const randomDigits = Math.floor(1000 + Math.random() * 9000);
      newId = "FG-" + randomDigits;
      if (existingIds.indexOf(newId) === -1) {
        isUnique = true;
      }
    }

    // Subir comprobante a Google Drive si fue adjuntado (opcional en este flujo)
    let fileUrl = "";
    if (fileObj && fileObj.data) {
      const comprobantesFolder = DriveApp.getFolderById(FOLDER_COMPROBANTES_ID);
      const cleanName = (nombre || "SinNombre").replace(/[^a-zA-Z0-9]/g, "_");
      const ext = fileObj.fileName.substring(fileObj.fileName.lastIndexOf('.'));
      const newFileName = "Comprobante_" + newId + "_" + cleanName + ext;
      const blob = Utilities.newBlob(Utilities.base64Decode(fileObj.data), fileObj.mimeType, newFileName);
      const comprobanteFile = comprobantesFolder.createFile(blob);
      fileUrl = comprobanteFile.getUrl();
    }

    const sourceFolder = DriveApp.getFolderById(FOLDER_FIGURITAS_ORIGEN_ID);
    const destFolder = DriveApp.getFolderById(FOLDER_FIGURITAS_DESTINO_ID);
    const estado = "Pendiente";

    figuritas.forEach(item => {
      const copiaUrl = buscarYCopiarFigurita(item.numero, item.nombreJugadora, sourceFolder, destFolder, newId);
      sheet.appendRow([
        timestamp,
        newId,
        nombre,
        categoria,
        telefono,
        String(item.numero).trim(),
        String(item.nombreJugadora).trim(),
        observaciones,
        fileUrl,
        copiaUrl,
        estado
      ]);
    });

    return { success: true, id: newId };
  } catch (e) {
    Logger.log("Error en guardarRegistroFiguGigante: " + e.toString());
    throw new Error("No se pudo registrar el pedido: " + e.message);
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
    let result;

    if (data.tipo === "figugigante") {
      result = guardarRegistroFiguGigante(
        data.nombre,
        data.categoria,
        data.telefono,
        data.figuritas,
        data.observaciones,
        data.fileObj
      );
    } else {
      result = guardarRegistroPreventa(
        data.nombre,
        data.categoria,
        data.telefono,
        data.cantidad,
        data.observaciones,
        data.fileObj
      );
    }

    return ContentService.createTextOutput(JSON.stringify(result))
      .setMimeType(ContentService.MimeType.JSON);
  } catch (err) {
    const errorResult = { success: false, error: err.toString() };
    return ContentService.createTextOutput(JSON.stringify(errorResult))
      .setMimeType(ContentService.MimeType.JSON);
  }
}
