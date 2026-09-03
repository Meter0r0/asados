/**
 * @OnlyCurrentDoc
 * El siguiente código crea una aplicación web para registrar pagos
 * y los guarda en esta hoja de cálculo.
 */

// ID de la Planilla de Google ACTUALIZADO
// Extraído de: https://docs.google.com/spreadsheets/d/16yTXsGuJMwXpoYCBXdMx1HQAUT1cX1LYJ3fNjRiiBwg/
const SPREADSHEET_ID = "16yTXsGuJMwXpoYCBXdMx1HQAUT1cX1LYJ3fNjRiiBwg";
 
// Nombre de la hoja donde se guardarán los datos.
// IMPORTANTE: Asegúrate de que en la nueva planilla exista una pestaña con este nombre exacto.
const SHEET_NAME = "ListaComensales";

// Planilla y pestaña del Punto de Venta transaccional
const SPREADSHEET_PV_ID = "1U_RVxTZMXZlHN1pbc-H0b020iDz4cvmMXdmkEWADIIY";
const SHEET_PV_NAME = "VentasPuntoVenta";
 
/**
 * Entrega el archivo HTML para la interfaz de la aplicación web.
 * Esta función se ejecuta cuando un usuario visita la URL de la aplicación.
 */
function doGet(e) {
  const page = e && e.parameter && e.parameter.page;
  if (page === 'inspect') {
    const data = inspectPVSpreadsheet();
    return HtmlService.createHtmlOutput(JSON.stringify(data, null, 2));
  }
  if (page === 'dashboard') {
    return HtmlService.createHtmlOutputFromFile('Dashboard')
      .setTitle('Dashboard de Ventas - ORC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'ventas') {
    return HtmlService.createHtmlOutputFromFile('Ventas')
      .setTitle('Caja y Ventas de Figuritas - ORC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'puntoventa') {
    const template = HtmlService.createTemplateFromFile('PuntoVenta');
    template.vendedor = e && e.parameter && e.parameter.vendedor || '';
    template.editId = e && e.parameter && e.parameter.edit || '';
    return template.evaluate()
      .setTitle('Punto de Venta - ORC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'historial') {
    const template = HtmlService.createTemplateFromFile('Historial');
    template.vendedor = e && e.parameter && e.parameter.vendedor || '';
    return template.evaluate()
      .setTitle('Historial de Ventas - ORC')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  if (page === 'cobros') {
    return HtmlService.createHtmlOutputFromFile('Index')
      .setTitle('Registro de Cobros - Asado 18/6 - Club Olivos')
      .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
      .addMetaTag('viewport', 'width=device-width, initial-scale=1');
  }
  // Por defecto (o si page === 'asistencia')
  return HtmlService.createHtmlOutputFromFile('Asistencia')
    .setTitle('Asistencia Vendedores - ORC')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL)
    .addMetaTag('viewport', 'width=device-width, initial-scale=1');
}

/**
 * Procesa los datos enviados desde el formulario HTML.
 * Es llamada desde el lado del cliente usando google.script.run.
 * @param {Object} formObject - El objeto con los datos del formulario.
 * @return {String} Un mensaje de éxito.
 */
function processForm(formObject) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    const sheet = ss.getSheetByName(SHEET_NAME);
    
    // Verificación de seguridad por si la hoja no existe
    if (!sheet) {
      throw new Error(`No se encontró la pestaña "${SHEET_NAME}" en la planilla.`);
    }

    // Obtener los datos del objeto del formulario
    const cobrador = formObject.cobrador;
    const categoria = formObject.categoria;
    const nombre = formObject.nombre;
    const cantidad = formObject.cantidad;
    const tipo = formObject.tipo || 'cobro'; // 'cobro' o 'devolucion'
    const esDevolucion = (tipo === 'devolucion');
    const montoCalculado = esDevolucion ? -Math.abs(Number(formObject.montoCalculado)) : Number(formObject.montoCalculado);
    const montoCobrado = esDevolucion ? -Math.abs(Number(formObject.montoCobrado)) : Number(formObject.montoCobrado);
    const timestamp = new Date();

    // Generar recibo aleatorio de 4 caracteres (Alfanumérico)
    const caracteres = "ABCDEFGHIJKLMNOPQRSTUVWXYZ0123456789";
    let recibo = "";
    for (let i = 0; i < 4; i++) {
      recibo += caracteres.charAt(Math.floor(Math.random() * caracteres.length));
    }

    // Añadir una nueva fila a la planilla
    sheet.appendRow([
      timestamp,
      cobrador,
      categoria,
      nombre,
      cantidad,
      montoCalculado,
      montoCobrado,
      recibo,
      esDevolucion ? 'DEVOLUCION' : 'COBRO'
    ]);

    // Calcular quién lleva más cobros (sumando montoCobrado de filas tipo COBRO)
    let lider = null;
    try {
      const allData = sheet.getDataRange().getValues();
      const totales = {}; // { cobrador: totalCobrado }
      for (let i = 0; i < allData.length; i++) {
        const filaCobrador = String(allData[i][1]).trim();
        const filaMonto = Number(allData[i][6]) || 0;
        const filaTipo = String(allData[i][8]).trim().toUpperCase();
        if (filaCobrador && filaTipo === 'COBRO') {
          totales[filaCobrador] = (totales[filaCobrador] || 0) + filaMonto;
        }
      }
      let maxMonto = 0;
      let liderNombre = '';
      for (const c in totales) {
        if (totales[c] > maxMonto) {
          maxMonto = totales[c];
          liderNombre = c;
        }
      }
      if (liderNombre) {
        lider = { nombre: liderNombre, total: maxMonto };
      }
    } catch (err) {
      Logger.log("Error calculando líder: " + err.toString());
    }

    Logger.log("Datos guardados exitosamente: " + JSON.stringify({ ...formObject, recibo }));
    return {
      success: true,
      recibo: recibo,
      nombre: nombre,
      monto: montoCobrado,
      tipo: tipo,
      lider: lider
    };

  } catch (e) {
    Logger.log("Error al guardar los datos: " + e.toString());
    // Lanzar un nuevo error para que sea capturado por el handler de fallo en el cliente
    throw new Error("Error: " + e.message);
  }
}

/**
 * Obtiene las asignaciones guardadas en la solapa "Asistencia".
 * Si la hoja no existe, la crea con su fila de encabezado.
 * @return {Object} Objeto con las asignaciones actuales ("Fecha|Cancha": "Colaborador").
 */
/**
 * Normaliza cualquier valor de fecha leído desde el Sheet al formato estándar ISO "YYYY-MM-DD"
 * con salvaguardas para contrarrestar desfases por diferencias horarias.
 */
function normalizarFecha(fechaVal, ss) {
  if (!fechaVal) return "";
  
  if (fechaVal instanceof Date) {
    // Sumar 12 horas para evitar desfases horarios que cambien el día a la fecha anterior
    const fechaAjustada = new Date(fechaVal.getTime() + 12 * 60 * 60 * 1000);
    return Utilities.formatDate(fechaAjustada, ss.getSpreadsheetTimeZone() || "GMT-3", "yyyy-MM-dd");
  }
  
  let str = String(fechaVal).trim();
  if (/^\d{4}-\d{2}-\d{2}/.test(str)) {
    return str.substring(0, 10);
  }
  
  // Procesar formatos alternativos de texto (ej: DD/MM/YYYY)
  const partes = str.split('/');
  if (partes.length === 3) {
    const dia = parseInt(partes[0], 10);
    const mes = parseInt(partes[1], 10);
    const anio = parseInt(partes[2], 10);
    if (!isNaN(dia) && !isNaN(mes) && !isNaN(anio)) {
      const pad = (n) => String(n).padStart(2, '0');
      const anioCompleto = anio < 100 ? 2000 + anio : anio;
      return `${anioCompleto}-${pad(mes)}-${pad(dia)}`;
    }
  }
  
  const parsed = Date.parse(str);
  if (!isNaN(parsed)) {
    const dateObj = new Date(parsed + 12 * 60 * 60 * 1000);
    return Utilities.formatDate(dateObj, ss.getSpreadsheetTimeZone() || "GMT-3", "yyyy-MM-dd");
  }
  
  return str;
}

function getAsistenciaData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Asistencia");
    if (!sheet) {
      sheet = ss.insertSheet("Asistencia");
      sheet.appendRow(["Timestamp", "Fecha", "Cancha", "Colaborador", "Usuario"]);
      sheet.getRange(1, 1, 1, 5).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    const values = sheet.getDataRange().getValues();
    const assignments = {};
    
    // Saltamos la cabecera (i = 1)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const fecha = normalizarFecha(row[1], ss);
      const cancha = String(row[2]).trim();
      const colaborador = String(row[3]).trim();
      if (fecha && cancha) {
        assignments[`${fecha}|${cancha}`] = colaborador;
      }
    }
    return assignments;
  } catch (e) {
    Logger.log("Error en getAsistenciaData: " + e.toString());
    throw new Error("No se pudieron cargar los datos de asistencia: " + e.message);
  }
}

/**
 * Registra una asignación de asistencia agregando una nueva fila.
 * Se evita el uso de Session.getActiveUser().getEmail() para no requerir login del usuario.
 * @param {String} fecha - La fecha en formato "YYYY-MM-DD"
 * @param {String} cancha - El punto de venta ("Cancha 1" o "Cancha 2")
 * @param {String} colaborador - El nombre del colaborador ("Joaco", "Fusche", etc. o vacío)
 * @return {Object} Resultado del guardado.
 */
function guardarAsistencia(fecha, cancha, colaborador) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Asistencia");
    if (!sheet) {
      sheet = ss.insertSheet("Asistencia");
      sheet.appendRow(["Timestamp", "Fecha", "Cancha", "Colaborador", "Usuario"]);
      sheet.getRange(1, 1, 1, 5).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    const timestamp = new Date();
    const usuario = "Usuario Web"; // Fijo para evitar autenticación de Google en clientes
    
    sheet.appendRow([
      timestamp,
      fecha,
      cancha,
      colaborador,
      usuario
    ]);
    
    return { success: true };
  } catch (e) {
    Logger.log("Error en guardarAsistencia: " + e.toString());
    throw new Error("No se pudo guardar la asignación: " + e.message);
  }
}

/**
 * Obtiene la URL de publicación de la Web App para que los clientes puedan navegar.
 * @return {String} La URL del servicio web.
 */
function getWebAppUrl() {
  try {
    return ScriptApp.getService().getUrl();
  } catch (e) {
    Logger.log("Error en getWebAppUrl: " + e.toString());
    return "";
  }
}

/**
 * Obtiene los registros de ventas guardados en la solapa "Ventas".
 * Si la hoja no existe, la crea con su fila de encabezado.
 * @return {Object} Objeto con los datos de ventas indexados por "Fecha|Cancha".
 */
function getVentasData() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Ventas");
    if (!sheet) {
      sheet = ss.insertSheet("Ventas");
      sheet.appendRow(["Fecha", "Cancha", "Registrador", "Paquetes", "Efectivo", "CBU", "Timestamp", "ID"]);
      sheet.getRange(1, 1, 1, 8).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    // Migración automática para agregar la columna "ID" si no existe
    let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    let idColIndex = headers.indexOf("ID");
    if (idColIndex === -1) {
      sheet.getRange(1, 8).setValue("ID").setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const ids = [];
        for (let r = 2; r <= lastRow; r++) {
          ids.push(["ID-" + new Date().getTime() + "-" + r + "-" + Math.floor(Math.random() * 1000)]);
        }
        sheet.getRange(2, 8, ids.length, 1).setValues(ids);
      }
      idColIndex = 7; // Columna 8 (0-indexed)
    }

    const values = sheet.getDataRange().getValues();
    const ventas = {}; // Map: key -> Array of records
    
    // Saltamos la cabecera (i = 1)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const fecha = normalizarFecha(row[0], ss);
      const cancha = String(row[1]).trim();
      const registrador = String(row[2]).trim();
      const paquetes = Number(row[3]) || 0;
      const efectivo = Number(row[4]) || 0;
      const cbu = Number(row[5]) || 0;
      const id = row[7] ? String(row[7]).trim() : ("ID-GEN-" + i + "-" + Math.floor(Math.random() * 1000));
      
      if (fecha && cancha) {
        const key = `${fecha}|${cancha}`;
        if (!ventas[key]) {
          ventas[key] = [];
        }
        ventas[key].push({
          id: id,
          registrador: registrador,
          paquetes: paquetes,
          efectivo: efectivo,
          cbu: cbu
        });
      }
    }
    return ventas;
  } catch (e) {
    Logger.log("Error en getVentasData: " + e.toString());
    throw new Error("No se pudieron cargar los datos de ventas: " + e.message);
  }
}

/**
 * Guarda o edita un registro de venta en la solapa "Ventas".
 * Si se pasa un recordId válido, edita ese registro. Si no, agrega un nuevo registro.
 * @param {String} recordId - El ID único del registro de venta (opcional)
 * @param {String} fecha - La fecha del turno ("YYYY-MM-DD")
 * @param {String} cancha - El punto de venta ("Cancha 1" o "Cancha 2")
 * @param {String} registrador - Quién registra el cobro ("Joaco", "Fusche", etc.)
 * @param {Number} paquetes - Cantidad de paquetes vendidos
 * @param {Number} efectivo - Monto recaudado en efectivo (ARS)
 * @param {Number} cbu - Monto recaudado en transferencia (ARS)
 * @return {Object} Objeto indicando el éxito de la operación.
 */
function guardarVenta(recordId, fecha, cancha, registrador, paquetes, efectivo, cbu) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_ID);
    let sheet = ss.getSheetByName("Ventas");
    if (!sheet) {
      sheet = ss.insertSheet("Ventas");
      sheet.appendRow(["Fecha", "Cancha", "Registrador", "Paquetes", "Efectivo", "CBU", "Timestamp", "ID"]);
      sheet.getRange(1, 1, 1, 8).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    }
    
    // Buscar la columna de ID
    let headers = sheet.getRange(1, 1, 1, Math.max(sheet.getLastColumn(), 1)).getValues()[0];
    let idColIndex = headers.indexOf("ID");
    if (idColIndex === -1) {
      sheet.getRange(1, 8).setValue("ID").setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
      idColIndex = 7;
    }

    const values = sheet.getDataRange().getValues();
    let rowIndex = -1;
    
    // Si nos pasaron un ID, buscamos la fila por ID
    if (recordId) {
      for (let i = 1; i < values.length; i++) {
        const rowId = values[i][idColIndex] ? String(values[i][idColIndex]).trim() : "";
        if (rowId === recordId) {
          rowIndex = i + 1; // 1-indexed
          break;
        }
      }
    }
    
    const timestamp = new Date();
    
    if (rowIndex !== -1) {
      // Editar registro existente (columnas: Fecha=1, Cancha=2, Registrador=3, Paquetes=4, Efectivo=5, CBU=6, Timestamp=7)
      sheet.getRange(rowIndex, 1, 1, 7).setValues([[
        fecha,
        cancha,
        registrador,
        Number(paquetes) || 0,
        Number(efectivo) || 0,
        Number(cbu) || 0,
        timestamp
      ]]);
    } else {
      // Generar nuevo ID único
      const newId = "ID-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
      // Agregar nueva fila
      sheet.appendRow([
        fecha,
        cancha,
        registrador,
        Number(paquetes) || 0,
        Number(efectivo) || 0,
        Number(cbu) || 0,
        timestamp,
        newId
      ]);
    }
    
    return { success: true };
  } catch (e) {
    Logger.log("Error en guardarVenta: " + e.toString());
    throw new Error("No se pudo registrar la venta: " + e.message);
  }
}

/**
 * Retorna metadatos del script actual para depuración en el cliente
 */
function getScriptInfo() {
  return {
    scriptId: ScriptApp.getScriptId(),
    version: "Build-20260706-0143",
    timeZone: Session.getScriptTimeZone()
  };
}

/**
 * Verifica si existe la columna "Albums30k" en la planilla del Punto de Venta
 * y realiza la migración insertando la columna después de "Albums".
 */
function verificarYMigrarColumnas(sheet) {
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return;

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0];
  const albums30kIndex = headers.indexOf("Albums30k");
  if (albums30kIndex === -1) {
    // Buscar la columna "Albums"
    let albumsIndex = headers.indexOf("Albums");
    if (albumsIndex === -1) {
      albumsIndex = headers.indexOf("Albums25k");
    }

    if (albumsIndex !== -1) {
      const insertCol = albumsIndex + 2; // Insertar después de la columna "Albums" (1-indexed + 1)
      sheet.insertColumnAfter(albumsIndex + 1);
      sheet.getRange(1, insertCol).setValue("Albums30k")
           .setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");

      // Completar filas existentes con 0
      const lastRow = sheet.getLastRow();
      if (lastRow > 1) {
        const fillValues = [];
        for (let r = 2; r <= lastRow; r++) {
          fillValues.push([0]);
        }
        sheet.getRange(2, insertCol, fillValues.length, 1).setValues(fillValues);
      }
    }
  }

  // Verificar/agregar columna "FiguGigante" al final de la planilla
  const headersActuales = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headersActuales.indexOf("FiguGigante") === -1) {
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue("FiguGigante")
         .setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const fillValues = [];
      for (let r = 2; r <= lastRow; r++) {
        fillValues.push([0]);
      }
      sheet.getRange(2, newCol, fillValues.length, 1).setValues(fillValues);
    }
  }

  // Verificar/agregar columna "Promo2x1Paq10" al final de la planilla
  const headersActuales2 = sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];
  if (headersActuales2.indexOf("Promo2x1Paq10") === -1) {
    const newCol = sheet.getLastColumn() + 1;
    sheet.getRange(1, newCol).setValue("Promo2x1Paq10")
         .setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");

    const lastRow = sheet.getLastRow();
    if (lastRow > 1) {
      const fillValues = [];
      for (let r = 2; r <= lastRow; r++) {
        fillValues.push([0]);
      }
      sheet.getRange(2, newCol, fillValues.length, 1).setValues(fillValues);
    }
  }
}

/**
 * Guarda o edita un registro de venta en la planilla del punto de venta.
 */
function guardarVentaPuntoVenta(recordId, vendedor, comprador, albums, albums30k, paq1, paq4, paq10, figuGigante, total, efectivo, cbu, contado, promo2x1) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    let sheet = ss.getSheetByName(SHEET_PV_NAME);
    if (!sheet) {
      sheet = ss.insertSheet(SHEET_PV_NAME);
      sheet.appendRow(["Timestamp", "ID", "Vendedor", "Comprador", "Albums", "Albums30k", "Paquetes1", "Paquetes4", "Paquetes10", "MontoTotal", "CobradoEfectivo", "CobradoTransferencia", "MontoContadoEfectivo", "DiferenciaEfectivo", "FiguGigante", "Promo2x1Paq10"]);
      sheet.getRange(1, 1, 1, 16).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    }

    // Asegurar la presencia de la columna Albums30k
    verificarYMigrarColumnas(sheet);

    const timestamp = new Date();
    const values = sheet.getDataRange().getValues();
    let rowIndex = -1;

    if (recordId) {
      for (let i = 1; i < values.length; i++) {
        const rowId = values[i][1] ? String(values[i][1]).trim() : "";
        if (rowId === recordId) {
          rowIndex = i + 1; // 1-indexed
          break;
        }
      }
    }

    const numContado = Number(contado) || 0;
    const numEfectivo = Number(efectivo) || 0;
    const diff = numContado - numEfectivo;

    if (rowIndex !== -1) {
      // Editar registro existente (columnas: Timestamp=1..DiferenciaEfectivo=14, FiguGigante=15, Promo2x1Paq10=16)
      sheet.getRange(rowIndex, 1, 1, 16).setValues([[
        timestamp,
        recordId,
        vendedor,
        comprador,
        Number(albums) || 0,
        Number(albums30k) || 0,
        Number(paq1) || 0,
        Number(paq4) || 0,
        Number(paq10) || 0,
        Number(total) || 0,
        numEfectivo,
        Number(cbu) || 0,
        numContado,
        diff,
        Number(figuGigante) || 0,
        Number(promo2x1) || 0
      ]]);
    } else {
      // Generar nuevo ID único
      const newId = "PV-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
      sheet.appendRow([
        timestamp,
        newId,
        vendedor,
        comprador,
        Number(albums) || 0,
        Number(albums30k) || 0,
        Number(paq1) || 0,
        Number(paq4) || 0,
        Number(paq10) || 0,
        Number(total) || 0,
        numEfectivo,
        Number(cbu) || 0,
        numContado,
        diff,
        Number(figuGigante) || 0,
        Number(promo2x1) || 0
      ]);
    }

    return { success: true };
  } catch (e) {
    Logger.log("Error en guardarVentaPuntoVenta: " + e.toString());
    throw new Error("No se pudo registrar la venta: " + e.message);
  }
}

/**
 * Obtiene el histórico de ventas para un vendedor en específico.
 */
function getVentasPuntoVenta(vendedor) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    const sheet = ss.getSheetByName(SHEET_PV_NAME);
    if (!sheet) {
      return [];
    }

    // Asegurar migración de columnas
    verificarYMigrarColumnas(sheet);

    const values = sheet.getDataRange().getValues();
    const ventas = [];

    // Saltamos la cabecera (i = 1)
    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowVendedor = String(row[2]).trim();
      
      // Filtrar por el vendedor seleccionado
      if (rowVendedor === vendedor) {
        ventas.push({
          timestamp: row[0] ? (row[0] instanceof Date ? row[0].toISOString() : String(row[0])) : "",
          id: String(row[1]).trim(),
          vendedor: rowVendedor,
          comprador: String(row[3]).trim(),
          albums: Number(row[4]) || 0,
          albums30k: Number(row[5]) || 0,
          paq1: Number(row[6]) || 0,
          paq4: Number(row[7]) || 0,
          paq10: Number(row[8]) || 0,
          total: Number(row[9]) || 0,
          efectivo: Number(row[10]) || 0,
          cbu: Number(row[11]) || 0,
          contado: Number(row[12]) || 0,
          diferencia: Number(row[13]) || 0,
          figuGigante: Number(row[14]) || 0,
          promo2x1: Number(row[15]) || 0
        });
      }
    }

    // Ordenar por fecha descendente (las más recientes primero)
    ventas.sort((a, b) => new Date(b.timestamp) - new Date(a.timestamp));
    
    return ventas;
  } catch (e) {
    Logger.log("Error en getVentasPuntoVenta: " + e.toString());
    throw new Error("No se pudo cargar el historial de ventas: " + e.message);
  }
}

/**
 * Elimina un registro de venta en la planilla del punto de venta.
 */
function eliminarVentaPuntoVenta(recordId) {
  try {
    if (!recordId) throw new Error("ID de registro no provisto.");
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    const sheet = ss.getSheetByName(SHEET_PV_NAME);
    if (!sheet) throw new Error("La planilla de ventas no existe.");

    const values = sheet.getDataRange().getValues();
    let rowIndex = -1;

    for (let i = 1; i < values.length; i++) {
      const rowId = values[i][1] ? String(values[i][1]).trim() : "";
      if (rowId === recordId) {
        rowIndex = i + 1; // 1-indexed
        break;
      }
    }

    if (rowIndex !== -1) {
      sheet.deleteRow(rowIndex);
      return { success: true };
    } else {
      throw new Error("No se encontró el registro con ID: " + recordId);
    }
  } catch (e) {
    Logger.log("Error en eliminarVentaPuntoVenta: " + e.toString());
    throw new Error("No se pudo eliminar la venta: " + e.message);
  }
}

/**
 * Retorna información de semana (clave, etiqueta y timestamp del lunes) para una fecha.
 */
function getWeekInfo(d) {
  if (!(d instanceof Date) || isNaN(d.getTime())) return null;
  const date = new Date(d.getTime());
  date.setHours(0, 0, 0, 0);
  const day = date.getDay();
  const diff = date.getDate() - day + (day === 0 ? -6 : 1); // lunes
  const monday = new Date(date.setDate(diff));
  
  const mondayStr = monday.toLocaleDateString('es-AR', { day: '2-digit', month: '2-digit' });
  const tempDate = new Date(monday.valueOf());
  tempDate.setDate(tempDate.getDate() + 4 - (tempDate.getDay() || 7));
  const yearStart = new Date(tempDate.getFullYear(), 0, 1);
  const weekNo = Math.ceil((((tempDate - yearStart) / 86400000) + 1) / 7);
  
  return {
    key: `${monday.getFullYear()}-W${weekNo < 10 ? '0' + weekNo : weekNo}`,
    label: `Sem ${weekNo} (${mondayStr})`,
    mondayTimestamp: monday.getTime()
  };
}

/**
 * Obtiene los datos acumulados para el Dashboard
 */
function getDashboardData() {
  try {
    if (!SPREADSHEET_PV_ID) {
      return { error: "El ID de la planilla de Punto de Venta (SPREADSHEET_PV_ID) no está configurado." };
    }
    
    let ss;
    try {
      ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    } catch (openErr) {
      return { error: "No se pudo abrir la planilla de Google Sheets (ID: " + SPREADSHEET_PV_ID + "). Verifique los permisos de acceso. Detalle: " + openErr.message };
    }

    let sheet = ss.getSheetByName(SHEET_PV_NAME);
    
    const stats = {
      productos: {
        albums: { qty: 0, revenue: 0 },
        albums30k: { qty: 0, revenue: 0 },
        paq1: { qty: 0, revenue: 0 },
        paq4: { qty: 0, revenue: 0 },
        paq10: { qty: 0, revenue: 0 },
        figuGigante: { qty: 0, revenue: 0 },
        promo2x1: { qty: 0, revenue: 0 }
      },
      vendedores: {},
      caja: {
        efectivo: 0,
        cbu: 0,
        cobradoTotal: 0,
        ventasTotal: 0,
        sinClasificar: 0,
        total: 0
      },
      preventaQty: 0,
      ultimasVentas: []
    };

    // Inicializar vendedores oficiales
    const vendedoresOficiales = ["Joaco", "Fusche", "MartinC", "MartinD", "Fortu", "Garu", "Facu", "Bar"];
    vendedoresOficiales.forEach(v => {
      stats.vendedores[v] = { efectivo: 0, cbu: 0, total: 0, sinClasificar: 0, albums: 0, albums30k: 0, paquetes: 0 };
    });

    if (!sheet) {
      return stats;
    }

    // Asegurar migración de columnas si la hoja es editable
    try {
      verificarYMigrarColumnas(sheet);
    } catch (migErr) {
      Logger.log("Advertencia migración columnas: " + migErr.message);
    }

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) {
      return stats;
    }

    // Mapeo dinámico de columnas según encabezados (normalizado sin espacios ni caracteres especiales)
    const headerRowClean = values[0].map(h => String(h || '').trim().toLowerCase().replace(/[^a-z0-9]/gi, ''));
    
    const findIndexByAliases = (aliases, fallback) => {
      for (const alias of aliases) {
        const cleanAlias = alias.toLowerCase().replace(/[^a-z0-9]/gi, '');
        const idx = headerRowClean.indexOf(cleanAlias);
        if (idx !== -1) return idx;
      }
      return fallback;
    };

    const colIndex = {
      timestamp: findIndexByAliases(["timestamp", "fecha"], 0),
      id: findIndexByAliases(["id"], 1),
      vendedor: findIndexByAliases(["vendedor", "registrador"], 2),
      comprador: findIndexByAliases(["comprador", "cliente"], 3),
      albums: findIndexByAliases(["albums", "albums25k", "albumes", "libros"], 4),
      albums30k: findIndexByAliases(["albums30k", "albumes30k", "libros30k"], 5),
      paq1: findIndexByAliases(["paquetes1", "paq1", "paquete1"], 6),
      paq4: findIndexByAliases(["paquetes4", "paq4", "paquete4"], 7),
      paq10: findIndexByAliases(["paquetes10", "paq10", "paquete10"], 8),
      total: findIndexByAliases(["montototal", "total", "monto", "totalrecaudado"], 9),
      efectivo: findIndexByAliases(["cobradoefectivo", "efectivo", "montoefectivo", "cobradoenefectivo"], 10),
      cbu: findIndexByAliases(["cobradotransferencia", "cobradocbu", "cbu", "transferencia"], 11),
      contado: findIndexByAliases(["montocontadoefectivo", "contado", "montocontado"], 12),
      diferencia: findIndexByAliases(["diferenciaefectivo", "diferencia"], 13),
      figuGigante: findIndexByAliases(["figugigante", "figugigantes"], 14),
      promo2x1: findIndexByAliases(["promo2x1paq10", "promo2x1", "2x1superpack10", "2x1paq10"], 15)
    };

    const getValue = (row, idx) => (idx >= 0 && idx < row.length ? Number(row[idx]) || 0 : 0);

    const semanasMap = {};

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      if (!row || row.length <= colIndex.vendedor) continue;

      const vendedorRaw = row[colIndex.vendedor] != null ? String(row[colIndex.vendedor]).trim() : "";
      if (!vendedorRaw) continue;

      let vendedor = vendedoresOficiales.find(v => v.toLowerCase() === vendedorRaw.toLowerCase());
      if (!vendedor) {
        vendedor = vendedorRaw;
        if (!stats.vendedores[vendedor]) {
          stats.vendedores[vendedor] = { efectivo: 0, cbu: 0, total: 0, sinClasificar: 0, albums: 0, albums30k: 0, paquetes: 0 };
        }
      }

      const albums = getValue(row, colIndex.albums);
      const albums30k = getValue(row, colIndex.albums30k);
      const paq1 = getValue(row, colIndex.paq1);
      const paq4 = getValue(row, colIndex.paq4);
      const paq10 = getValue(row, colIndex.paq10);
      const figuGigante = getValue(row, colIndex.figuGigante);
      const promo2x1 = getValue(row, colIndex.promo2x1);
      const rawTotal = getValue(row, colIndex.total);
      const efectivo = getValue(row, colIndex.efectivo);
      const cbu = getValue(row, colIndex.cbu);

      const prodTotal = (albums * 25000 + albums30k * 30000 + paq1 * 3000 + paq4 * 10000 + paq10 * 20000 + figuGigante * 10000 + promo2x1 * 20000);
      const cobradoMetodos = efectivo + cbu;
      const totalVenta = rawTotal > 0 ? rawTotal : (cobradoMetodos > 0 ? cobradoMetodos : prodTotal);
      const sinClasificarFila = cobradoMetodos > 0 ? Math.max(0, totalVenta - cobradoMetodos) : totalVenta;
      const paquetesFila = paq1 + (paq4 * 4) + (paq10 * 10) + (promo2x1 * 10);

      // Acumular cantidades e ingresos
      stats.productos.albums.qty += albums;
      stats.productos.albums.revenue += albums * 25000;

      stats.productos.albums30k.qty += albums30k;
      stats.productos.albums30k.revenue += albums30k * 30000;

      stats.productos.paq1.qty += paq1;
      stats.productos.paq1.revenue += paq1 * 3000;

      stats.productos.paq4.qty += paq4;
      stats.productos.paq4.revenue += paq4 * 10000;

      stats.productos.paq10.qty += paq10;
      stats.productos.paq10.revenue += paq10 * 20000;

      stats.productos.figuGigante.qty += figuGigante;
      stats.productos.figuGigante.revenue += figuGigante * 10000;

      stats.productos.promo2x1.qty += promo2x1;
      stats.productos.promo2x1.revenue += promo2x1 * 20000;

      // Acumular por vendedor
      stats.vendedores[vendedor].efectivo += efectivo;
      stats.vendedores[vendedor].cbu += cbu;
      stats.vendedores[vendedor].total += totalVenta;
      stats.vendedores[vendedor].sinClasificar += sinClasificarFila;
      stats.vendedores[vendedor].albums += albums;
      stats.vendedores[vendedor].albums30k += albums30k;
      stats.vendedores[vendedor].paquetes += paquetesFila;

      // Acumular total global
      stats.caja.efectivo += efectivo;
      stats.caja.cbu += cbu;
      stats.caja.cobradoTotal += cobradoMetodos;
      stats.caja.ventasTotal += totalVenta;
      stats.caja.sinClasificar += sinClasificarFila;
      stats.caja.total += totalVenta;

      // Acumular ventas por semana
      const rawTimestamp = row[colIndex.timestamp];
      let dateObj = null;
      if (rawTimestamp instanceof Date && !isNaN(rawTimestamp.getTime())) {
        dateObj = rawTimestamp;
      } else if (rawTimestamp != null && rawTimestamp !== "") {
        const parsed = new Date(rawTimestamp);
        if (!isNaN(parsed.getTime())) dateObj = parsed;
      }

      if (dateObj) {
        const weekInfo = getWeekInfo(dateObj);
        if (weekInfo) {
          if (!semanasMap[weekInfo.key]) {
            semanasMap[weekInfo.key] = {
              key: weekInfo.key,
              label: weekInfo.label,
              mondayTimestamp: weekInfo.mondayTimestamp,
              paquetes: 0,
              revenue: 0
            };
          }
          semanasMap[weekInfo.key].paquetes += paquetesFila;
          semanasMap[weekInfo.key].revenue += totalVenta;
        }
      }
    }

    // Ordenar y calcular acumulado semanal
    let acumuladoPacks = 0;
    const ventasSemanales = Object.values(semanasMap)
      .sort((a, b) => a.mondayTimestamp - b.mondayTimestamp)
      .map(w => {
        acumuladoPacks += w.paquetes;
        return {
          key: w.key,
          label: w.label,
          paquetes: w.paquetes,
          revenue: w.revenue,
          acumulado: acumuladoPacks
        };
      });

    // Calcular proyecciones
    const totalPacksVendidos = stats.productos.paq1.qty + (stats.productos.paq4.qty * 4) + (stats.productos.paq10.qty * 10) + (stats.productos.promo2x1.qty * 10);
    const targetPacks = 37000;
    const remainingPacks = Math.max(0, targetPacks - totalPacksVendidos);

    let promedioSemanal = 0;
    if (ventasSemanales.length > 0) {
      const recentWeeks = ventasSemanales.slice(-4);
      const sumRecent = recentWeeks.reduce((acc, w) => acc + w.paquetes, 0);
      promedioSemanal = Math.round(sumRecent / Math.max(1, recentWeeks.length));
      if (promedioSemanal === 0) {
        promedioSemanal = Math.round(totalPacksVendidos / Math.max(1, ventasSemanales.length));
      }
    }

    const semanasRestantes = (promedioSemanal > 0 && remainingPacks > 0) ? Math.ceil(remainingPacks / promedioSemanal) : 0;
    
    let fechaEstimadaStr = "N/D";
    if (remainingPacks === 0) {
      fechaEstimadaStr = "🎯 ¡Objetivo Alcanzado!";
    } else if (semanasRestantes > 0) {
      const hoy = new Date();
      const fechaFin = new Date(hoy.getTime() + (semanasRestantes * 7 * 24 * 60 * 60 * 1000));
      fechaEstimadaStr = fechaFin.toLocaleDateString('es-AR', { day: '2-digit', month: 'long', year: 'numeric' });
    } else {
      fechaEstimadaStr = "Sin ritmo registrado";
    }

    stats.ventasSemanales = ventasSemanales;
    stats.proyeccion = {
      promedioSemanal: promedioSemanal,
      semanasRestantes: semanasRestantes,
      fechaEstimada: fechaEstimadaStr,
      remainingPacks: remainingPacks,
      totalSemanasActivas: ventasSemanales.length
    };

    // Agregar últimas ventas
    const limit = Math.max(1, values.length - 15);
    for (let i = values.length - 1; i >= limit; i--) {
      const row = values[i];
      if (i === 0 || !row) break;
      const timestamp = row[colIndex.timestamp];
      const recordId = row[colIndex.id];
      const vendedor = row[colIndex.vendedor];
      const comprador = row[colIndex.comprador];

      let dateString = "";
      try {
        if (timestamp instanceof Date && !isNaN(timestamp.getTime())) {
          dateString = timestamp.toISOString();
        } else if (timestamp != null) {
          dateString = String(timestamp);
        }
      } catch (dateErr) {
        dateString = String(timestamp || '');
      }

      stats.ultimasVentas.push({
        id: String(recordId || ''),
        timestamp: dateString,
        vendedor: String(vendedor || ''),
        comprador: String(comprador || ''),
        albums: getValue(row, colIndex.albums),
        albums30k: getValue(row, colIndex.albums30k),
        paq1: getValue(row, colIndex.paq1),
        paq4: getValue(row, colIndex.paq4),
        paq10: getValue(row, colIndex.paq10),
        figuGigante: getValue(row, colIndex.figuGigante),
        promo2x1: getValue(row, colIndex.promo2x1),
        total: getValue(row, colIndex.total),
        efectivo: getValue(row, colIndex.efectivo),
        cbu: getValue(row, colIndex.cbu),
        contado: getValue(row, colIndex.contado),
        diferencia: getValue(row, colIndex.diferencia)
      });
    }

    // Obtener la cantidad de álbumes pendientes de entrega en la pestaña "Preventa" (filtrando por Col I: Estado = "Pendiente")
    let preventaQty = 0;
    try {
      const preventaSheet = ss.getSheetByName("Preventa");
      if (preventaSheet) {
        const preventaValues = preventaSheet.getDataRange().getValues();
        if (preventaValues.length > 1) {
          const headers = preventaValues[0].map(h => String(h || '').trim().toLowerCase());
          let qtyColIndex = headers.indexOf("cantidad de álbumes");
          if (qtyColIndex === -1) qtyColIndex = headers.indexOf("cantidad de albumes");
          if (qtyColIndex === -1) qtyColIndex = 5;

          let estadoColIndex = headers.indexOf("estado");
          if (estadoColIndex === -1) estadoColIndex = 8;

          for (let i = 1; i < preventaValues.length; i++) {
            const row = preventaValues[i];
            const estadoVal = String(row[estadoColIndex] || '').trim().toLowerCase();
            // Sumar únicamente si el estado es "pendiente" (o si aún no se definió estado)
            if (estadoVal === "pendiente" || estadoVal === "") {
              preventaQty += Number(row[qtyColIndex]) || 0;
            }
          }
        }
      }
    } catch (prevErr) {
      Logger.log("Advertencia preventa: " + prevErr.message);
    }
    stats.preventaQty = preventaQty;

    // Garantizar que todos los valores sean serializables limpiamente en Apps Script
    return JSON.parse(JSON.stringify(stats));
  } catch (e) {
    Logger.log("Error en getDashboardData: " + e.toString());
    return { error: "Error en el servidor al obtener datos del dashboard: " + e.message };
  }
}

/**
 * Obtiene un único registro de venta por su ID.
 */
function getVentaPorId(recordId) {
  try {
    if (!recordId) throw new Error("ID de registro no provisto.");
    
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    const sheet = ss.getSheetByName(SHEET_PV_NAME);
    if (!sheet) throw new Error("La planilla de ventas no existe.");

    // Asegurar migración de columnas
    verificarYMigrarColumnas(sheet);

    const values = sheet.getDataRange().getValues();

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const rowId = String(row[1]).trim();
      if (rowId === recordId) {
        return {
          timestamp: row[0] ? (row[0] instanceof Date ? row[0].toISOString() : String(row[0])) : "",
          id: rowId,
          vendedor: String(row[2]).trim(),
          comprador: String(row[3]).trim(),
          albums: Number(row[4]) || 0,
          albums30k: Number(row[5]) || 0,
          paq1: Number(row[6]) || 0,
          paq4: Number(row[7]) || 0,
          paq10: Number(row[8]) || 0,
          total: Number(row[9]) || 0,
          efectivo: Number(row[10]) || 0,
          cbu: Number(row[11]) || 0,
          contado: Number(row[12]) || 0,
          diferencia: Number(row[13]) || 0,
          figuGigante: Number(row[14]) || 0,
          promo2x1: Number(row[15]) || 0
        };
      }
    }
    throw new Error("No se encontró la venta con el ID especificado.");
  } catch (e) {
    Logger.log("Error en getVentaPorId: " + e.toString());
    throw new Error("No se pudo cargar la venta: " + e.message);
  }
}

/**
 * Función de depuración temporal para inspeccionar la planilla de Punto de Venta
 */
function inspectPVSpreadsheet() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    const sheets = ss.getSheets();
    const result = [];
    for (let i = 0; i < sheets.length; i++) {
      const sheet = sheets[i];
      const name = sheet.getName();
      const lastRow = sheet.getLastRow();
      const lastCol = sheet.getLastColumn();
      let firstRows = [];
      if (lastRow > 0) {
        firstRows = sheet.getRange(1, 1, Math.min(lastRow, 5), Math.max(lastCol, 1)).getValues();
      }
      result.push({ name: name, lastRow: lastRow, lastCol: lastCol, firstRows: firstRows });
    }
    return result;
  } catch (e) {
    return { error: e.toString() };
  }
}

/**
 * Registra un comensal/interesado en la preventa de álbumes.
 */
function guardarRegistroPreventa(nombre, categoria, telefono, cantidad, observaciones) {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    let sheet = ss.getSheetByName("Preventa");
    if (!sheet) {
      sheet = ss.insertSheet("Preventa");
      sheet.appendRow(["Timestamp", "ID", "Nombre", "Categoria", "Telefono", "Cantidad", "Observaciones", "Estado"]);
      sheet.getRange(1, 1, 1, 8).setBackground("#434343").setFontColor("#ffffff").setFontWeight("bold");
    }

    const timestamp = new Date();
    const newId = "PR-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
    const estado = "Pendiente";

    sheet.appendRow([
      timestamp,
      newId,
      nombre,
      categoria,
      telefono,
      Number(cantidad) || 1,
      observaciones,
      estado
    ]);

    return { success: true, id: newId };
  } catch (e) {
    Logger.log("Error en guardarRegistroPreventa: " + e.toString());
    throw new Error("No se pudo registrar en la preventa: " + e.message);
  }
}

/**
 * Exporta el listado de reservas de preventa pública a un PDF Base64.
 */
function exportarPreventaPdf() {
  try {
    const ss = SpreadsheetApp.openById(SPREADSHEET_PV_ID);
    const sheet = ss.getSheetByName("Preventa");
    if (!sheet) throw new Error("No hay registros de preventa aún.");

    const values = sheet.getDataRange().getValues();
    if (values.length <= 1) throw new Error("No hay registros de preventa para exportar.");

    // Generar contenido HTML premium para el PDF
    let html = `
      <html>
        <head>
          <style>
            body { font-family: 'Helvetica Neue', Helvetica, Arial, sans-serif; color: #333; margin: 20px; }
            h1 { text-align: center; color: #e65c00; font-size: 24px; margin-bottom: 5px; }
            h2 { text-align: center; color: #666; font-size: 12px; margin-top: 0; text-transform: uppercase; letter-spacing: 1px; }
            table { width: 100%; border-collapse: collapse; margin-top: 20px; font-size: 10px; }
            th { background-color: #241203; color: white; padding: 10px; border: 1px solid #ddd; text-align: left; text-transform: uppercase; font-size: 9px; }
            td { padding: 8px; border: 1px solid #ddd; }
            tr:nth-child(even) { background-color: #f9f9f9; }
            .footer { text-align: center; margin-top: 30px; font-size: 10px; color: #999; border-top: 1px solid #eee; padding-top: 10px; }
            .badge { background-color: #e6f4ea; color: #137333; padding: 3px 8px; font-weight: bold; border-radius: 4px; display: inline-block; }
          </style>
        </head>
        <body>
          <h1>Club Olivos Hockey</h1>
          <h2>Reserva de Álbumes - Listado Oficial</h2>
          <table>
            <thead>
              <tr>
                <th style="width: 15%">Fecha</th>
                <th style="width: 15%">ID</th>
                <th style="width: 25%">Jugadora</th>
                <th style="width: 15%">Categoría</th>
                <th style="width: 15%">Contacto</th>
                <th style="width: 7%">Cant.</th>
                <th style="width: 13%">Estado</th>
              </tr>
            </thead>
            <tbody>
    `;

    // Cabeceras de la planilla
    const headers = values[0];
    const idxTimestamp = headers.indexOf("Timestamp");
    const idxId = headers.indexOf("ID");
    const idxNombre = headers.indexOf("Nombre de la Jugadora");
    const idxCategoria = headers.indexOf("Categoría");
    const idxTelefono = headers.indexOf("Teléfono de Contacto");
    const idxCantidad = headers.indexOf("Cantidad de Álbumes");
    const idxEstado = headers.indexOf("Estado");

    for (let i = 1; i < values.length; i++) {
      const row = values[i];
      const dateVal = row[idxTimestamp];
      let dateStr = "";
      if (dateVal instanceof Date) {
        dateStr = dateVal.toLocaleDateString('es-AR') + ' ' + dateVal.toLocaleTimeString('es-AR', { hour: '2-digit', minute: '2-digit' });
      } else {
        dateStr = String(dateVal);
      }

      const id = String(row[idxId]);
      const nombre = String(row[idxNombre]);
      const categoria = String(row[idxCategoria]);
      const telefono = String(row[idxTelefono]);
      const cantidad = String(row[idxCantidad]);
      const estado = String(row[idxEstado] || 'Pendiente');

      html += `
        <tr>
          <td>${dateStr}</td>
          <td style="font-family: monospace; font-size: 9px;">${id}</td>
          <td style="font-weight: bold;">${nombre}</td>
          <td>${categoria}</td>
          <td>${telefono}</td>
          <td style="text-align: center; font-weight: bold;">${cantidad}</td>
          <td><span class="badge">${estado}</span></td>
        </tr>
      `;
    }

    html += `
            </tbody>
          </table>
          <div class="footer">
            Generado automáticamente el ${new Date().toLocaleString('es-AR')} • Club Olivos Hockey
          </div>
        </body>
      </html>
    `;

    // Convertir HTML a PDF Blob y luego a Base64
    const htmlOutput = HtmlService.createHtmlOutput(html);
    const pdfBlob = htmlOutput.getAs('application/pdf').setName("Reserva_Albumes_OLIVOS.pdf");
    const base64 = Utilities.base64Encode(pdfBlob.getBytes());
    
    return {
      success: true,
      pdfBase64: base64,
      fileName: "Reserva_Albumes_OLIVOS.pdf"
    };
  } catch (e) {
    Logger.log("Error en exportarPreventaPdf: " + e.toString());
    throw new Error("No se pudo generar el PDF: " + e.message);
  }
}