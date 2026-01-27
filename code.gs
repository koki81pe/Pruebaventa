// MOD-001: ENCABEZADO DEL PROYECTO [INICIO]
/*
******************************************
PROYECTO: Prueba Ventas
ARCHIVO: code.gs
VERSIÓN: 01.01
FECHA: 27/01/2026 09:40 (UTC-5)
******************************************
*/
// MOD-001: FIN

// MOD-002: CONFIGURACIÓN GLOBAL [INICIO]
/**
 * Configuración central del sistema
 * Contiene API Key, nombre de hoja y URL del modelo Gemini
 */
const CONFIG = {
  API_KEY: 'AIzaSyDVejACx27_tL3j3vX65LXojKsRrfbdZ1U', 
  SHEET_NAME: 'Ventas',
  MODEL_URL: 'https://generativelanguage.googleapis.com/v1beta/models/gemini-2.5-flash:generateContent'
};
// MOD-002: FIN

// MOD-003: INTERFAZ DE USUARIO - MENÚ [INICIO]
/**
 * Crea el menú personalizado al abrir el Sheet
 * Opciones: Registrar por Voz y Diagnóstico de Modelos
 */
function onOpen() {
  SpreadsheetApp.getUi()
      .createMenu('Ventas voz')
      .addItem('Registrar por Voz', 'mostrarDialogoVoz')
      .addItem('🔍 Diagnóstico de Modelos', 'verModelosDisponibles')
      .addToUi();
}
// MOD-003: FIN

// MOD-004: INTERFAZ DE USUARIO - DIÁLOGO VOZ [INICIO]
/**
 * Muestra el diálogo modal con el botón de micrófono
 * Carga el archivo HTML index.html
 */
function mostrarDialogoVoz() {
  const html = HtmlService.createHtmlOutputFromFile('index')
      .setTitle('Asistente de Voz')
      .setWidth(400)
      .setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(html, ' ');
}
// MOD-004: FIN

// MOD-005: LÓGICA DE NEGOCIO - REGISTRO EN SHEET [INICIO]
/**
 * Registra los datos extraídos en la hoja de cálculo
 * CORRECCIÓN v01.01: Agrega fórmula en columna E (Precio x Cantidad)
 * @param {Object} datos - Objeto con producto, precio y cantidad
 * @return {String} Mensaje de confirmación o error
 */
function registrarVentaEnSheet(datos) {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) throw new Error(`No se encontró la hoja: ${CONFIG.SHEET_NAME}`);

    // Formatear fecha como DD/MM/AAAA
    const fechaCompleta = new Date();
    const dia = String(fechaCompleta.getDate()).padStart(2, '0');
    const mes = String(fechaCompleta.getMonth() + 1).padStart(2, '0');
    const anio = fechaCompleta.getFullYear();
    const fechaFormateada = `${dia}/${mes}/${anio}`;
    
    // Obtener la última fila para calcular la nueva posición
    const ultimaFila = sheet.getLastRow() + 1;
    
    // Orden de columnas: Fecha, Producto, Precio, Cantidad, Total (fórmula)
    sheet.appendRow([
      fechaFormateada, 
      datos.producto || "Desconocido", 
      datos.precio || 0, 
      datos.cantidad || 1,
      '' // Columna E vacía temporalmente
    ]);
    
    // Insertar fórmula en columna E (Total = Precio x Cantidad)
    const celdaFormula = sheet.getRange(ultimaFila, 5); // Fila nueva, columna E
    celdaFormula.setFormula(`=C${ultimaFila}*D${ultimaFila}`);

    return `✅ Registrado: ${datos.cantidad} ${datos.producto} a S/ ${datos.precio}`;
  } catch (e) {
    return "❌ Error al escribir en Sheet: " + e.toString();
  }
}
// MOD-005: FIN

// MOD-006: INTEGRACIÓN IA - PROCESAMIENTO GEMINI [INICIO]
/**
 * Envía el texto transcrito a Gemini AI para extraer datos estructurados
 * @param {String} textoVoz - Transcripción de voz del usuario
 * @return {Object} Datos extraídos: {producto, precio, cantidad}
 */
function procesarVozConGemini(textoVoz) {
  try {
    const prompt = `Actúa como un sistema de inventario. Extrae los datos de esta venta: "${textoVoz}".
    Responde estrictamente en formato JSON plano con esta estructura:
    {"producto": string, "precio": number, "cantidad": number}
    Si no detectas precio o cantidad, usa null para precio y 1 para cantidad. No escribas nada más que el JSON.`;

    const payload = {
      "contents": [{ "parts": [{ "text": prompt }] }]
    };

    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    const response = UrlFetchApp.fetch(`${CONFIG.MODEL_URL}?key=${CONFIG.API_KEY}`, options);
    const resContent = JSON.parse(response.getContentText());
    
    if (resContent.error) throw new Error(resContent.error.message);

    // Extraer y limpiar el texto de la respuesta
    let rawText = resContent.candidates[0].content.parts[0].text;
    let jsonString = rawText.replace(/```json|```/g, "").trim();
    
    return JSON.parse(jsonString);

  } catch (e) {
    throw new Error("Fallo en Gemini: " + e.toString());
  }
}
// MOD-006: FIN

// MOD-007: ORQUESTADOR PRINCIPAL [INICIO]
/**
 * Función principal que coordina el flujo completo
 * Recibe transcripción → Procesa con Gemini → Registra en Sheet
 * @param {String} transcripcion - Texto capturado por reconocimiento de voz
 * @return {String} Mensaje de resultado
 */
function ejecutorPrincipal(transcripcion) {
  const datosExtraidos = procesarVozConGemini(transcripcion);
  return registrarVentaEnSheet(datosExtraidos);
}
// MOD-007: FIN

// MOD-008: UTILIDAD - DIAGNÓSTICO DE MODELOS [INICIO]
/**
 * Lista todos los modelos Gemini disponibles en la cuenta
 * Útil para debugging y verificación de acceso a la API
 */
function verModelosDisponibles() {
  const url = `https://generativelanguage.googleapis.com/v1beta/models?key=${CONFIG.API_KEY}`;
  const response = UrlFetchApp.fetch(url, {muteHttpExceptions: true});
  const res = JSON.parse(response.getContentText());
  const ui = SpreadsheetApp.getUi();
  
  if (res.models) {
    let lista = res.models.map(m => m.name).join('\n');
    ui.alert("Modelos disponibles en tu cuenta:\n" + lista);
  } else {
    ui.alert("Error de Google: " + response.getContentText());
  }
}
// MOD-008: FIN

// MOD-009: NOTAS TÉCNICAS [INICIO]
/**
 * NOTAS DE IMPLEMENTACIÓN
 * 
 * Historial de versiones:
 * - v01.00: Versión inicial con formato fecha DD/MM/AAAA
 * - v01.01: Agrega columna E con fórmula =C*D (Precio x Cantidad)
 * 
 * Dependencias:
 * - Google Apps Script
 * - Gemini AI API (v1beta)
 * - Archivo HTML: index.html (interfaz de usuario)
 * 
 * Estructura de datos en Sheet:
 * Columna A: Fecha (DD/MM/AAAA)
 * Columna B: Producto (texto)
 * Columna C: Precio (número)
 * Columna D: Cantidad (número)
 * Columna E: Total (fórmula =C*D)
 * 
 * Próximas mejoras sugeridas:
 * - Validación de duplicados
 * - Histórico de modificaciones
 * - Exportación a PDF
 * - Dashboard de ventas
 */
// MOD-009: FIN
