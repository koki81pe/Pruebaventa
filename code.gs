// MOD-001: ENCABEZADO DEL PROYECTO [INICIO]
/*
******************************************
PROYECTO: Prueba Ventas
ARCHIVO: code.gs
VERSIÓN: 01.22
FECHA: 29/01/2026 23:44 (UTC-5)
******************************************
*/
// MOD-001: FIN

// MOD-002: CONFIGURACIÓN GLOBAL [INICIO]
/**
 * Configuración central del sistema
 * Contiene API Key, nombres de hojas y URL del modelo Gemini
 */
const CONFIG = {
  API_KEY: 'AIzaSyAqs2cSnSym-wTCsJ_dKoqijE9qPa6k-NA', 
  SHEET_NAME: 'Ventas',
  MPAGO_SHEET: 'Mpago',
  CAT_SHEET: 'Cat',
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
      .createMenu('🔈Ventas voz')
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

// MOD-005-S01: FUNCIÓN PRINCIPAL [INICIO]
/**
 * Registra los datos extraídos en la hoja de cálculo (8 columnas A-H)
 * V01.20: AGREGADO modo al objeto resultado para visualización HTML
 * V01.19: Última fila busca por columna A (fecha) en lugar de getLastRow()
 */
function registrarVentaEnSheet(datos) {
  console.log("┌─────────────────────────────────────────────────────┐");
  console.log("📝 MOD-005-S01 INICIO - Datos recibidos:", JSON.stringify(datos));
  
  try {
    // VALIDACIÓN OBLIGATORIA: producto y precio
    console.log("📝 MOD-005-S01 - Validando producto:", datos.producto);
    console.log("📝 MOD-005-S01 - Validando precio:", datos.precio, "tipo:", typeof datos.precio);
    
    if (!datos.producto || datos.precio == null || datos.precio === 0) {
      console.error("🔴 MOD-005-S01 ERROR - Validación falló");
      console.error("🔴 MOD-005-S01 - producto:", datos.producto);
      console.error("🔴 MOD-005-S01 - precio:", datos.precio);
      return {
        exito: false,
        mensaje: `❌ Falta producto o precio válido: ${JSON.stringify(datos)}`
      };
    }
    
    console.log("📝 MOD-005-S01 - Validación OK, obteniendo sheet...");
    
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.SHEET_NAME);
    
    if (!sheet) {
      console.error("🔴 MOD-005-S01 ERROR - Sheet no encontrado:", CONFIG.SHEET_NAME);
      throw new Error(`No se encontró la hoja: ${CONFIG.SHEET_NAME}`);
    }
    
    console.log("📝 MOD-005-S01 - Sheet encontrado:", CONFIG.SHEET_NAME);
    console.log("📝 MOD-005-S01 - Llamando a prepararDatosValidados...");
    
    // Preparar datos validados
    const datosValidados = prepararDatosValidados(datos);
    
    console.log("📝 MOD-005-S01 - Datos validados recibidos:", JSON.stringify(datosValidados));
    
    // V01.19: Obtener última fila por columna A (fecha)
    const ultimaFila = obtenerUltimaFilaPorFecha(sheet) + 1;
    console.log("📝 MOD-005-S01 - Última fila calculada por fecha:", ultimaFila);
    
    // Insertar fila con datos
    console.log("📝 MOD-005-S01 - Llamando a insertarFilaDatos...");
    insertarFilaDatos(sheet, datosValidados, ultimaFila);
    
    // Aplicar fórmulas y formatos
    console.log("📝 MOD-005-S01 - Llamando a aplicarFormulasYFormatos...");
    aplicarFormulasYFormatos(sheet, ultimaFila);
    
    // V01.20: Calcular total para retornar
    const total = datosValidados.precio * datosValidados.cantidad;
    
    const resultado = {
      exito: true,
      mensaje: `✅ ${datosValidados.cantidad} ${datosValidados.producto} a S/ ${datosValidados.precio}`,
      datos: {
        producto: datosValidados.producto,
        precio: datosValidados.precio,
        cantidad: datosValidados.cantidad,
        modo: datosValidados.modo,  // 🆕 v01.20: PARA HTML
        total: total
      }
    };
    
    console.log("📝 MOD-005-S01 SUCCESS v01.20 - Resultado CON MODO:", JSON.stringify(resultado));
    console.log("└─────────────────────────────────────────────────────┘");
    
    return resultado;
    
  } catch (e) {
    console.error("🔴 MOD-005-S01 ERROR - Excepción:", e.toString());
    console.error("🔴 MOD-005-S01 ERROR - Stack:", e.stack);
    console.log("└─────────────────────────────────────────────────────┘");
    return {
      exito: false,
      mensaje: `❌ Error al escribir en Sheet: ${e.toString()}`
    };
  }
}
// MOD-005-S01: FIN

// MOD-005-S02: PREPARAR DATOS VALIDADOS [INICIO]
/**
 * Valida y capitaliza los campos necesarios para el registro
 * v01.20: Debug removido - DEBUG HTML L3 maneja logging
 */
function prepararDatosValidados(datos) {
  // CLASIFICAR categoria y modo usando listas de referencia
  const categorias = obtenerCategorias();
  const modosPago = obtenerModosPago();
  
  const categoriaValidada = validarCampo(datos.categoria, categorias, true);
  const modoValidado = validarCampo(datos.modo, modosPago, false);
  
  // CAPITALIZAR PRIMERA LETRA
  const categoriaFinal = capitalizarPrimeraLetra(categoriaValidada);
  const productoFinal = capitalizarPrimeraLetra(datos.producto);
  const modoFinal = capitalizarPrimeraLetra(modoValidado);
  
  // Formatear fecha como DD/MM/AAAA
  const fechaCompleta = new Date();
  const dia = String(fechaCompleta.getDate()).padStart(2, '0');
  const mes = String(fechaCompleta.getMonth() + 1).padStart(2, '0');
  const anio = fechaCompleta.getFullYear();
  const fechaFormateada = `${dia}/${mes}/${anio}`;
  
  const resultado = {
    fecha: fechaFormateada,
    categoria: categoriaFinal,
    producto: productoFinal,
    modo: modoFinal,
    precio: datos.precio,
    cantidad: datos.cantidad || 1
  };
  
  return resultado;
}
// MOD-005-S02: FIN

// MOD-005-S03: INSERTAR FILA DE DATOS [INICIO]
/**
 * Inserta una nueva fila con los datos validados
 * ESTRUCTURA: A:Fecha, B:Cat, C:Producto, D:Modo, E:Precio, F:Cant, G:Total, H:Acumulado
 */
function insertarFilaDatos(sheet, datosValidados, ultimaFila) {
  sheet.appendRow([
    datosValidados.fecha,      // A
    datosValidados.categoria,  // B: Capitalizada
    datosValidados.producto,   // C: Capitalizada
    datosValidados.modo,       // D: Capitalizada
    datosValidados.precio,     // E
    datosValidados.cantidad,   // F
    '',                        // G: Fórmula Total
    ''                         // H: Fórmula Acumulado (NUEVO)
  ]);
}
// MOD-005-S03: FIN

// MOD-005-S04: APLICAR FÓRMULAS Y FORMATOS [INICIO]
/**
 * Aplica las fórmulas de Total y Acumulado con formato numérico
 * NUEVO V01.09: Formato números #,##0.00
 */
function aplicarFormulasYFormatos(sheet, ultimaFila) {
  // Fórmula Total en G (E*F)
  const celdaTotal = sheet.getRange(ultimaFila, 7);  // Columna G
  celdaTotal.setFormula(`=E${ultimaFila}*F${ultimaFila}`);
  
  // FORMATO NÚMERO G: 5,350.67 (NUEVO V01.09)
  celdaTotal.setNumberFormat("#,##0.00");
  
  // Fórmula Acumulado en H (NUEVO V01.09)
  const celdaAcumulado = sheet.getRange(ultimaFila, 8);  // Columna H
  if (ultimaFila === 2) {
    // Primera fila de datos: H2 = G2
    celdaAcumulado.setFormula(`=G${ultimaFila}`);
  } else {
    // Filas siguientes: Hn = Gn + H(n-1)
    celdaAcumulado.setFormula(`=G${ultimaFila}+H${ultimaFila-1}`);
  }
  
  // FORMATO NÚMERO H: 5,350.67 (NUEVO V01.09)
  celdaAcumulado.setNumberFormat("#,##0.00");
}
// MOD-005-S04: FIN

// MOD-005: FIN

// MOD-006: INTEGRACIÓN IA - PROCESAMIENTO GEMINI [INICIO]
/**
 * Envía el texto transcrito a Gemini AI para extraer MÚLTIPLES ventas
 * LOGS PERMANENTES: Debug completo de Gemini
 */
function procesarVozConGemini(textoVoz) {
  console.log("───────────────────────────────────────────────────────");
  console.log("🤖 MOD-006 INICIO - Texto recibido:", textoVoz);
  
  try {
    const prompt = `Actúa como sistema de inventario. Extrae TODAS las ventas de: "${textoVoz}".

REGLAS:
- SIEMPRE detecta: producto + precio (números con ".") 
- cantidad = 1 si no se menciona
- categoria y modo son OPCIONALES (dejar null si no se mencionan)
- Separa ventas: ";", "y", "," 
- Remate "todos son X, pagados en Y" aplica a todas anteriores

EJEMPLO "Un peluche a 35 soles":
[{"producto": "peluche", "precio": 35, "cantidad": 1, "categoria": null, "modo": null}]

Responde SOLO JSON array válido:`;

    console.log("🤖 MOD-006 - Prompt preparado (primeras 100 chars):", prompt.substring(0, 100) + "...");

    const payload = {
      "contents": [{ "parts": [{ "text": prompt }] }]
    };

    const options = {
      "method": "post",
      "contentType": "application/json",
      "payload": JSON.stringify(payload),
      "muteHttpExceptions": true
    };

    console.log("🤖 MOD-006 - Enviando request a Gemini...");
    console.log("🤖 MOD-006 - URL:", CONFIG.MODEL_URL);
    
    const response = UrlFetchApp.fetch(`${CONFIG.MODEL_URL}?key=${CONFIG.API_KEY}`, options);
    const responseCode = response.getResponseCode();
    const responseText = response.getContentText();
    
    console.log("🤖 MOD-006 - Response HTTP Code:", responseCode);
    console.log("🤖 MOD-006 - Response completo:", responseText);
    
    if (responseCode !== 200) {
      console.error("🔴 MOD-006 ERROR - HTTP no es 200:", responseCode);
      throw new Error("HTTP " + responseCode + ": " + responseText);
    }
    
    const resContent = JSON.parse(responseText);
    
    if (resContent.error) {
      console.error("🔴 MOD-006 ERROR - Gemini devolvió error:", JSON.stringify(resContent.error));
      throw new Error(resContent.error.message);
    }

    let rawText = resContent.candidates[0].content.parts[0].text;
    console.log("🤖 MOD-006 - RAW text de Gemini:", rawText);
    
    let jsonString = rawText.replace(/```json|```/g, "").trim();
    console.log("🤖 MOD-006 - JSON limpio:", jsonString);
    
    const ventasArray = JSON.parse(jsonString);
    
    console.log("🤖 MOD-006 - Parsed exitoso, tipo:", typeof ventasArray);
    console.log("🤖 MOD-006 - Es array?:", Array.isArray(ventasArray));
    console.log("🤖 MOD-006 - Length:", ventasArray.length);
    console.log("🤖 MOD-006 - Contenido:", JSON.stringify(ventasArray));
    
    if (!Array.isArray(ventasArray)) {
      console.error("🔴 MOD-006 ERROR - No es un array válido");
      throw new Error("Gemini no devolvió array válido");
    }
    
    console.log("🤖 MOD-006 SUCCESS - Retornando", ventasArray.length, "ventas");
    console.log("───────────────────────────────────────────────────────");
    return ventasArray;

  } catch (e) {
    console.error("🔴 MOD-006 ERROR - Excepción capturada:", e.toString());
    console.error("🔴 MOD-006 ERROR - Tipo:", e.name);
    console.error("🔴 MOD-006 ERROR - Stack:", e.stack);
    console.log("───────────────────────────────────────────────────────");
    
    // NO retornar fallback, mejor propagar el error
    throw new Error("Error procesando con Gemini: " + e.toString());
  }
}
// MOD-006: FIN

// MOD-007: ORQUESTADOR PRINCIPAL [INICIO]
/**
 * Función principal que coordina el flujo completo MÚLTIPLE VENTAS
 * V01.19: Retorna objeto estructurado con datos para mostrar en HTML
 * @param {String} transcripcion - Texto capturado por reconocimiento de voz
 * @return {Object} Objeto con exito, mensaje y ventas[]
 */
function ejecutorPrincipal(transcripcion) {
  console.log("═══════════════════════════════════════════════════════");
  console.log("🎯 MOD-007 INICIO - Transcripción recibida:", transcripcion);
  
  try {
    // Obtener array de ventas desde Gemini
    console.log("🎯 MOD-007 - Llamando a procesarVozConGemini...");
    const ventasArray = procesarVozConGemini(transcripcion);
    
    console.log("🎯 MOD-007 - Ventas recibidas de Gemini:", ventasArray.length);
    console.log("🎯 MOD-007 - Array completo:", JSON.stringify(ventasArray));
    
    let ventasRegistradas = [];
    let mensajes = [];
    let totalVentas = 0;
    let totalGeneral = 0;
    
    // Registrar cada venta individualmente
    for (let i = 0; i < ventasArray.length; i++) {
      const venta = ventasArray[i];
      console.log(`🎯 MOD-007 - Procesando venta ${i + 1}/${ventasArray.length}:`, JSON.stringify(venta));
      
      const resultado = registrarVentaEnSheet(venta);
      console.log(`🎯 MOD-007 - Resultado venta ${i + 1}:`, JSON.stringify(resultado));
      
      if (resultado.exito) {
        ventasRegistradas.push(resultado.datos);
        mensajes.push(resultado.mensaje);
        totalVentas++;
        totalGeneral += resultado.datos.total;
      } else {
        mensajes.push(resultado.mensaje);
      }
    }
    
    // V01.19: Retornar objeto estructurado
    const respuesta = {
      exito: totalVentas > 0,
      totalVentas: totalVentas,
      totalGeneral: totalGeneral,
      ventas: ventasRegistradas,
      mensajes: mensajes
    };
    
    console.log("🎯 MOD-007 SUCCESS - Respuesta:", JSON.stringify(respuesta));
    console.log("═══════════════════════════════════════════════════════");
    
    return respuesta;
    
  } catch (e) {
    console.error("🔴 MOD-007 ERROR - Excepción:", e.toString());
    console.error("🔴 MOD-007 ERROR - Stack:", e.stack);
    console.log("═══════════════════════════════════════════════════════");
    return {
      exito: false,
      totalVentas: 0,
      totalGeneral: 0,
      ventas: [],
      mensajes: [`❌ Error en orquestador: ${e.toString()}`]
    };
  }
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

// MOD-009: FUNCIONES AUXILIARES [INICIO]
/**
 * UTILIDADES PARA VALIDACIÓN DE LISTAS REFERENCIA + CAPITALIZACIÓN
 * v01.21: Debug removido - DEBUG HTML L3 maneja logging
 * V01.18: ELIMINADO Camino 1 (>65% directo con .find())
 * ÚNICO CAMINO: Ranking completo >=30% con tie-breaker orden hoja
 * Mpago: similitud 30% | Cat: prefijo 70% + Jaccard 30%
 */

function similitudLetras(texto1, texto2) {
  // NORMALIZAR TILDES
  const normalizar = (str) => str.toLowerCase()
    .replace(/[áä]/g,'a').replace(/[éë]/g,'e')
    .replace(/[íï]/g,'i').replace(/[óö]/g,'o')
    .replace(/[úü]/g,'u')
    .replace(/[aeiou]/g, '').replace(/[^a-z]/g, '');
  
  const set1 = new Set(normalizar(texto1));
  const set2 = new Set(normalizar(texto2));
  
  if (set1.size === 0 || set2.size === 0) return 0;
  
  const interseccion = new Set([...set1].filter(x => set2.has(x)));
  const union = new Set([...set1, ...set2]);
  
  return interseccion.size / union.size;
}

function scorePrefijo(texto, opcion) {
  const t = texto.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,"");
  const o = opcion.toLowerCase().normalize('NFD').replace(/[\u0300-\u036f]/g,"");
  const minLen = Math.min(t.length, o.length);
  let coincidencias = 0;
  
  for(let i = 0; i < minLen; i++) {
    if(t[i] === o[i]) coincidencias++;
    else break;
  }
  
  return coincidencias / Math.max(t.length, o.length);
}

function capitalizarPrimeraLetra(texto) {
  if (!texto || typeof texto !== 'string') return '';
  return texto.charAt(0).toUpperCase() + texto.slice(1).toLowerCase();
}

function obtenerModosPago() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.MPAGO_SHEET);
    
    if (!sheet) {
      return [];
    }
    
    const valores = sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
      .flat()
      .filter(Boolean)
      .map(modo => modo.toString().toLowerCase().trim());
    
    return [...new Set(valores)];
    
  } catch (e) {
    console.error('MOD-009 obtenerModosPago ERROR:', e.toString());
    return [];
  }
}

function obtenerCategorias() {
  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const sheet = ss.getSheetByName(CONFIG.CAT_SHEET);
    
    if (!sheet) {
      return [];
    }
    
    const valores = sheet.getRange('A2:A' + sheet.getLastRow()).getValues()
      .flat()
      .filter(Boolean)
      .map(cat => cat.toString().toLowerCase().trim());
    
    return [...new Set(valores)];
    
  } catch (e) {
    console.error('MOD-009 obtenerCategorias ERROR:', e.toString());
    return [];
  }
}

function obtenerUltimaFilaPorFecha(sheet) {
  try {
    const maxRows = sheet.getMaxRows();
    const columnA = sheet.getRange(1, 1, maxRows, 1).getValues();
    
    // Buscar desde abajo hacia arriba la primera celda no vacía
    for (let i = columnA.length - 1; i >= 0; i--) {
      const valor = columnA[i][0];
      
      if (valor !== null && valor !== undefined && valor !== '') {
        return i + 1; // +1 porque arrays empiezan en 0
      }
    }
    
    return 1; // Cabecera si no hay datos
    
  } catch (e) {
    console.error('MOD-009 obtenerUltimaFilaPorFecha ERROR:', e.toString());
    return sheet.getLastRow(); // Fallback
  }
}

function validarCampo(texto, listaOpciones, esCategoria = false) {
  if (!texto || listaOpciones.length === 0) {
    return '';
  }
  
  const textoLimpio = texto.toString().toLowerCase().trim();
  
  // FIX DIRECTO "llave" → "yape"
  if (textoLimpio === 'llave') {
    return listaOpciones.find(opcion => opcion === 'yape') || '';
  }
  
  // PRIORIDAD 1: Exacto/substring
  const coincidencia = listaOpciones.find(opcion => 
    opcion === textoLimpio || 
    textoLimpio.includes(opcion) || 
    opcion.includes(textoLimpio)
  );
  
  if (coincidencia) {
    return coincidencia;
  }
  
  if (esCategoria) {
    // CAMINO ÚNICO: RANKING COMPLETO >=30% (V01.18)
    const matchesRefinados = listaOpciones
      .map(opcion => {
        const prefijoScore = scorePrefijo(textoLimpio, opcion);
        const jaccardScore = similitudLetras(textoLimpio, opcion);
        return {
          opcion,
          scoreFinal: (prefijoScore * 0.7) + (jaccardScore * 0.3),
          prefijo: prefijoScore,
          jaccard: jaccardScore,
          lenDiff: Math.abs(textoLimpio.length - opcion.length),
          indiceOriginal: listaOpciones.indexOf(opcion)
        };
      })
      .filter(m => m.scoreFinal >= 0.30)
      .sort((a, b) => {
        if (Math.abs(a.scoreFinal - b.scoreFinal) > 0.001) {
          return b.scoreFinal - a.scoreFinal;
        }
        return a.indiceOriginal - b.indiceOriginal;
      });
    
    return matchesRefinados[0]?.opcion || '';
    
  } else {
    // MPAGO: simple 30%
    const matches = listaOpciones
      .map(opcion => ({
        opcion,
        score: similitudLetras(textoLimpio, opcion)
      }))
      .filter(match => match.score > 0.30)
      .sort((a, b) => b.score - a.score);
    
    return matches[0]?.opcion || '';
  }
}
// MOD-009: FIN

// MOD-010: WEB APP PÚBLICA [INICIO]
/**
 * EndPoint público para acceso vía LINK directo
 * Mantiene menú Sheet + agrega Web App independiente
 */
function doGet() {
  return HtmlService.createHtmlOutputFromFile('index')
    .setTitle('🔈 Ventas Voz Web')
    .setXFrameOptionsMode(HtmlService.XFrameOptionsMode.ALLOWALL);
}
// MOD-010: FIN

// MOD-099: NOTAS TÉCNICAS [INICIO]
/**
 * Historial versiones:
 * v01.19: ✅ MOD-005-S01 busca última fila por columna A (fecha)
 *         ✅ MOD-007 retorna objeto estructurado para HTML
 *         ✅ index.html muestra ventas registradas (Opción A - lista simple)
 *         ✅ MOD-009 nueva función obtenerUltimaFilaPorFecha()
 * v01.18: ✅ MOD-009 ELIMINADO Camino 1 (>65% .find) → Solo ranking >=30%
 * v01.10: ✅ MOD-009 similitud 30% + fix llave→yape
 * v01.09: ✅ Múltiples ventas + 8 columnas + Mayúscula + Acumulado
 * 
 * ESTRUCTURA Ventas: Fecha | Cat | Producto | Modo | Precio | Cant | Total | Acumulado
 * Listas ref: Mpago(A2+), Cat(A2+) → validación: exacto/llave/ranking>=30%
 * Formato: G,H → "#,##0.00"
 * 
 * FIX v01.18: "Cosme" ahora detecta correctamente "Cosmético" vs "Comisión"
 * Método: Ranking completo con prefijo 70% + Jaccard 30% + tie-breaker orden hoja
 * 
 * FIX v01.19: Última fila se busca por columna A (fecha) para manejar filas vacías/extras
 * Visualización: HTML muestra ventas registradas formato "Producto - S/ X × Y = S/ Z"
 */
// MOD-099: FIN
