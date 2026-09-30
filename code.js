/**
 * SISTEMA DE GENERACIÓN DE PLANEACIÓN DIDÁCTICA CON GEMINI API
 * Especializado para Licenciatura en Derecho / DGIRE
 */

const GEMINI_CONFIG = {
  MODEL: 'gemini-3.6-flash',
  TEMPERATURE: 0.6
};

/**
 * Función central con Reintento Automático Extendida (5 Intentos)
 */
function ejecutarConsultaGemini(systemPrompt, userPrompt) {
  const apiKey = PropertiesService.getScriptProperties().getProperty('GEMINI_API_KEY');
  
  if (!apiKey) {
    throw new Error("API Key no encontrada en ScriptProperties.");
  }

  const url = `https://generativelanguage.googleapis.com/v1beta/models/${GEMINI_CONFIG.MODEL}:generateContent?key=${apiKey}`;

  const payload = {
    "system_instruction": {
      "parts": [{ "text": systemPrompt }]
    },
    "contents": [{
      "role": "user",
      "parts": [{ "text": `Tema a desarrollar: "${userPrompt}"` }]
    }],
    "generationConfig": {
      "temperature": GEMINI_CONFIG.TEMPERATURE,
      "maxOutputTokens": 2000 // Asegura que las respuestas no se corten
    }
  };

  const options = {
    'method': 'post',
    'contentType': 'application/json',
    'payload': JSON.stringify(payload),
    'muteHttpExceptions': true
  };

  // Pausa preventiva de 8 segundos antes de llamar a la API para respetar la cuota gratuita
  Utilities.sleep(200);

  // Intentar hasta 5 veces si la API responde con error de límite (429)
  for (let intento = 1; intento <= 5; intento++) {
    try {
      const response = UrlFetchApp.fetch(url, options);
      const json = JSON.parse(response.getContentText());

      if (response.getResponseCode() === 200) {
        return json.candidates[0].content.parts[0].text.trim();
      }

      // Si nos pasamos de la cuota (Error 429), esperamos 15 segundos y reintentamos
      if (response.getResponseCode() === 429 || json.error?.code === 429) {
        Logger.log(`Intento ${intento}: Cuota excedida. Esperando 15 segundos...`);
        Utilities.sleep(15000); 
        continue;
      }

      return `[Error API ${response.getResponseCode()}]: ${json.error?.message || 'Error de conexión'}`;

    } catch (e) {
      if (intento === 5) return `[Error Script]: ${e.toString()}`;
      Utilities.sleep(5000);
    }
  }

  return "[Error]: Cuota excedida. Por favor espera 1 minuto antes de volver a ejecutar.";
}

// ============================================================================
// PROMPTS ESPECIALIZADOS
// ============================================================================

function obtenerPromptObjetivos() {
  return `Eres un diseñador instruccional y docente experto en la Licenciatura en Derecho.
Tu tarea es redactar los "Objetivos o Resultados de Aprendizaje" para el tema proporcionado.

DIRECTRICES OBLIGATORIAS:
- Basado en las Taxonomías de Bloom o Marzano para nivel universitario. Usa verbos observables y medibles (ej. identificar, analizar, comparar, evaluar, argumentar).
- Abarca desde la comprensión conceptual hasta el análisis o aplicación a casos jurídicos.
- Redacta de 3 a 4 puntos clave, claros y directos.
- FORMATO: Presenta cada objetivo mediante una lista de viñetas breves. Resalta en **negrita** el verbo principal.
- RESTRICCIÓN: Sin introducciones, saludos ni textos finales. Devuelve únicamente los objetivos.`;
}

function obtenerPromptActividades() {
  return `Eres un docente universitario experto en la enseñanza del Derecho.
Tu tarea es redactar las "Actividades de Enseñanza-Aprendizaje" estructuradas rigurosamente en 3 momentos pedagógicos.

DIRECTRICES DE ESTRUCTURA (OBLIGATORIAS):
Debes usar de forma estricta los tres encabezados siguientes en negrita y en líneas separadas:

**INICIO:**
- Plantea una actividad breve de activación (pregunta detonadora, análisis de una noticia jurídica reciente, caso hipotético o recuperación de conocimientos previos).

**DESARROLLO:**
- Diseña una actividad práctica e interactiva (análisis de artículos/jurisprudencia, resolución colaborativa de un problema legal, debate guiado o simulación).

**CIERRE:**
- Establece una actividad de síntesis y evaluación (reflexión grupal, conclusiones técnicas, retroalimentación o formulación de un argumento final).

DIRECTRICES DE ESTILO:
- Mantén un lenguaje técnico-jurídico adecuado pero sumamente directo y concreto.
- Cada fase (Inicio, Desarrollo y Cierre) debe constar de 2 a 4 líneas descriptivas.
- RESTRICCIÓN: Asegúrate de concluir completamente cada oración. No incluyas saludos ni textos de presentación.`;
}

// ============================================================================
// FUNCIONES DE EJECUCIÓN
// ============================================================================

function generarPlaneacionFilaSeleccionada() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const fila = hoja.getActiveCell().getRow();

  if (fila < 2) {
    SpreadsheetApp.getUi().alert("Por favor selecciona una fila válida (fila 2 o superior).");
    return;
  }

  procesarFila(hoja, fila);
}

/**
 * Detecta dinámicamente cualquier rango seleccionado (ejemplo: C12 a C16)
 * y actualiza sus respectivas filas en las columnas D y E.
 */
function generarPlaneacionRangoSeleccionado() {
  const hoja = SpreadsheetApp.getActiveSpreadsheet().getActiveSheet();
  const rango = hoja.getActiveRange();
  
  const filaInicio = rango.getRow();
  const numFilas = rango.getNumRows();
  const filaFin = filaInicio + numFilas - 1;

  if (filaInicio < 2) {
    SpreadsheetApp.getUi().alert("Asegúrate de seleccionar filas a partir de la fila 2.");
    return;
  }

  const confirmacion = SpreadsheetApp.getUi().alert(
    'Confirmación de Procesamiento',
    `¿Deseas procesar el rango de filas de la ${filaInicio} a la ${filaFin}?`,
    SpreadsheetApp.getUi().ButtonSet.YES_NO
  );

  if (confirmacion !== SpreadsheetApp.getUi().Button.YES) return;

  // Bucle dinámico que recorre exactamente el rango sombreado
  for (let fila = filaInicio; fila <= filaFin; fila++) {
    procesarFila(hoja, fila);
  }

  SpreadsheetApp.getUi().alert(`✨ ¡Planeación completada con éxito para las filas ${filaInicio} a ${filaFin}!`);
}

function procesarFila(hoja, fila) {
  const tema = hoja.getRange(fila, 3).getValue(); // Lee siempre el tema de la Columna C

  if (!tema || tema.toString().trim() === "") return;

  // Notificar estado en las Columnas D y E de esa misma fila
  hoja.getRange(fila, 4).setValue("⏳ Generando Objetivos...");
  hoja.getRange(fila, 5).setValue("⏳ Generando Actividades...");
  SpreadsheetApp.flush();

  // 1. Procesar Objetivos (Escribe en Columna D)
  const objetivos = ejecutarConsultaGemini(obtenerPromptObjetivos(), tema);
  hoja.getRange(fila, 4).setValue(objetivos);
  SpreadsheetApp.flush();

  // 2. Procesar Actividades (Escribe en Columna E)
  const actividades = ejecutarConsultaGemini(obtenerPromptActividades(), tema);
  hoja.getRange(fila, 5).setValue(actividades);
  SpreadsheetApp.flush();
}

function onOpen() {
  SpreadsheetApp.getUi()
    .createMenu('✨ IA Planeación')
    .addItem('⚡ Procesar solo la fila activa', 'generarPlaneacionFilaSeleccionada')
    .addItem('🚀 Procesar TODAS las filas sombreadas/seleccionadas', 'generarPlaneacionRangoSeleccionado')
    .addToUi();
}