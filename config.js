/**
 * CONFIG.GS
 */

// ============================================================================
// CONFIGURACIÓN DE PLANILLAS
// ============================================================================

// URL de la planilla actual (Seguimiento de O2)
const PLANILLA_O2_URL = 'https://docs.google.com/spreadsheets/d/14QQccn4qmot0XKp-7tEwRSOJ1h2NOP-_fBk385pNrg0';

// URL de la planilla externa (Excel Cuidados Paliativos 2026)
const PLANILLA_CCP_URL = 'https://docs.google.com/spreadsheets/d/1XT9rGd-j3X_zKUopJ4YwQKXizGLI5Xkfr5S9U2fAlSs';

// Rango de importación de la planilla CCP
const PLANILLA_CCP_RANGO = 'Consolidado!A:AH';

// ============================================================================
// IDS DE CARPETAS DE GOOGLE DRIVE
// ============================================================================

const FOLDER_IDS = {
  RECETAS: '1c6HcZohX-5uvCg7w4E4YhjLBAliA12Gx',
  FICHAS_GENERADAS: '1waY_yPLevI5wTD6kruAMOKtcIv7UC8GD',
  FICHAS_FIRMADAS: '1qS3Yzp3lHX7Qx1JFH1pbR8_zwVshK1Vx',
  CERTIFICADOS_DEFUNCION: '1Heic-YbKSftCtcXtIqScpfQSySi1VCPN',
  DOCUMENTOS_RETIRO: '1SQ0uCJlne3lz5e_IsvGVm2vy1PkL3OOC',
  COMPROBANTES_ENTREGA: '1liVGs19LxH0LhgeO8qBtY9JpBMGc20jt',
  ORDENES_ENTREGA: '149k7IlFZRndNPLcHaFHsVwXEtxYGVU30'
};

// ============================================================================
// NOMBRES DE HOJAS
// ============================================================================

const HOJAS = {
  HISTORIAL: 'Historial Pacientes',
  INSTALACION: 'Instalación',
  RECARGAS: 'Recargas',
  CAJAS: 'Cajas',
  DATOS_CCP: 'Datos_CCP' // Hoja oculta para caché de datos
};

// ============================================================================
// ÍNDICES DE COLUMNAS - HISTORIAL PACIENTES
// ============================================================================

const COL_HISTORIAL = {
  // Sección Validación (a-i)
  VACIO_A: 1,
  ID_PACIENTE: 2,           // b
  NOMBRE_PROBLEMA: 3,        // c
  FECHA_GARANTIA: 4,         // d
  FECHA_IPD: 5,              // e
  ESTABLECIMIENTO: 6,        // f
  VIGENCIA_GARANTIA: 7,      // g
  ESTADO_PREVISIONAL: 8,     // h
  LEY_HCUH: 9,              // i
  
  // Sección Ingreso (j-aa)
  FECHA_INGRESO: 10,         // j
  RUT: 11,                   // k
  DV: 12,                    // l
  NOMBRE: 13,                // m
  APELLIDO_PAT: 14,          // n
  APELLIDO_MAT: 15,          // o
  DIAGNOSTICO: 16,           // p
  EDAD: 17,                  // q
  SEXO: 18,                  // r
  DOMICILIO: 19,             // s
  NUMERO: 20,                // t
  COMUNA: 21,                // u
  TELEFONO_PACIENTE: 22,     // v
  NOMBRE_TUTOR: 23,          // w
  TELEFONO_TUTOR: 24,        // x
  EMAIL: 25,                 // y
  NOTIFICADO_O2: 26,         // z
  
  // Sección Indicación (aa-ai)
  SOLICITANTE: 27,           // aa
  FLUJO_MEDICO: 28,          // ab
  TERAPIA: 29,               // ac
  TIPO_CONCENTRADOR: 30,     // ad
  RECETA_URL: 31,            // ae
  OBSERVACIONES: 32,         // af
  FICHA_FIRMADA_URL: 33,     // ag
  PACIENTE_VIGENTE: 34,      // ah
  REQUIERE_RECARGA: 35,      // ai
  
  // Sección Egreso (aj-ap)
  FECHA_EGRESO: 36,          // aj
  MOTIVO_EGRESO: 37,         // ak
  FECHA_ALTA_MEDICA: 38,     // al
  CERTIFICADO_DEF_URL: 39,   // am
  FECHA_RETIRO_EQUIPO: 40,   // an
  OBSERVACIONES_EGRESO: 41,  // ao
  DOC_RETIRO_URL: 42         // ap
};

// ============================================================================
// ÍNDICES DE COLUMNAS - INSTALACIÓN
// ============================================================================

const COL_INSTALACION = {
  ID_PACIENTE: 1,            // a
  RUT: 2,                    // b
  DV: 3,                     // c
  NOMBRE: 4,                 // d
  APELLIDO_PAT: 5,           // e
  APELLIDO_MAT: 6,           // f
  DIRECCION: 7,              // g
  FECHA_ENVIO_ORDEN: 8,      // h
  CONFIRMACION_EQUIPOS: 9,   // i
  PACIENTE_CONTESTA: 10,     // j
  FECHA_EVAL_TECNICA: 11,    // k
  FECHA_INSTALACION: 12,     // l
  FECHA_CAPACITACION: 13,    // m
  NOMBRE_CAPACITADO: 14,     // n
  COMPROBANTE_URL: 15,       // o
  CUMPLIMIENTO_ENTREGA: 16   // p
};

// ============================================================================
// ÍNDICES DE COLUMNAS - RECARGAS
// ============================================================================

const COL_RECARGAS = {
  ID_PACIENTE: 1,            // a
  FECHA_SOLICITUD: 2,        // b
  RUT: 3,                    // c
  DV: 4,                     // d
  NOMBRE: 5,                 // e
  APELLIDO_PAT: 6,           // f
  APELLIDO_MAT: 7,           // g
  DIRECCION: 8,              // h
  OBSERVACIONES: 9,          // i
  NUM_SERVICIO: 10,          // j
  FECHA_ENTREGA: 11,         // k
  ORDEN_URL: 12,             // l
  CUMPLIMIENTO_ENTREGA: 13   // m
};

// ============================================================================
// ÍNDICES DE COLUMNAS - CAJAS
// ============================================================================

const COL_CAJAS = {
  ID_PACIENTE: 1,            // a
  RUT: 2,                    // b
  DV: 3,                     // c
  NOMBRE_COMPLETO: 4,        // d
  EDAD: 5,                   // e
  TELEFONO: 6,               // f
  MAIL: 7,                   // g
  FECHA_EMISION_RECETA: 8,   // h
  FECHA_INGRESO: 9,          // i
  FECHA_SOLICITUD_INST: 10,  // j
  FECHA_ENTREGA_RECARGA: 11, // k
  FECHA_EGRESO: 12,          // l
  CRI: 13                    // m
};

// ============================================================================
// ÍNDICES DE COLUMNAS - DATOS CCP (desde planilla externa)
// ============================================================================

const COL_CCP = {
  NOMBRE_COMPLETO: 1,        // a - "Nombre completo"
  RUT: 2,                    // b - "Rut" (formato: 12345678-9)
  SEXO: 3,                   // c - "Sexo"
  EDAD: 4,                   // d - "Edad"
  COMUNA: 5,                 // e - "Comuna"
  DIAGNOSTICO: 6,            // f - "Diagnósticos"
  CESFAM: 7,                 // g - "CESFAM"
  FECHA_IPD: 11              // k - "Fecha IPD"
};

// ============================================================================
// LISTAS DE VALIDACIÓN (Data Validation)
// ============================================================================

const LISTAS = {
  SI_NO: ['Sí', 'No'],
  
  COMUNA: ['Independencia', 'Tiltil', "Otra"],
  
  TERAPIA: [
    'Cilindro 10 m3 + Cilindro tipo E',
    'Cilindro de respaldo de 10 m3',
    'Otro'
  ],
  
  CONCENTRADOR: [
    'Concentrador 5 lpm + Cilindro tipo E',
    'Concentrador 10 lpm + Cilindro tipo E',
    'Otro'
  ],
  
  MOTIVO_EGRESO: [
    'Fallecido',
    'Traslado',
    'Otro'
  ]
};

// ============================================================================
// CONFIGURACIÓN DE EMAILS
// ============================================================================

const EMAILS = {
  PROVEEDOR: [
    //'ci-ingresosvitalaire@airliquide.com',
    //'maria.alvarado@airliquide.com'
    'fzunigac@hcuch.cl',
    'vchavez@hcuch.cl'
  ],
  CALL_CENTER: '6004264114'
};

// ============================================================================
// PLAZOS Y ALERTAS
// ============================================================================

const PLAZOS = {
  GARANTIA_VIGENTE: 3,        // días
  GARANTIA_POR_VENCER: 5,     // días
  INSTALACION_MAX: 1,         // día corrido
  RECARGA_MAX: 1,             // día corrido
  RETIRO_MAX: 3               // días hábiles
};

// ============================================================================
// COLORES PARA FORMATO CONDICIONAL
// ============================================================================

const COLORES = {
  VIGENTE: '#d9ead3',         // Verde claro
  POR_VENCER: '#fff2cc',      // Amarillo claro
  VENCIDA: '#f4cccc',         // Rojo claro
  HEADER: '#4a86e8',          // Azul
  ALERTA: '#e06666'           // Rojo
};

// ============================================================================
// MENSAJES DEL SISTEMA
// ============================================================================

const MENSAJES = {
  ERROR_NO_SELECCION: 'Por favor, seleccione al menos una fila válida.',
  ERROR_SIN_FICHA: 'El paciente seleccionado no tiene ficha de ingreso firmada.',
  ERROR_SIN_RUT: 'Por favor, ingrese el RUT del paciente primero.',
  ERROR_SIN_ALTA: 'El paciente seleccionado no tiene fecha de alta médica.',
  
  EXITO_ENVIO_INST: 'Paciente(s) enviado(s) a hoja Instalación correctamente.',
  EXITO_ENVIO_REC: 'Paciente(s) enviado(s) a hoja Recargas correctamente.',
  EXITO_ENVIO_CAJAS: 'Paciente(s) enviado(s) a hoja Cajas correctamente.',
  EXITO_ARCHIVO: 'Archivo subido correctamente.',
  EXITO_FICHA: 'Ficha de ingreso generada correctamente.',
  
  CONFIRMACION_ENVIO: '¿Está seguro que desea enviar este(os) paciente(s)?'
};

// ============================================================================
// CONFIGURACIÓN DE LOGOS (BASE64)
// ============================================================================

// Nota: Los logos completos en base64 se cargarán desde un archivo separado
// debido a su tamaño. Esta es una referencia.
const LOGOS = {
  VITALAIRE: 'data:image/png;base64,', // Se completará en FichaTemplate.html
  HCUCH: 'data:image/png;base64,'      // Se completará en FichaTemplate.html
};

// ============================================================================
// FUNCIONES DE UTILIDAD PARA CONFIGURACIÓN
// ============================================================================

/**
 * Obtiene la planilla activa
 * @returns {Spreadsheet}
 */
function getPlanillaO2() {
  return SpreadsheetApp.openById(
    PLANILLA_O2_URL.match(/\/d\/([a-zA-Z0-9-_]+)/)[1]
  );
}

/**
 * Obtiene una hoja específica por nombre
 * @param {string} nombreHoja - Nombre de la hoja
 * @returns {Sheet}
 */
function getHoja(nombreHoja) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  return ss.getSheetByName(nombreHoja);
}

/**
 * Obtiene una carpeta de Drive por su ID
 * @param {string} folderId - ID de la carpeta
 * @returns {Folder}
 */
function getCarpeta(folderId) {
  return DriveApp.getFolderById(folderId);
}

/**
 * Valida que todas las carpetas estén accesibles
 * @returns {Object} - {success: boolean, errores: Array}
 */
function validarCarpetas() {
  const errores = [];
  
  for (const [nombre, id] of Object.entries(FOLDER_IDS)) {
    try {
      DriveApp.getFolderById(id);
    } catch (e) {
      errores.push(`Error accediendo a carpeta ${nombre}: ${e.message}`);
    }
  }
  
  return {
    success: errores.length === 0,
    errores: errores
  };
}

/**
 * Valida conexión con planilla CCP
 * @returns {boolean}
 */
function validarConexionCCP() {
  try {
    const planillaId = PLANILLA_CCP_URL.match(/\/d\/([a-zA-Z0-9-_]+)/)[1];
    const planilla = SpreadsheetApp.openById(planillaId);
    const hoja = planilla.getSheetByName('Consolidado');
    return hoja !== null;
  } catch (e) {
    Logger.log('Error conectando con planilla CCP: ' + e.message);
    return false;
  }
}

// ============================================================================
// LOGGING Y DEBUG
// ============================================================================

const DEBUG_MODE = true;

/**
 * Log de depuración
 * @param {string} mensaje
 */
function debug(mensaje) {
  if (DEBUG_MODE) {
    Logger.log('[DEBUG] ' + mensaje);
  }
}

/**
 * Log de error
 * @param {string} mensaje
 * @param {Error} error
 */
function logError(mensaje, error) {
  Logger.log('[ERROR] ' + mensaje);
  if (error) {
    Logger.log('[ERROR] Stack: ' + error.stack);
  }
}
