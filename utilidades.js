/**
 * UTILIDADES.GS
 */

// ============================================================================
// UTILIDADES DE FECHAS
// ============================================================================

/**
 * Obtiene la fecha actual en formato DD-MM-YYYY
 * @returns {string}
 */
function getFechaHoy() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy');
}

/**
 * Obtiene la fecha y hora actual en formato DD-MM-YYYY HH:mm:ss
 * @returns {string}
 */
function getFechaHoraHoy() {
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd-MM-yyyy HH:mm:ss');
}

/**
 * Calcula días corridos entre dos fechas
 * @param {Date} fecha1
 * @param {Date} fecha2
 * @returns {number}
 */
function calcularDiasCorridos(fecha1, fecha2) {
  const unDia = 24 * 60 * 60 * 1000; // milisegundos en un día
  return Math.round(Math.abs((fecha1 - fecha2) / unDia));
}

/**
 * Valida si una fecha está dentro del plazo
 * @param {Date} fechaInicio
 * @param {Date} fechaActual
 * @param {number} diasLimite
 * @returns {boolean}
 */
function estaDentroPlazo(fechaInicio, fechaActual, diasLimite) {
  const dias = calcularDiasCorridos(fechaInicio, fechaActual);
  return dias <= diasLimite;
}

// ============================================================================
// UTILIDADES DE RUT
// ============================================================================

/**
 * Valida formato de RUT chileno
 * @param {string} rut - RUT sin formato (solo números)
 * @returns {boolean}
 */
function validarRUT(rut) {
  if (!rut || rut.length < 7) return false;
  
  // Validación de dígitos numéricos
  return /^\d+$/.test(rut);
}

/**
 * Calcula dígito verificador de RUT
 * @param {string} rut - RUT sin dígito verificador
 * @returns {string} - Dígito verificador (puede ser 'K')
 */
function calcularDV(rut) {
  let suma = 0;
  let multiplo = 2;
  
  // Recorrer de derecha a izquierda
  for (let i = rut.length - 1; i >= 0; i--) {
    suma += parseInt(rut.charAt(i)) * multiplo;
    multiplo = multiplo === 7 ? 2 : multiplo + 1;
  }
  
  const resto = suma % 11;
  const dv = 11 - resto;
  
  if (dv === 11) return '0';
  if (dv === 10) return 'K';
  return dv.toString();
}

/**
 * Formatea RUT con guión
 * @param {string} rut - RUT sin formato
 * @param {string} dv - Dígito verificador
 * @returns {string} - RUT formateado (12345678-9)
 */
function formatearRUT(rut, dv) {
  return `${rut}-${dv}`;
}

/**
 * Separa RUT y DV de un RUT con guión
 * @param {string} rutCompleto - RUT con formato 12345678-9
 * @returns {Object} - {rut: string, dv: string}
 */
function separarRUT(rutCompleto) {
  const partes = rutCompleto.split('-');
  return {
    rut: partes[0] || '',
    dv: partes[1] || ''
  };
}

// ============================================================================
// UTILIDADES DE NOMBRES
// ============================================================================

/**
 * Separa nombre completo en partes
 * @param {string} nombreCompleto - Nombre completo del paciente
 * @returns {Object} - {nombre: string, apellidoPat: string, apellidoMat: string}
 */
function separarNombre(nombreCompleto) {
  if (!nombreCompleto) {
    return { nombre: '', apellidoPat: '', apellidoMat: '' };
  }
  
  const partes = nombreCompleto.trim().split(/\s+/);
  
  return {
    nombre: partes[0] || '',
    apellidoPat: partes[1] || '',
    apellidoMat: partes[2] || ''
  };
}

/**
 * Concatena nombre completo desde partes
 * @param {string} nombre
 * @param {string} apellidoPat
 * @param {string} apellidoMat
 * @returns {string}
 */
function concatenarNombre(nombre, apellidoPat, apellidoMat) {
  return `${nombre} ${apellidoPat} ${apellidoMat}`.trim();
}

// ============================================================================
// UTILIDADES DE TEXTO
// ============================================================================

/**
 * Limpia y normaliza texto
 * @param {string} texto
 * @returns {string}
 */
function limpiarTexto(texto) {
  if (!texto) return '';
  return texto.toString().trim();
}

/**
 * Convierte texto a mayúsculas sin tildes
 * @param {string} texto
 * @returns {string}
 */
function normalizarTexto(texto) {
  if (!texto) return '';
  
  return texto.toString()
    .toUpperCase()
    .normalize('NFD')
    .replace(/[\u0300-\u036f]/g, '');
}

/**
 * Genera un slug (parte final y única de una URL) seguro para nombres de archivo
 * @param {string} texto
 * @returns {string}
 */
function generarSlug(texto) {
  return normalizarTexto(texto)
    .replace(/[^A-Z0-9]/g, '_')
    .replace(/_+/g, '_');
}

// ============================================================================
// UTILIDADES DE ARCHIVOS
// ============================================================================

/**
 * Genera nombre de archivo con timestamp
 * @param {string} prefijo - Prefijo del archivo (ej: 'RECETA')
 * @param {string} rut - RUT del paciente
 * @param {string} nombre - Nombre del paciente
 * @param {string} extension - Extensión del archivo (ej: 'pdf')
 * @returns {string}
 */
function generarNombreArchivo(prefijo, rut, nombre, extension = 'pdf') {
  const timestamp = Utilities.formatDate(
    new Date(),
    Session.getScriptTimeZone(),
    'ddMMyyyy_HHmmss'
  );
  
  const nombreLimpio = generarSlug(nombre);
  
  return `${prefijo}_${rut}_${nombreLimpio}_${timestamp}.${extension}`;
}

/**
 * Obtiene la extensión de un archivo desde su nombre
 * @param {string} nombreArchivo
 * @returns {string}
 */
function obtenerExtension(nombreArchivo) {
  const partes = nombreArchivo.split('.');
  return partes.length > 1 ? partes[partes.length - 1].toLowerCase() : '';
}

/**
 * Valida que un archivo sea PDF
 * @param {string} nombreArchivo
 * @returns {boolean}
 */
function esPDF(nombreArchivo) {
  return obtenerExtension(nombreArchivo) === 'pdf';
}

// ============================================================================
// UTILIDADES DE HOJAS DE CÁLCULO
// ============================================================================

/**
 * Obtiene el valor de una celda de forma segura
 * @param {Sheet} hoja
 * @param {number} fila
 * @param {number} columna
 * @returns {*}
 */
function obtenerValorCelda(hoja, fila, columna) {
  try {
    return hoja.getRange(fila, columna).getValue();
  } catch (e) {
    logError(`Error obteniendo celda [${fila},${columna}]`, e);
    return '';
  }
}

/**
 * Establece el valor de una celda de forma segura
 * @param {Sheet} hoja
 * @param {number} fila
 * @param {number} columna
 * @param {*} valor
 */
function establecerValorCelda(hoja, fila, columna, valor) {
  try {
    hoja.getRange(fila, columna).setValue(valor);
  } catch (e) {
    logError(`Error estableciendo celda [${fila},${columna}]`, e);
  }
}

/**
 * Obtiene los datos de una fila completa
 * @param {Sheet} hoja
 * @param {number} fila
 * @returns {Array}
 */
function obtenerFilaCompleta(hoja, fila) {
  try {
    const ultimaColumna = hoja.getLastColumn();
    return hoja.getRange(fila, 1, 1, ultimaColumna).getValues()[0];
  } catch (e) {
    logError(`Error obteniendo fila ${fila}`, e);
    return [];
  }
}

/**
 * Busca una fila por ID de paciente
 * @param {Sheet} hoja
 * @param {string} idPaciente
 * @param {number} columnaID - Número de columna donde está el ID
 * @returns {number} - Número de fila (0 si no se encuentra)
 */
function buscarFilaPorID(hoja, idPaciente, columnaID) {
  const datos = hoja.getRange(2, columnaID, hoja.getLastRow() - 1, 1).getValues();
  
  for (let i = 0; i < datos.length; i++) {
    if (datos[i][0] === idPaciente) {
      return i + 2; // +2 porque empezamos en fila 2 y los arrays en 0
    }
  }
  
  return 0; // No encontrado
}

/**
 * Obtiene la última fila con datos
 * @param {Sheet} hoja
 * @returns {number}
 */
function obtenerUltimaFila(hoja) {
  return hoja.getLastRow();
}

/**
 * Inserta una nueva fila al final
 * @param {Sheet} hoja
 * @param {Array} datos - Array con los valores de la fila
 * @returns {number} - Número de la fila insertada
 */
function insertarNuevaFila(hoja, datos) {
  const ultimaFila = hoja.getLastRow();
  const nuevaFila = ultimaFila + 1;
  
  hoja.getRange(nuevaFila, 1, 1, datos.length).setValues([datos]);
  
  return nuevaFila;
}

// ============================================================================
// UTILIDADES DE VALIDACIÓN
// ============================================================================

/**
 * Valida que una celda no esté vacía
 * @param {*} valor
 * @returns {boolean}
 */
function noEstaVacio(valor) {
  return valor !== null && valor !== undefined && valor !== '';
}

/**
 * Valida que una URL sea válida
 * @param {string} url
 * @returns {boolean}
 */
function esURLValida(url) {
  if (!url) return false;
  
  try {
    // Verificar que comience con http:// o https://
    return /^https?:\/\/.+/.test(url);
  } catch (e) {
    return false;
  }
}

/**
 * Valida que una fecha sea válida
 * @param {*} fecha
 * @returns {boolean}
 */
function esFechaValida(fecha) {
  return fecha instanceof Date && !isNaN(fecha);
}

/**
 * Valida que un email sea válido
 * @param {string} email
 * @returns {boolean}
 */
function esEmailValido(email) {
  if (!email) return false;
  
  const regex = /^[^\s@]+@[^\s@]+\.[^\s@]+$/;
  return regex.test(email);
}

// ============================================================================
// UTILIDADES DE UI
// ============================================================================

/**
 * Muestra un mensaje de alerta
 * @param {string} titulo
 * @param {string} mensaje
 */
function mostrarAlerta(titulo, mensaje) {
  const ui = SpreadsheetApp.getUi();
  ui.alert(titulo, mensaje, ui.ButtonSet.OK);
}

/**
 * Muestra un mensaje de confirmación
 * @param {string} titulo
 * @param {string} mensaje
 * @returns {boolean} - true si el usuario confirmó
 */
function mostrarConfirmacion(titulo, mensaje) {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.alert(titulo, mensaje, ui.ButtonSet.YES_NO);
  
  return respuesta === ui.Button.YES;
}

/**
 * Muestra un mensaje de éxito
 * @param {string} mensaje
 */
function mostrarExito(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, '✅ Éxito', 3);
}

/**
 * Muestra un mensaje de error
 * @param {string} mensaje
 */
function mostrarError(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, '❌ Error', 5);
}

/**
 * Muestra un mensaje de advertencia
 * @param {string} mensaje
 */
function mostrarAdvertencia(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, '⚠️ Advertencia', 4);
}

/**
 * Muestra un mensaje de procesando
 * @param {string} mensaje
 */
function mostrarProcesando(mensaje) {
  SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, '⏳ Procesando...', -1);
}

// ============================================================================
// UTILIDADES DE DRIVE
// ============================================================================

/**
 * Verifica si un archivo existe en una carpeta
 * @param {Folder} carpeta
 * @param {string} nombreArchivo
 * @returns {boolean}
 */
function archivoExiste(carpeta, nombreArchivo) {
  const archivos = carpeta.getFilesByName(nombreArchivo);
  return archivos.hasNext();
}

/**
 * Mueve un archivo a la papelera
 * @param {string} fileId - ID del archivo
 */
function moverAPapelera(fileId) {
  try {
    DriveApp.getFileById(fileId).setTrashed(true);
  } catch (e) {
    logError('Error moviendo archivo a papelera', e);
  }
}

/**
 * Copia un archivo a otra carpeta
 * @param {string} fileId - ID del archivo original
 * @param {string} folderId - ID de la carpeta destino
 * @returns {File}
 */
function copiarArchivo(fileId, folderId) {
  try {
    const archivo = DriveApp.getFileById(fileId);
    const carpetaDestino = DriveApp.getFolderById(folderId);
    return archivo.makeCopy(archivo.getName(), carpetaDestino);
  } catch (e) {
    logError('Error copiando archivo', e);
    return null;
  }
}

// ============================================================================
// UTILIDADES DE ID ÚNICO
// ============================================================================

/**
 * Genera ID único de paciente
 * @param {string} rut - RUT sin DV
 * @returns {string} - Formato: RUT_YYYY_CCP
 */
function generarIDPaciente(rut) {
  const anio = new Date().getFullYear();
  return `${rut}_${anio}_CCP`;
}

/**
 * Extrae el RUT de un ID de paciente
 * @param {string} idPaciente - ID en formato RUT_YYYY_CCP
 * @returns {string} - RUT sin DV
 */
function extraerRUTdeID(idPaciente) {
  if (!idPaciente) return '';
  return idPaciente.split('_')[0];
}

// ============================================================================
// UTILIDADES DE FORMATO CONDICIONAL
// ============================================================================

/**
 * Aplica color a una celda según el estado
 * @param {Sheet} hoja
 * @param {number} fila
 * @param {number} columna
 * @param {string} estado - 'vigente', 'por_vencer', 'vencida'
 */
function aplicarColorEstado(hoja, fila, columna, estado) {
  const rango = hoja.getRange(fila, columna);
  
  switch (estado.toLowerCase()) {
    case 'vigente':
      rango.setBackground(COLORES.VIGENTE);
      break;
    case 'por vencer':
      rango.setBackground(COLORES.POR_VENCER);
      break;
    case 'vencida':
      rango.setBackground(COLORES.VENCIDA);
      break;
    default:
      rango.setBackground('#ffffff');
  }
}

// ============================================================================
// UTILIDADES DE CONVERSIÓN
// ============================================================================

/**
 * Convierte un valor a booleano de forma segura
 * @param {*} valor
 * @returns {boolean}
 */
function aBooleano(valor) {
  if (typeof valor === 'boolean') return valor;
  if (typeof valor === 'string') {
    return valor.toLowerCase() === 'sí' || valor.toLowerCase() === 'si' || valor === '1';
  }
  if (typeof valor === 'number') return valor === 1;
  return false;
}

/**
 * Convierte un booleano a texto Sí/No
 * @param {boolean} valor
 * @returns {string}
 */
function aTextoSiNo(valor) {
  return valor ? 'Sí' : 'No';
}