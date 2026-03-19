/**
 * MENUS.GS
 */

/**
 * Función que se ejecuta al abrir la planilla
 * Crea el menú personalizado "Oxigenoterapia"
 */
function onOpen() {
  const ui = SpreadsheetApp.getUi();
  
  ui.createMenu('💻 Oxigenoterapia')
    // Sección: Documentos
    .addSubMenu(ui.createMenu('📄 Documentos')
      .addItem('📋 Generar Ficha de Ingreso', 'generarFichaIngreso')
      .addItem('📎 Subir Receta Médica', 'subirRecetaMedica')
      .addItem('✍️ Subir Ficha Firmada', 'subirFichaFirmada')
      .addItem('📄 Subir Certificado Defunción', 'subirCertificadoDefuncion')
      .addItem('📦 Subir Documento Retiro', 'subirDocumentoRetiro')
      .addItem('✅ Subir Comprobante Entrega', 'subirComprobanteEntrega')
      .addItem('📄 Subir Orden Entrega', 'subirOrdenEntrega'))
    
    .addSeparator()
    
    // Sección: Envío de datos
    .addSubMenu(ui.createMenu('🔄 Enviar Pacientes')
      .addItem('🏗️ Enviar a Instalación', 'enviarAInstalacion')
      .addItem('🔄 Enviar a Recargas', 'enviarARecargas')
      .addItem('💰 Enviar a Cajas', 'enviarACajas'))
    
    .addSeparator()
    
    // Sección: Utilidades
    .addSubMenu(ui.createMenu('🔧 Utilidades')
      .addItem('🔄 Sincronizar Datos CCP', 'sincronizarDatosCCP')
      .addItem('✅ Validar Configuración', 'validarConfiguracion')
      .addItem('📊 Ver Estadísticas', 'mostrarEstadisticas'))
    
    .addSeparator()
    
    // Ayuda
    .addItem('❓ Ayuda', 'mostrarAyuda')
    
    .addToUi();
  
  debug('Menú personalizado creado exitosamente');
}

// ============================================================================
// FUNCIONES DE VALIDACIÓN PREVIAS
// ============================================================================

/**
 * Obtiene la hoja activa y valida que sea correcta
 * @param {Array<string>} hojasPermitidas - Nombres de hojas permitidas
 * @returns {Sheet|null}
 */
function validarHojaActiva(hojasPermitidas) {
  const hojaActiva = SpreadsheetApp.getActiveSheet();
  const nombreHoja = hojaActiva.getName();
  
  if (!hojasPermitidas.includes(nombreHoja)) {
    mostrarError(`Esta acción solo está disponible en las hojas: ${hojasPermitidas.join(', ')}`);
    return null;
  }
  
  return hojaActiva;
}

/**
 * Obtiene la fila activa y valida que no sea el encabezado
 * @returns {number|null}
 */
function validarFilaActiva() {
  const filaActiva = SpreadsheetApp.getActiveRange().getRow();
  
  if (filaActiva === 1) {
    mostrarError('Por favor, seleccione una fila de datos (no el encabezado)');
    return null;
  }
  
  return filaActiva;
}

/**
 * Obtiene las filas seleccionadas
 * @returns {Array<number>}
 */
function obtenerFilasSeleccionadas() {
  const rango = SpreadsheetApp.getActiveRange();
  const primeraFila = rango.getRow();
  const numFilas = rango.getNumRows();
  
  const filas = [];
  for (let i = 0; i < numFilas; i++) {
    const fila = primeraFila + i;
    if (fila > 1) { // Excluir encabezado
      filas.push(fila);
    }
  }
  
  return filas;
}

// ============================================================================
// FUNCIONES DE SUBIDA DE DOCUMENTOS
// ============================================================================

/**
 * Subir Receta Médica
 */
function subirRecetaMedica() {
  const hoja = validarHojaActiva([HOJAS.HISTORIAL]);
  if (!hoja) return;
  
  const fila = validarFilaActiva();
  if (!fila) return;
  
  // Validar que haya RUT ingresado
  const rut = obtenerValorCelda(hoja, fila, COL_HISTORIAL.RUT);
  if (!rut) {
    mostrarError(MENSAJES.ERROR_SIN_RUT);
    return;
  }
  
  const nombre = obtenerValorCelda(hoja, fila, COL_HISTORIAL.NOMBRE);
  const apellido = obtenerValorCelda(hoja, fila, COL_HISTORIAL.APELLIDO_PAT);
  
  // Abrir diálogo de carga
  const html = HtmlService.createHtmlOutputFromFile('FileUploader')
    .setWidth(500)
    .setHeight(350)
    .setSandboxMode(HtmlService.SandboxMode.IFRAME);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '📎 Subir Receta Médica');
  
  // Guardar contexto en propiedades temporales
  const props = PropertiesService.getUserProperties();
  props.setProperty('UPLOAD_TIPO', 'receta');
  props.setProperty('UPLOAD_FILA', fila.toString());
  props.setProperty('UPLOAD_RUT', rut);
  props.setProperty('UPLOAD_NOMBRE', `${nombre}_${apellido}`);
}

/**
 * Subir Ficha Firmada
 */
function subirFichaFirmada() {
  const hoja = validarHojaActiva([HOJAS.HISTORIAL]);
  if (!hoja) return;
  
  const fila = validarFilaActiva();
  if (!fila) return;
  
  const rut = obtenerValorCelda(hoja, fila, COL_HISTORIAL.RUT);
  if (!rut) {
    mostrarError(MENSAJES.ERROR_SIN_RUT);
    return;
  }
  
  const nombre = obtenerValorCelda(hoja, fila, COL_HISTORIAL.NOMBRE);
  const apellido = obtenerValorCelda(hoja, fila, COL_HISTORIAL.APELLIDO_PAT);
  
  const html = HtmlService.createHtmlOutputFromFile('FileUploader')
    .setWidth(500)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '✍️ Subir Ficha de Ingreso Firmada');
  
  const props = PropertiesService.getUserProperties();
  props.setProperty('UPLOAD_TIPO', 'ficha_firmada');
  props.setProperty('UPLOAD_FILA', fila.toString());
  props.setProperty('UPLOAD_RUT', rut);
  props.setProperty('UPLOAD_NOMBRE', `${nombre}_${apellido}`);
}

/**
 * Subir Certificado de Defunción
 */
function subirCertificadoDefuncion() {
  const hoja = validarHojaActiva([HOJAS.HISTORIAL]);
  if (!hoja) return;
  
  const fila = validarFilaActiva();
  if (!fila) return;
  
  const rut = obtenerValorCelda(hoja, fila, COL_HISTORIAL.RUT);
  if (!rut) {
    mostrarError(MENSAJES.ERROR_SIN_RUT);
    return;
  }
  
  // Validar que motivo egreso sea "Fallecido"
  const motivoEgreso = obtenerValorCelda(hoja, fila, COL_HISTORIAL.MOTIVO_EGRESO);
  if (motivoEgreso !== 'Fallecido') {
    if (!mostrarConfirmacion('Advertencia', 
        'El motivo de egreso no es "Fallecido". ¿Desea continuar de todas formas?')) {
      return;
    }
  }
  
  const nombre = obtenerValorCelda(hoja, fila, COL_HISTORIAL.NOMBRE);
  const apellido = obtenerValorCelda(hoja, fila, COL_HISTORIAL.APELLIDO_PAT);
  
  const html = HtmlService.createHtmlOutputFromFile('FileUploader')
    .setWidth(500)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '📄 Subir Certificado de Defunción');
  
  const props = PropertiesService.getUserProperties();
  props.setProperty('UPLOAD_TIPO', 'certificado');
  props.setProperty('UPLOAD_FILA', fila.toString());
  props.setProperty('UPLOAD_RUT', rut);
  props.setProperty('UPLOAD_NOMBRE', `${nombre}_${apellido}`);
}

/**
 * Subir Documento de Retiro
 */
function subirDocumentoRetiro() {
  const hoja = validarHojaActiva([HOJAS.HISTORIAL]);
  if (!hoja) return;
  
  const fila = validarFilaActiva();
  if (!fila) return;
  
  const rut = obtenerValorCelda(hoja, fila, COL_HISTORIAL.RUT);
  if (!rut) {
    mostrarError(MENSAJES.ERROR_SIN_RUT);
    return;
  }
  
  const nombre = obtenerValorCelda(hoja, fila, COL_HISTORIAL.NOMBRE);
  const apellido = obtenerValorCelda(hoja, fila, COL_HISTORIAL.APELLIDO_PAT);
  
  const html = HtmlService.createHtmlOutputFromFile('FileUploader')
    .setWidth(500)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '📦 Subir Documento de Retiro');
  
  const props = PropertiesService.getUserProperties();
  props.setProperty('UPLOAD_TIPO', 'retiro');
  props.setProperty('UPLOAD_FILA', fila.toString());
  props.setProperty('UPLOAD_RUT', rut);
  props.setProperty('UPLOAD_NOMBRE', `${nombre}_${apellido}`);
}

/**
 * Subir Comprobante de Entrega (desde hoja Instalación)
 */
function subirComprobanteEntrega() {
  const hoja = validarHojaActiva([HOJAS.INSTALACION]);
  if (!hoja) return;
  
  const fila = validarFilaActiva();
  if (!fila) return;
  
  const rut = obtenerValorCelda(hoja, fila, COL_INSTALACION.RUT);
  if (!rut) {
    mostrarError(MENSAJES.ERROR_SIN_RUT);
    return;
  }
  
  const nombre = obtenerValorCelda(hoja, fila, COL_INSTALACION.NOMBRE);
  const apellido = obtenerValorCelda(hoja, fila, COL_INSTALACION.APELLIDO_PAT);
  
  const html = HtmlService.createHtmlOutputFromFile('FileUploader')
    .setWidth(500)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '✅ Subir Comprobante de Entrega');
  
  const props = PropertiesService.getUserProperties();
  props.setProperty('UPLOAD_TIPO', 'comprobante');
  props.setProperty('UPLOAD_FILA', fila.toString());
  props.setProperty('UPLOAD_RUT', rut);
  props.setProperty('UPLOAD_NOMBRE', `${nombre}_${apellido}`);
  props.setProperty('UPLOAD_HOJA', HOJAS.INSTALACION);
}

/**
 * Subir Orden de Entrega (desde hoja Recargas)
 */
function subirOrdenEntrega() {
  const hoja = validarHojaActiva([HOJAS.RECARGAS]);
  if (!hoja) return;
  
  const fila = validarFilaActiva();
  if (!fila) return;
  
  const rut = obtenerValorCelda(hoja, fila, COL_RECARGAS.RUT);
  if (!rut) {
    mostrarError(MENSAJES.ERROR_SIN_RUT);
    return;
  }
  
  const nombre = obtenerValorCelda(hoja, fila, COL_RECARGAS.NOMBRE);
  const apellido = obtenerValorCelda(hoja, fila, COL_RECARGAS.APELLIDO_PAT);
  
  const html = HtmlService.createHtmlOutputFromFile('FileUploader')
    .setWidth(500)
    .setHeight(350);
  
  SpreadsheetApp.getUi().showModelessDialog(html, '📄 Subir Orden de Entrega');
  
  const props = PropertiesService.getUserProperties();
  props.setProperty('UPLOAD_TIPO', 'orden');
  props.setProperty('UPLOAD_FILA', fila.toString());
  props.setProperty('UPLOAD_RUT', rut);
  props.setProperty('UPLOAD_NOMBRE', `${nombre}_${apellido}`);
  props.setProperty('UPLOAD_HOJA', HOJAS.RECARGAS);
}

// ============================================================================
// FUNCIÓN DE AYUDA
// ============================================================================

/**
 * Muestra diálogo de ayuda
 */
function mostrarAyuda() {
  const html = `
    <div style="font-family: Arial; padding: 20px; line-height: 1.6;">
      <h2>📖 Sistema de Oxigenoterapia HCUCH</h2>
      
      <h3>🔄 Flujo de Trabajo:</h3>
      <ol>
        <li><strong>Ingreso:</strong> Completar datos del paciente en "Historial Pacientes"</li>
        <li><strong>Documentos:</strong> Generar y subir ficha de ingreso</li>
        <li><strong>Instalación:</strong> Enviar a hoja "Instalación" cuando ficha esté firmada</li>
        <li><strong>Recargas:</strong> Marcar si requiere recarga y enviar a hoja correspondiente</li>
        <li><strong>Cierre:</strong> Registrar alta médica y enviar a "Cajas"</li>
      </ol>
      
      <h3>📋 Botones Principales:</h3>
      <ul>
        <li><strong>Generar Ficha:</strong> Crea PDF con datos del paciente</li>
        <li><strong>Subir Documentos:</strong> Adjunta archivos y genera enlaces</li>
        <li><strong>Enviar a...</strong>: Mueve pacientes entre hojas del flujo</li>
      </ul>
      
      <h3>⚠️ Validaciones Automáticas:</h3>
      <ul>
        <li>Garantía GES-4: máximo 5 días</li>
        <li>Instalación: máximo 1 día corrido</li>
        <li>Recargas: máximo 1 día corrido</li>
      </ul>
      <br>
      <button onclick="google.script.host.close()">Cerrar</button>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(600)
    .setHeight(500);
  
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, '❓ Ayuda - Sistema Oxigenoterapia');
}

/**
 * Valida la configuración del sistema
 */
function validarConfiguracion() {
  mostrarProcesando('Validando configuración...');
  
  const errores = [];
  
  // Validar carpetas
  const resultadoCarpetas = validarCarpetas();
  if (!resultadoCarpetas.success) {
    errores.push(...resultadoCarpetas.errores);
  }
  
  // Validar conexión CCP
  if (!validarConexionCCP()) {
    errores.push('No se puede conectar con la planilla "Excel Cuidados Paliativos 2026"');
  }
  
  // Validar hojas existen
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  for (const [nombre, nombreHoja] of Object.entries(HOJAS)) {
    if (!ss.getSheetByName(nombreHoja)) {
      errores.push(`Hoja "${nombreHoja}" no encontrada`);
    }
  }
  
  // Mostrar resultados
  if (errores.length === 0) {
    mostrarExito('✅ Configuración validada correctamente. Todo funciona bien.');
  } else {
    const mensaje = '❌ Errores encontrados:\n\n' + errores.join('\n');
    mostrarAlerta('Errores de Configuración', mensaje);
  }
}

/**
 * Muestra estadísticas básicas
 */
function mostrarEstadisticas() {
  const hojaHistorial = getHoja(HOJAS.HISTORIAL);
  const hojaInstalacion = getHoja(HOJAS.INSTALACION);
  const hojaRecargas = getHoja(HOJAS.RECARGAS);
  const hojaCajas = getHoja(HOJAS.CAJAS);
  
  const totalPacientes = hojaHistorial.getLastRow() - 1;
  const totalInstalaciones = hojaInstalacion.getLastRow() - 1;
  const totalRecargas = hojaRecargas.getLastRow() - 1;
  const totalFacturacion = hojaCajas.getLastRow() - 1;
  
  // Contar pacientes vigentes
  let pacientesVigentes = 0;
  const datosVigencia = hojaHistorial.getRange(2, COL_HISTORIAL.PACIENTE_VIGENTE, 
                                                totalPacientes, 1).getValues();
  for (let i = 0; i < datosVigencia.length; i++) {
    if (datosVigencia[i][0] === 'Sí') {
      pacientesVigentes++;
    }
  }
  
  const html = `
    <div style="font-family: Arial; padding: 20px;">
      <h2>📊 Estadísticas del Sistema</h2>
      
      <div style="background: #e8f4f8; padding: 15px; border-radius: 8px; margin: 10px 0;">
        <h3>👥 Pacientes</h3>
        <p><strong>Total registrados:</strong> ${totalPacientes}</p>
        <p><strong>Vigentes:</strong> ${pacientesVigentes}</p>
        <p><strong>Egresados:</strong> ${totalPacientes - pacientesVigentes}</p>
      </div>
      
      <div style="background: #fff4e6; padding: 15px; border-radius: 8px; margin: 10px 0;">
        <h3>🏗️ Instalaciones</h3>
        <p><strong>Total:</strong> ${totalInstalaciones}</p>
      </div>
      
      <div style="background: #e8f5e9; padding: 15px; border-radius: 8px; margin: 10px 0;">
        <h3>🔄 Recargas</h3>
        <p><strong>Total:</strong> ${totalRecargas}</p>
      </div>
      
      <div style="background: #f3e5f5; padding: 15px; border-radius: 8px; margin: 10px 0;">
        <h3>💰 Facturación</h3>
        <p><strong>Total casos:</strong> ${totalFacturacion}</p>
      </div>
      
      <p style="margin-top: 20px; color: #666;">
        <small>Actualizado: ${getFechaHoraHoy()}</small>
      </p>
      
      <button onclick="google.script.host.close()">Cerrar</button>
    </div>
  `;
  
  const htmlOutput = HtmlService.createHtmlOutput(html)
    .setWidth(500)
    .setHeight(600);
  
  SpreadsheetApp.getUi().showModelessDialog(htmlOutput, '📊 Estadísticas');
}