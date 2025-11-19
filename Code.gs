// ====================================
// SISTEMA DE SEGUIMIENTO DE PRÁCTICAS v2.8 - ESTRUCTURA ACTUALIZADA
// Nuevas columnas de etapas: Aliados, Plataformas, Conexión laboral, Por su cuenta, No busca trabajar, Empleado, No terminó la formación
// FIX: Sistema de bloqueo para evitar doble procesamiento
// ====================================

const LLAMADAS_PARA_FINALIZAR = 6;

// MAPEO DE COLUMNAS v2.8
const COLUMNAS = {
  ID: 1,
  NOMBRE: 2,
  TELEFONO: 3,
  FORMACION: 4,
  // NUEVAS COLUMNAS DE ETAPAS (5-11)
  ALIADOS: 5,
  PLATAFORMAS: 6,
  CONEXION_LABORAL: 7,
  POR_SU_CUENTA: 8,
  NO_BUSCA_TRABAJAR: 9,
  EMPLEADO: 10,
  NO_TERMINO_FORMACION: 11,
  // COLUMNAS ORIGINALES (ahora 12-18)
  ETAPA_ACTUAL: 12,
  RESULTADOS: 13,
  DOCUMENTOS: 14,
  FECHA: 15,
  NOTAS: 16,
  TOTAL_LLAMADAS: 17,
  PROCESAR: 18
};

const TOTAL_COLUMNAS = 18;
const TOTAL_COLUMNAS_DESTINO = 17; // Hojas de llamadas no tienen checkbox

function onOpen() {
  const ui = SpreadsheetApp.getUi();
  ui.createMenu('🎓 Sistema de Seguimiento')
    .addItem('📖 Ver Instrucciones', 'mostrarHojaInstrucciones')
    .addSeparator()
    .addSubMenu(ui.createMenu('🔧 Configuración')
      .addItem('⚙️ Configuración Inicial', 'configurarHojasCorregido')
      .addItem('🔄 ACTUALIZAR Estructura de Tabla', 'actualizarEstructuraTabla')
      .addItem('🔗 Configurar Google Sheet Externo', 'configurarSheetExterno')
      .addItem('🎨 Aplicar Diseño Profesional', 'aplicarFormato')
      .addItem('⚙️ Ver Configuración Actual', 'verConfiguracion'))
    .addSeparator()
    .addSubMenu(ui.createMenu('📥 Importación')
      .addItem('📥 Importar Datos 2024', 'importarDatos2024')
      .addItem('📥 Importar Datos 2025', 'importarDatos2025')
      .addItem('📥 Importar Desde Otra Hoja', 'importarDatosPersonalizados')
      .addItem('🧪 Probar Conexión Externa', 'probarConexionSheetExterno')
      .addItem('🎯 Finalizar Participantes por Etapa', 'finalizarParticipantesPorEtapa'))
    .addSeparator()
    .addItem('⚡ ACTIVAR Procesamiento Automático', 'activarProcesomientoAutomatico')
    .addItem('🔴 DESACTIVAR Procesamiento Automático', 'desactivarProcesomientoAutomatico')
    .addItem('🔄 Procesar Llamadas Marcadas', 'procesarLlamadasManualesCorregido')
    .addItem('🧪 Procesar Una Llamada (Prueba)', 'procesarUnaLlamadaPrueba')
    .addSeparator()
    .addSubMenu(ui.createMenu('📊 Reportes y Análisis')
      .addItem('📈 Generar Reporte Completo', 'generarReporte')
      .addItem('🔍 Diagnosticar Sistema', 'diagnosticarSistema')
      .addItem('🔍 Diagnosticar Doble Procesamiento', 'diagnosticarDobleProcesamiento')
      .addItem('🔍 Verificar Checkboxes', 'verificarCheckboxes'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🛠️ Mantenimiento')
      .addItem('🔧 Reparar Sistema', 'repararSistema')
      .addItem('🧹 Limpiar Datos Vacíos', 'limpiarDatosVacios')
      .addItem('🧹 Limpiar Bloqueos de Procesamiento', 'limpiarBloqueosProcesamiento')
      .addItem('💾 Crear Respaldo', 'crearRespaldo')
      .addItem('🔄 Resetear Sistema Completo', 'resetearSistema'))
    .addSeparator()
    .addSubMenu(ui.createMenu('🎯 Gestión de Etapas')
      .addItem('🔧 Configurar Desplegable Etapas', 'configurarDesplegableEtapaActual')
      .addItem('🔄 Actualizar Etapas por Llamadas', 'actualizarEtapaSegunLlamadas')
      .addItem('📊 Reporte de Etapas', 'generarReporteEtapas')
      .addItem('🔍 Buscar por Etapa', 'buscarPorEtapa'))
    .addToUi();

  verificarEstadoProcesomientoAutomatico();
}

// ====================================
// FUNCIÓN NUEVA: ACTUALIZAR ESTRUCTURA DE TABLA SIN BORRAR DATOS
// ====================================
function actualizarEstructuraTabla() {
  const ui = SpreadsheetApp.getUi();

  const confirmacion = ui.alert(
    '🔄 ACTUALIZAR Estructura de Tabla v2.8',
    '⚠️ Esta función agregará 7 columnas nuevas de etapas:\n\n' +
    '• Aliados\n' +
    '• Plataformas\n' +
    '• Conexión laboral\n' +
    '• Por su cuenta\n' +
    '• No busca trabajar\n' +
    '• Empleado\n' +
    '• No terminó la formación\n\n' +
    '✅ SIN BORRAR datos existentes\n\n' +
    '¿Continuar?',
    ui.ButtonSet.YES_NO
  );

  if (confirmacion !== ui.Button.YES) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojasAProcesar = [
      '📋 Seguimiento General',
      '📞 Llamada 1',
      '📞 Llamada 2',
      '📞 Llamada 3',
      '📞 Llamada 4',
      '📞 Llamada 5',
      '✅ Finalizados'
    ];

    let hojasActualizadas = 0;
    const resultados = [];

    hojasAProcesar.forEach(nombreHoja => {
      const hoja = ss.getSheetByName(nombreHoja);
      if (!hoja) {
        resultados.push(`❌ ${nombreHoja}: No encontrada`);
        return;
      }

      // Verificar si ya tiene las columnas nuevas
      const encabezados = hoja.getRange(1, 1, 1, hoja.getLastColumn()).getValues()[0];
      if (encabezados[4] === 'Aliados' || encabezados.length >= 18) {
        resultados.push(`ℹ️ ${nombreHoja}: Ya actualizada`);
        return;
      }

      // Insertar 7 columnas después de la columna 4 (Formación)
      hoja.insertColumnsAfter(4, 7);

      // Agregar encabezados de las nuevas columnas
      const nuevosEncabezados = [
        'Aliados',
        'Plataformas',
        'Conexión laboral',
        'Por su cuenta',
        'No busca trabajar',
        'Empleado',
        'No terminó la formación'
      ];

      hoja.getRange(1, 5, 1, 7).setValues([nuevosEncabezados]);

      // Aplicar formato a los encabezados nuevos
      const rangoNuevosEncabezados = hoja.getRange(1, 5, 1, 7);
      rangoNuevosEncabezados.setBackground('#1f4e79');
      rangoNuevosEncabezados.setFontColor('#ffffff');
      rangoNuevosEncabezados.setFontWeight('bold');
      rangoNuevosEncabezados.setFontSize(11);
      rangoNuevosEncabezados.setHorizontalAlignment('center');

      // Ajustar anchos de columnas
      for (let col = 5; col <= 11; col++) {
        hoja.setColumnWidth(col, 150);
      }

      hojasActualizadas++;
      resultados.push(`✅ ${nombreHoja}: Actualizada exitosamente`);
    });

    // Reconfigurar validaciones y checkboxes con nuevas posiciones
    reconfigurarValidacionesActualizadas();

    const mensaje = `🔄 ACTUALIZACIÓN COMPLETADA v2.8\n\n` +
      `📊 Hojas procesadas: ${hojasActualizadas}/${hojasAProcesar.length}\n\n` +
      `📋 DETALLES:\n${resultados.join('\n')}\n\n` +
      `✅ Estructura actualizada sin pérdida de datos`;

    ui.alert('✅ Actualización Exitosa', mensaje, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Actualización',
      `Error actualizando estructura: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function reconfigurarValidacionesActualizadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');

  if (!hojaGeneral) return;

  const ultimaFila = Math.max(hojaGeneral.getLastRow(), 100);

  try {
    // FORMACIÓN - columna 4 (sin cambios)
    const opcionesFormacion = [
      'Barista I', 'Barista II', 'Barista III', 'Barista IV',
      'Barismo', 'Barismo 1', 'Barismo 2', 'Barismo 3', 'Barismo 4', 'Barismo 5',
      'Barismo 6', 'Barismo 7', 'Barismo 8', 'Barismo 9', 'Barismo 10',
      'Gastronomía', 'Gastronomía I', 'Gastronomía II', 'Gastronomía III',
      'Gastronomía IV', 'Gastronomía V',
      'Gastronomía 1', 'Gastronomía 2', 'Gastronomía 3', 'Gastronomía 4', 'Gastronomía 5',
      'Panadería', 'Repostería', 'Sommelier', 'Food Manager',
      'Análisis de datos E-commerce', 'SAC', 'Ofimática',
      'Otra'
    ];

    const rangoFormacion = hojaGeneral.getRange(2, COLUMNAS.FORMACION, ultimaFila - 1, 1);
    const validacionFormacion = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesFormacion)
      .setAllowInvalid(true)
      .setHelpText('Selecciona la formación\n🆕 Nuevas: Análisis de datos, SAC, Ofimática')
      .build();
    rangoFormacion.setDataValidation(validacionFormacion);

    // ETAPA ACTUAL - ahora columna 12
    const opcionesEtapa = [
      'Inicial', 'Aliados', 'Plataformas', 'Conexión Laboral',
      'Por su cuenta', 'No busca trabajar', 'Empleado',
      'No termino la formación', 'Finalizado'
    ];

    const rangoEtapa = hojaGeneral.getRange(2, COLUMNAS.ETAPA_ACTUAL, ultimaFila - 1, 1);
    const validacionEtapa = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesEtapa)
      .setAllowInvalid(true)
      .build();
    rangoEtapa.setDataValidation(validacionEtapa);

    // DOCUMENTOS FALTANTES - ahora columna 14
    const opcionesDocumentos = [
      'Salud', 'Manipulación', 'Pulmones', 'Policiacos',
      'Penales', 'Ninguno', 'CV', 'NIT', 'Fotografía', 'DPI'
    ];

    const rangoDocumentos = hojaGeneral.getRange(2, COLUMNAS.DOCUMENTOS, ultimaFila - 1, 1);
    const validacionDocumentos = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesDocumentos)
      .setAllowInvalid(true)
      .setHelpText('📄 SELECCIÓN MÚLTIPLE:\n1. Selecciona documento\n2. Se agrega automáticamente\n3. Para remover: selecciona el mismo\n4. "Ninguno" borra todo')
      .build();
    rangoDocumentos.setDataValidation(validacionDocumentos);

    // CHECKBOXES - ahora columna 18
    const rangoCheckbox = hojaGeneral.getRange(2, COLUMNAS.PROCESAR, ultimaFila - 1, 1);
    rangoCheckbox.insertCheckboxes();

    console.log('✅ Validaciones reconfiguradas con nueva estructura');

  } catch (error) {
    console.error('Error reconfigurando validaciones:', error);
  }
}

// ====================================
// FUNCIÓN onEdit ACTUALIZADA - v2.8 CON NUEVAS COLUMNAS
// ====================================
function onEdit(e) {
  // Validaciones iniciales
  if (!e || !e.range || !e.source) {
    return;
  }

  const hoja = e.source.getActiveSheet();
  const rango = e.range;
  const fila = rango.getRow();
  const columna = rango.getColumn();

  // Verificar que estamos en la hoja correcta
  if (hoja.getName() !== '📋 Seguimiento General' || fila <= 1) {
    return;
  }

  // *** DOCUMENTOS FALTANTES (columna 14) ***
  if (columna === COLUMNAS.DOCUMENTOS && rango.getNumRows() === 1 && rango.getNumColumns() === 1) {
    console.log(`Cambio detectado en Documentos Faltantes - Fila: ${fila}`);
    procesarSeleccionDocumentos(hoja, fila, e.value, e.oldValue);
    return;
  }

  // *** ETAPA ACTUAL (columna 12) ***
  if (columna === COLUMNAS.ETAPA_ACTUAL && rango.getNumRows() === 1 && rango.getNumColumns() === 1) {
    const nuevaEtapa = rango.getValue();
    if (nuevaEtapa && nuevaEtapa.toString().toLowerCase().includes('finalizado')) {
      Utilities.sleep(1000);
      procesarFinalizacionPorEtapa(hoja, fila);
      return;
    }
  }

  // Verificar procesamiento automático
  const propiedades = PropertiesService.getScriptProperties();
  const procesamientoActivo = propiedades.getProperty('PROCESAMIENTO_AUTOMATICO') === 'true';

  if (!procesamientoActivo) {
    return;
  }

  // *** PROCESAMIENTO DE CHECKBOX (columna 18) ***
  if (columna !== COLUMNAS.PROCESAR) {
    return;
  }

  if (rango.getNumRows() !== 1 || rango.getNumColumns() === 1) {
    return;
  }

  const nuevoValor = rango.getValue();
  if (typeof nuevoValor !== 'boolean' || nuevoValor !== true) {
    return;
  }

  // *** SISTEMA DE BLOQUEO PARA EVITAR DOBLE PROCESAMIENTO ***
  const lock = LockService.getScriptLock();

  try {
    // Verificar si ya está procesando
    const marcaProcesamiento = propiedades.getProperty(`proc_${fila}`);

    if (marcaProcesamiento) {
      const tiempoTranscurrido = new Date().getTime() - parseInt(marcaProcesamiento);
      if (tiempoTranscurrido < 10000) { // Menos de 10 segundos
        console.log(`⚠️ Fila ${fila} ya está siendo procesada. Ignorando.`);
        return;
      }
    }

    // Intentar adquirir bloqueo
    if (!lock.tryLock(5000)) {
      console.log('⚠️ No se pudo obtener bloqueo. Saliendo.');
      return;
    }

    // Marcar inicio de procesamiento
    propiedades.setProperty(`proc_${fila}`, new Date().getTime().toString());

    // Pequeña pausa para asegurar
    Utilities.sleep(500);

    const nombre = hoja.getRange(fila, COLUMNAS.NOMBRE).getValue();
    console.log(`🔄 Procesando fila ${fila}: ${nombre}`);

    const resultado = procesarLlamadaOptimizada(hoja, fila);

    if (resultado.exito) {
      let mensaje;
      if (resultado.movido) {
        mensaje = `🎯 ${nombre} completó ${LLAMADAS_PARA_FINALIZAR} llamadas y fue movido a Finalizados`;
      } else {
        mensaje = `📞 ${nombre} procesado correctamente (${resultado.totalLlamadas}/${LLAMADAS_PARA_FINALIZAR} llamadas)`;
      }
      mostrarNotificacionDiscreta(mensaje);
      console.log(`✅ Procesamiento exitoso: ${nombre}`);
    } else {
      throw new Error(resultado.error);
    }

  } catch (error) {
    console.error('❌ Error en onEdit:', error);

    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error en procesamiento: ${error.message}`,
      'Error del Sistema',
      5
    );

    try {
      if (e && e.range) {
        e.range.setValue(false);
      }
    } catch (cleanupError) {
      console.error('Error limpiando checkbox:', cleanupError);
    }

  } finally {
    // Limpiar marca de procesamiento
    try {
      propiedades.deleteProperty(`proc_${fila}`);
      lock.releaseLock();
      console.log(`🔓 Bloqueo liberado para fila ${fila}`);
    } catch (releaseError) {
      console.error('Error liberando bloqueo:', releaseError);
    }
  }
}

// ====================================
// LIMPIEZA DE BLOQUEOS
// ====================================
function limpiarBloqueosProcesamiento() {
  const ui = SpreadsheetApp.getUi();

  try {
    const propiedades = PropertiesService.getScriptProperties();
    const todasLasPropiedades = propiedades.getProperties();

    let bloqueosLimpiados = 0;

    Object.keys(todasLasPropiedades).forEach(clave => {
      if (clave.startsWith('proc_') || clave.startsWith('procesando_fila_')) {
        propiedades.deleteProperty(clave);
        bloqueosLimpiados++;
      }
    });

    if (bloqueosLimpiados > 0) {
      ui.alert('🧹 Bloqueos Limpiados',
        `Se limpiaron ${bloqueosLimpiados} bloqueos de procesamiento.\n\n✅ Sistema listo para usar.`,
        ui.ButtonSet.OK);
    } else {
      ui.alert('✅ Sin Bloqueos',
        'No se encontraron bloqueos para limpiar.\n\nEl sistema está funcionando correctamente.',
        ui.ButtonSet.OK);
    }

    console.log(`✅ ${bloqueosLimpiados} bloqueos limpiados`);

  } catch (error) {
    ui.alert('❌ Error',
      `Error limpiando bloqueos: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

// ====================================
// DIAGNÓSTICO DE DOBLE PROCESAMIENTO
// ====================================
function diagnosticarDobleProcesamiento() {
  const ui = SpreadsheetApp.getUi();
  const propiedades = PropertiesService.getScriptProperties();

  let diagnostico = '🔍 DIAGNÓSTICO DE DOBLE PROCESAMIENTO v2.8\n\n';

  // 1. Verificar triggers
  const triggers = ScriptApp.getProjectTriggers();
  const triggersOnEdit = triggers.filter(t =>
    t.getHandlerFunction() === 'onEdit' &&
    t.getEventType() === ScriptApp.EventType.ON_EDIT
  );

  diagnostico += `📌 TRIGGERS onEdit:\n`;
  diagnostico += `• Cantidad: ${triggersOnEdit.length}\n`;

  if (triggersOnEdit.length > 1) {
    diagnostico += `❌ PROBLEMA ENCONTRADO: ${triggersOnEdit.length} triggers (debería ser 1)\n`;
    diagnostico += `💡 Solución: Desactivar y reactivar procesamiento automático\n\n`;
  } else if (triggersOnEdit.length === 1) {
    diagnostico += `✅ Correcto: Solo 1 trigger onEdit\n\n`;
  } else {
    diagnostico += `ℹ️ Sin triggers (modo manual)\n\n`;
  }

  // 2. Verificar bloqueos activos
  const todasLasPropiedades = propiedades.getProperties();
  const bloqueosActivos = Object.keys(todasLasPropiedades).filter(k =>
    k.startsWith('proc_') || k.startsWith('procesando_fila_')
  );

  diagnostico += `🔒 BLOQUEOS ACTIVOS:\n`;
  diagnostico += `• Cantidad: ${bloqueosActivos.length}\n`;

  if (bloqueosActivos.length > 0) {
    diagnostico += `⚠️ Bloqueos encontrados:\n`;
    bloqueosActivos.slice(0, 5).forEach(bloqueo => {
      diagnostico += `  - ${bloqueo}\n`;
    });
    if (bloqueosActivos.length > 5) {
      diagnostico += `  ... y ${bloqueosActivos.length - 5} más\n`;
    }
    diagnostico += `💡 Usar: Mantenimiento → Limpiar Bloqueos\n\n`;
  } else {
    diagnostico += `✅ No hay bloqueos activos\n\n`;
  }

  // 3. Estado del procesamiento
  const procesamientoActivo = propiedades.getProperty('PROCESAMIENTO_AUTOMATICO') === 'true';
  diagnostico += `⚡ PROCESAMIENTO AUTOMÁTICO:\n`;
  diagnostico += `• Estado: ${procesamientoActivo ? '✅ ACTIVO' : '📋 INACTIVO'}\n\n`;

  // 4. Recomendaciones
  diagnostico += `💡 RECOMENDACIONES:\n`;

  if (triggersOnEdit.length > 1) {
    diagnostico += `1️⃣ URGENTE: Eliminar triggers duplicados\n`;
    diagnostico += `   → Desactivar procesamiento automático\n`;
    diagnostico += `   → Esperar 5 segundos\n`;
    diagnostico += `   → Activar procesamiento automático\n\n`;
  }

  if (bloqueosActivos.length > 0) {
    diagnostico += `2️⃣ Limpiar bloqueos atorados\n`;
    diagnostico += `   → Mantenimiento → Limpiar Bloqueos\n\n`;
  }

  if (triggersOnEdit.length === 1 && bloqueosActivos.length === 0) {
    diagnostico += `✅ Sistema configurado correctamente\n\n`;
    diagnostico += `Si aún ves doble procesamiento:\n`;
    diagnostico += `  • Recarga la página del Google Sheet\n`;
    diagnostico += `  • Verifica que no haya otro usuario editando\n`;
    diagnostico += `  • Espera 1-2 segundos entre clics en checkboxes\n`;
  }

  ui.alert('🔍 Diagnóstico v2.8', diagnostico, ui.ButtonSet.OK);
}

// ====================================
// CONFIGURACIÓN DEL SISTEMA
// ====================================
function configurarHojasCorregido() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  try {
    ui.alert('Configurando Sistema v2.8...',
      'Creando estructura optimizada...\nConfigurando para ' + LLAMADAS_PARA_FINALIZAR + ' llamadas\nNUEVAS FORMACIONES: Análisis de datos E-commerce, SAC, Ofimática\nNUEVAS COLUMNAS: Aliados, Plataformas, Conexión laboral, etc.\nFIX: Sistema anti-doble procesamiento',
      ui.ButtonSet.OK);

    let hojaGeneral = ss.getSheetByName('📋 Seguimiento General');
    if (!hojaGeneral) {
      hojaGeneral = ss.insertSheet('📋 Seguimiento General');
    }

    const hojasLlamadas = [
      '📞 Llamada 1', '📞 Llamada 2', '📞 Llamada 3',
      '📞 Llamada 4', '📞 Llamada 5'
    ];

    hojasLlamadas.forEach(nombre => {
      if (!ss.getSheetByName(nombre)) {
        ss.insertSheet(nombre);
      }
    });

    if (!ss.getSheetByName('✅ Finalizados')) {
      ss.insertSheet('✅ Finalizados');
    }

    configurarHojaGeneralCorregida();
    configurarHojasLlamadasCorregidas();

    ui.alert('✅ Sistema v2.8 Configurado',
      `Sistema configurado exitosamente!\n\n📋 Hojas creadas correctamente\n🎨 Formato aplicado\n✅ Validaciones configuradas\n⚡ Sistema anti-doble procesamiento activado\n📞 Configurado para ${LLAMADAS_PARA_FINALIZAR} llamadas\n🆕 Formaciones: Análisis de datos E-commerce, SAC, Ofimática\n🆕 Columnas de etapas agregadas`,
      ui.ButtonSet.OK);

  } catch (error) {
    console.error('❌ Error en configurarHojasCorregido:', error);
    ui.alert('❌ Error de Configuración',
      `Error: ${error.message}\n\n🔧 Intenta nuevamente`,
      ui.ButtonSet.OK);
  }
}

function configurarHojaGeneralCorregida() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('📋 Seguimiento General');

  if (!hoja) return;

  hoja.clear();
  hoja.clearConditionalFormatRules();

  try {
    const rangeCompleto = hoja.getRange(1, 1, hoja.getMaxRows(), hoja.getMaxColumns());
    rangeCompleto.clearDataValidations();
  } catch (error) {
    limpiarValidacionesPorRangos(hoja);
  }

  configurarEncabezadosYValidacionesCorregidos(hoja);
}

function configurarEncabezadosYValidacionesCorregidos(hoja) {
  const encabezados = [
    'Creamos ID',
    'Nombre Completo',
    'Teléfono',
    'Formación',
    // NUEVAS COLUMNAS DE ETAPAS
    'Aliados',
    'Plataformas',
    'Conexión laboral',
    'Por su cuenta',
    'No busca trabajar',
    'Empleado',
    'No terminó la formación',
    // COLUMNAS ORIGINALES
    'Etapa Actual',
    'Resultados Obtenidos',
    'Documentos Faltantes',
    'Fecha Original',
    'Notas/Última Llamada',
    'Total Llamadas',
    'Procesar'
  ];

  hoja.getRange(1, 1, 1, TOTAL_COLUMNAS).setValues([encabezados]);
  configurarValidacionesCompletas(hoja);
}

function configurarValidacionesCompletas(hoja) {
  try {
    // FORMACIÓN - CON NUEVAS OPCIONES v2.8
    const opcionesFormacion = [
      'Barista I', 'Barista II', 'Barista III', 'Barista IV',
      'Barismo', 'Barismo 1', 'Barismo 2', 'Barismo 3', 'Barismo 4', 'Barismo 5',
      'Barismo 6', 'Barismo 7', 'Barismo 8', 'Barismo 9', 'Barismo 10',
      'Gastronomía', 'Gastronomía I', 'Gastronomía II', 'Gastronomía III',
      'Gastronomía IV', 'Gastronomía V',
      'Gastronomía 1', 'Gastronomía 2', 'Gastronomía 3', 'Gastronomía 4', 'Gastronomía 5',
      'Panadería', 'Repostería', 'Sommelier', 'Food Manager',
      'Análisis de datos E-commerce', 'SAC', 'Ofimática',
      'Otra'
    ];

    const rangoFormacion = hoja.getRange(2, COLUMNAS.FORMACION, 998, 1);
    const validacionFormacion = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesFormacion)
      .setAllowInvalid(true)
      .setHelpText('Selecciona la formación\n🆕 Nuevas: Análisis de datos, SAC, Ofimática')
      .build();
    rangoFormacion.setDataValidation(validacionFormacion);

    // ETAPA ACTUAL
    const opcionesEtapa = [
      'Inicial', 'Aliados', 'Plataformas', 'Conexión Laboral',
      'Por su cuenta', 'No busca trabajar', 'Empleado',
      'No termino la formación', 'Finalizado'
    ];

    const rangoEtapa = hoja.getRange(2, COLUMNAS.ETAPA_ACTUAL, 998, 1);
    const validacionEtapa = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesEtapa)
      .setAllowInvalid(true)
      .build();
    rangoEtapa.setDataValidation(validacionEtapa);

    // DOCUMENTOS FALTANTES
    const opcionesDocumentos = [
      'Salud', 'Manipulación', 'Pulmones', 'Policiacos',
      'Penales', 'Ninguno', 'CV', 'NIT', 'Fotografía', 'DPI'
    ];

    const rangoDocumentos = hoja.getRange(2, COLUMNAS.DOCUMENTOS, 998, 1);
    const validacionDocumentos = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesDocumentos)
      .setAllowInvalid(true)
      .setHelpText('📄 SELECCIÓN MÚLTIPLE:\n1. Selecciona documento\n2. Se agrega automáticamente\n3. Para remover: selecciona el mismo\n4. "Ninguno" borra todo')
      .build();
    rangoDocumentos.setDataValidation(validacionDocumentos);

    // CHECKBOXES
    const rangoCheckbox = hoja.getRange(2, COLUMNAS.PROCESAR, 998, 1);
    rangoCheckbox.insertCheckboxes();

  } catch (error) {
    console.error('Error configurando validaciones:', error);
  }
}

function configurarHojasLlamadasCorregidas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojas = [
    '📞 Llamada 1', '📞 Llamada 2', '📞 Llamada 3',
    '📞 Llamada 4', '📞 Llamada 5', '✅ Finalizados'
  ];

  hojas.forEach(nombreHoja => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;

    hoja.clear();
    hoja.clearConditionalFormatRules();

    const encabezadosDestino = [
      'Creamos ID', 'Nombre Completo', 'Teléfono', 'Formación',
      // NUEVAS COLUMNAS DE ETAPAS
      'Aliados', 'Plataformas', 'Conexión laboral', 'Por su cuenta',
      'No busca trabajar', 'Empleado', 'No terminó la formación',
      // COLUMNAS ORIGINALES
      'Etapa Actual', 'Resultados Obtenidos', 'Documentos Faltantes',
      'Fecha de Llamada', 'Notas/Historial', 'Total Llamadas'
    ];

    hoja.getRange(1, 1, 1, TOTAL_COLUMNAS_DESTINO).setValues([encabezadosDestino]);
  });
}

// ====================================
// PROCESAMIENTO DE FORMACIONES v2.8
// ====================================
function procesarFormacionInteligente(valor, nombreHoja) {
  if (!valor) {
    const hojaNombre = (nombreHoja || '').toLowerCase();
    if (hojaNombre.includes('gastro')) return 'Gastronomía';
    if (hojaNombre.includes('food') || hojaNombre.includes('manager')) return 'Food Manager';
    if (hojaNombre.includes('análisis') || hojaNombre.includes('datos')) return 'Análisis de datos E-commerce';
    if (hojaNombre.includes('sac')) return 'SAC';
    if (hojaNombre.includes('ofimatica')) return 'Ofimática';
    return 'Barismo';
  }

  const texto = valor.toString().toLowerCase().trim();

  // PATRONES PARA NUEVAS FORMACIONES
  const patronesAnalisisDatos = /an[aá]lisis\s*(de\s*)?datos?/i;
  const patronesSAC = /\bsac\b/i;
  const patronesOfimatica = /ofim[aá]tica/i;

  if (patronesAnalisisDatos.test(texto)) return 'Análisis de datos E-commerce';
  if (patronesSAC.test(texto)) return 'SAC';
  if (patronesOfimatica.test(texto)) return 'Ofimática';

  // Patrones existentes
  const patronesBarista = /barista\s*(i{1,4}|1|2|3|4)\b/i;
  const patronesBarismo = /barismo\s*(\d+)?\b/i;
  const patronesGastronomia = /gastronom[íi]a\s*(i{1,5}|[1-5])?\b/i;
  const patronesFoodManager = /food\s*manager/i;
  const patronesPanaderia = /panader[íi]a/i;
  const patronesReposteria = /reporter[íi]a/i;
  const patronesSommelier = /sommelier/i;

  if (patronesFoodManager.test(texto)) return 'Food Manager';

  const matchBarista = texto.match(patronesBarista);
  if (matchBarista) {
    let nivel = matchBarista[1];
    const conversion = { 'i': 'I', 'ii': 'II', 'iii': 'III', 'iv': 'IV' };
    if (conversion[nivel.toLowerCase()]) {
      nivel = conversion[nivel.toLowerCase()];
    }
    return `Barista ${nivel.toUpperCase()}`;
  }

  const matchGastro = texto.match(patronesGastronomia);
  if (matchGastro) {
    const nivel = matchGastro[1];
    if (nivel) {
      const conversion = { 'i': 'I', 'ii': 'II', 'iii': 'III', 'iv': 'IV', 'v': 'V' };
      const nivelFinal = conversion[nivel.toLowerCase()] || nivel;
      return `Gastronomía ${nivelFinal.toUpperCase()}`;
    }
    return 'Gastronomía';
  }

  const matchBarismo = texto.match(patronesBarismo);
  if (matchBarismo || /baris/i.test(texto)) {
    const numero = matchBarismo ? matchBarismo[1] : '';
    return numero ? `Barismo ${numero}` : 'Barismo';
  }

  if (patronesPanaderia.test(texto)) return 'Panadería';
  if (patronesReposteria.test(texto)) return 'Repostería';
  if (patronesSommelier.test(texto)) return 'Sommelier';

  const numeroMatch = texto.match(/(\d+)/);
  if (numeroMatch) {
    return `Barismo ${numeroMatch[1]}`;
  }

  return 'Barismo';
}

// ====================================
// DETECCIÓN DE COLUMNAS
// ====================================
function detectarColumnasOptimizadas(encabezados) {
  const encabezadosLower = encabezados.map(h => h.toString().toLowerCase().trim());

  return {
    id: buscarColumnaInteligente(encabezadosLower, ['creamos id', 'id', 'codigo', 'identificador', 'student_id']),
    nombre: buscarColumnaInteligente(encabezadosLower, ['nombre', 'nombre completo', 'name', 'participante', 'estudiante']),
    telefono: buscarColumnaInteligente(encabezadosLower, ['telefono', 'teléfono', 'phone', 'celular', 'contacto']),
    formacion: buscarColumnaInteligente(encabezadosLower, ['formacion', 'formación', 'programa', 'curso']),
    etapa: buscarColumnaInteligente(encabezadosLower, ['etapa actual', 'etapa', 'fase', 'estado']),
    resultados: buscarColumnaInteligente(encabezadosLower, ['resultados obtenidos', 'resultados', 'logros']),
    documentos: buscarColumnaInteligente(encabezadosLower, ['documentos faltantes', 'documentos', 'papeleria']),
    paso: buscarColumnaInteligente(encabezadosLower, ['paso a seguir', 'siguiente paso', 'accion']),
    fecha: buscarColumnaInteligente(encabezadosLower, ['fecha', 'date', 'timestamp']),
    notas: buscarColumnaInteligente(encabezadosLower, ['notas', 'observaciones', 'comentarios'])
  };
}

function buscarColumnaInteligente(encabezados, palabrasClave) {
  for (let i = 0; i < encabezados.length; i++) {
    for (let palabra of palabrasClave) {
      if (encabezados[i] === palabra) {
        return { encontrada: true, indice: i, nombre: encabezados[i], tipo: 'exacto' };
      }
    }
  }

  for (let i = 0; i < encabezados.length; i++) {
    for (let palabra of palabrasClave) {
      if (encabezados[i].includes(palabra)) {
        return { encontrada: true, indice: i, nombre: encabezados[i], tipo: 'parcial' };
      }
    }
  }

  return { encontrada: false, indice: -1, nombre: '', tipo: 'no_encontrado' };
}

// PROCESAMIENTO DE IMPORTACIÓN
// ====================================
function procesarDatosImportacionOptimizada(datosOriginales, nombreHoja) {
  if (!datosOriginales || datosOriginales.length <= 1) {
    throw new Error('No hay datos válidos para procesar');
  }

  const encabezados = datosOriginales[0];
  const filasDatos = datosOriginales.slice(1);

  console.log('Encabezados detectados:', encabezados);

  const columnas = detectarColumnasOptimizadas(encabezados);

  console.log('Columnas detectadas:', columnas);

  const datosProcessados = [];
  let filasVacias = 0;

  filasDatos.forEach((fila, index) => {
    const nombre = columnas.nombre.encontrada ?
      obtenerValorColumnaSeguro(fila, columnas.nombre) : '';

    if (!nombre || nombre.length < 2) {
      filasVacias++;
      return;
    }

    const formacionOriginal = columnas.formacion.encontrada ?
      obtenerValorColumnaSeguro(fila, columnas.formacion) : '';
    const formacionProcesada = procesarFormacionInteligente(formacionOriginal, nombreHoja);

    const etapaOriginal = columnas.etapa.encontrada ?
      obtenerValorColumnaSeguro(fila, columnas.etapa) : '';
    const etapaProcesada = procesarEtapaInteligente(etapaOriginal);

    const resultadosObtenidos = columnas.resultados.encontrada ?
      obtenerValorColumnaSeguro(fila, columnas.resultados) : '';

    const fechaOriginal = columnas.fecha.encontrada ?
      procesarFechaOriginal(obtenerValorColumnaSeguro(fila, columnas.fecha)) : '';

    const registro = [
      obtenerValorColumnaSeguro(fila, columnas.id) || generarIdTemporal(index),
      nombre.trim(),
      obtenerValorColumnaSeguro(fila, columnas.telefono) || '',
      formacionProcesada,
      '', '', '', '', '', '', '', // 7 columnas nuevas de etapas (vacías al importar)
      etapaProcesada,
      resultadosObtenidos,
      obtenerValorColumnaSeguro(fila, columnas.documentos) || 'Ninguno',
      fechaOriginal,
      generarNotasImportacion(obtenerValorColumnaSeguro(fila, columnas.notas), nombreHoja),
      0,
      false
    ];

    datosProcessados.push(registro);
  });

  console.log(`Procesados: ${datosProcessados.length} registros válidos, ${filasVacias} filas vacías omitidas`);
  return datosProcessados;
}

function obtenerValorColumnaSeguro(fila, columnaInfo) {
  if (!columnaInfo.encontrada || columnaInfo.indice >= fila.length) {
    return '';
  }

  const valor = fila[columnaInfo.indice];
  if (valor === null || valor === undefined) {
    return '';
  }

  return valor.toString().trim();
}

function generarIdTemporal(index) {
  const timestamp = new Date().getTime();
  return `TEMP_${timestamp}_${index}`;
}

function procesarFechaOriginal(valor) {
  if (!valor) return '';

  try {
    const fecha = new Date(valor);
    if (isNaN(fecha.getTime())) return valor.toString();

    return Utilities.formatDate(fecha, Session.getScriptTimeZone(), 'dd/MM/yyyy');
  } catch (error) {
    return valor.toString();
  }
}

function generarNotasImportacion(notasOriginales, nombreHoja) {
  const notasBase = notasOriginales || '';
  return `${notasBase}${notasBase ? ' | ' : ''}Importado de "${nombreHoja}"`.trim();
}

function procesarEtapaInteligente(valor) {
  if (!valor) return 'Inicial';

  const texto = valor.toString().toLowerCase().trim();

  const mapeoEtapas = {
    'aliados': 'Aliados',
    'plataformas': 'Plataformas',
    'conexion laboral': 'Conexión Laboral',
    'conexión laboral': 'Conexión Laboral',
    'por su cuenta': 'Por su cuenta',
    'no busca trabajar': 'No busca trabajar',
    'empleado': 'Empleado',
    'no termino': 'No termino la formación',
    'finalizado': 'Finalizado',
    'inicial': 'Inicial'
  };

  for (const [patron, etapa] of Object.entries(mapeoEtapas)) {
    if (texto.includes(patron)) {
      return etapa;
    }
  }

  return 'Inicial';
}

// ====================================
// PROCESAMIENTO DE LLAMADAS v2.8 - CON ESCRITURA EN COLUMNAS DE ETAPAS
// ====================================
function procesarLlamadaOptimizada(hojaGeneral, fila) {
  const timestamp = new Date().getTime();
  console.log(`[${timestamp}] 🔄 Iniciando procesamiento de fila ${fila}`);

  try {
    const rangoDatos = hojaGeneral.getRange(fila, 1, 1, TOTAL_COLUMNAS);
    const datos = rangoDatos.getValues()[0];

    const nombre = datos[COLUMNAS.NOMBRE - 1] ? datos[COLUMNAS.NOMBRE - 1].toString().trim() : '';
    if (!nombre) {
      return { exito: false, error: 'Nombre vacío o inválido', movido: false, totalLlamadas: 0 };
    }

    let llamadasActuales = parseInt(datos[COLUMNAS.TOTAL_LLAMADAS - 1]) || 0;
    const nuevasLlamadas = llamadasActuales + 1;
    const fechaActual = new Date();

    console.log(`📊 ${nombre}: ${llamadasActuales} → ${nuevasLlamadas} llamadas`);

    let nombreHojaDestino;
    let finalizando = false;

    if (nuevasLlamadas >= LLAMADAS_PARA_FINALIZAR) {
      nombreHojaDestino = '✅ Finalizados';
      finalizando = true;
    } else {
      nombreHojaDestino = `📞 Llamada ${nuevasLlamadas}`;
    }

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDestino = ss.getSheetByName(nombreHojaDestino);

    if (!hojaDestino) {
      return { exito: false, error: `Hoja "${nombreHojaDestino}" no encontrada`, movido: false, totalLlamadas: nuevasLlamadas };
    }

    const notasActuales = datos[COLUMNAS.NOTAS - 1] || '';
    const etapaActual = datos[COLUMNAS.ETAPA_ACTUAL - 1] || 'En proceso';
    const notasActualizadas = `${notasActuales}${notasActuales ? ' | ' : ''}Llamada ${nuevasLlamadas} - ${Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'dd/MM/yyyy')}`.trim();

    // *** DETERMINAR EN QUÉ COLUMNA DE ETAPA ESCRIBIR ***
    const infoEtapa = determinarColumnaEtapa(etapaActual, nuevasLlamadas);

    // Crear array con las 7 columnas de etapas
    const columnasEtapas = [
      datos[COLUMNAS.ALIADOS - 1] || '',
      datos[COLUMNAS.PLATAFORMAS - 1] || '',
      datos[COLUMNAS.CONEXION_LABORAL - 1] || '',
      datos[COLUMNAS.POR_SU_CUENTA - 1] || '',
      datos[COLUMNAS.NO_BUSCA_TRABAJAR - 1] || '',
      datos[COLUMNAS.EMPLEADO - 1] || '',
      datos[COLUMNAS.NO_TERMINO_FORMACION - 1] || ''
    ];

    // Escribir en la columna correspondiente según la etapa
    if (infoEtapa.indice !== -1) {
      const textoExistente = columnasEtapas[infoEtapa.indice];
      const fechaFormateada = Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
      const nuevoTexto = textoExistente ?
        `${textoExistente} | ${fechaFormateada} - Llamada ${nuevasLlamadas}` :
        `${fechaFormateada} - Llamada ${nuevasLlamadas}`;
      columnasEtapas[infoEtapa.indice] = nuevoTexto;
    }

    const datosDestino = [
      datos[COLUMNAS.ID - 1] || '',
      nombre,
      datos[COLUMNAS.TELEFONO - 1] || '',
      datos[COLUMNAS.FORMACION - 1] || 'Barismo',
      // 7 columnas de etapas actualizadas
      columnasEtapas[0],
      columnasEtapas[1],
      columnasEtapas[2],
      columnasEtapas[3],
      columnasEtapas[4],
      columnasEtapas[5],
      columnasEtapas[6],
      // Columnas restantes
      etapaActual,
      datos[COLUMNAS.RESULTADOS - 1] || '',
      datos[COLUMNAS.DOCUMENTOS - 1] || 'Ninguno',
      Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      notasActualizadas,
      nuevasLlamadas
    ];

    const resultadoInsercion = insertarEnHojaDestinoOptimizada(hojaDestino, datosDestino, nombreHojaDestino);

    if (!resultadoInsercion.exito) {
      return { exito: false, error: resultadoInsercion.error, movido: false, totalLlamadas: nuevasLlamadas };
    }

    if (finalizando) {
      hojaGeneral.deleteRow(fila);
      console.log(`✅ ${nombre} movido a Finalizados`);
      return { exito: true, error: null, movido: true, totalLlamadas: nuevasLlamadas };
    } else {
      // Actualizar la fila actual con las nuevas columnas de etapas
      hojaGeneral.getRange(fila, COLUMNAS.ALIADOS, 1, 7).setValues([columnasEtapas]);
      hojaGeneral.getRange(fila, COLUMNAS.NOTAS).setValue(notasActualizadas);
      hojaGeneral.getRange(fila, COLUMNAS.TOTAL_LLAMADAS).setValue(nuevasLlamadas);
      hojaGeneral.getRange(fila, COLUMNAS.PROCESAR).setValue(false);

      console.log(`✅ ${nombre} procesado - Llamada ${nuevasLlamadas} - Etapa: ${infoEtapa.nombre}`);
      return { exito: true, error: null, movido: false, totalLlamadas: nuevasLlamadas };
    }

  } catch (error) {
    console.error(`❌ Error procesando fila ${fila}:`, error);
    return { exito: false, error: error.message, movido: false, totalLlamadas: 0 };
  }
}

// *** NUEVA FUNCIÓN: Determinar en qué columna de etapa escribir ***
function determinarColumnaEtapa(etapa, numeroLlamada) {
  const etapaLower = etapa ? etapa.toString().toLowerCase().trim() : '';

  // Mapeo de etapas a índices de columnas (0-6 para columnas 5-11)
  const mapeoEtapas = {
    'aliados': { indice: 0, nombre: 'Aliados' },
    'plataformas': { indice: 1, nombre: 'Plataformas' },
    'conexion laboral': { indice: 2, nombre: 'Conexión Laboral' },
    'conexión laboral': { indice: 2, nombre: 'Conexión Laboral' },
    'por su cuenta': { indice: 3, nombre: 'Por su cuenta' },
    'no busca trabajar': { indice: 4, nombre: 'No busca trabajar' },
    'empleado': { indice: 5, nombre: 'Empleado' },
    'no termino': { indice: 6, nombre: 'No terminó la formación' },
    'no terminó': { indice: 6, nombre: 'No terminó la formación' }
  };

  // Buscar coincidencia
  for (const [patron, info] of Object.entries(mapeoEtapas)) {
    if (etapaLower.includes(patron)) {
      return info;
    }
  }

  // Si no hay coincidencia, escribir en la primera columna por defecto
  return { indice: -1, nombre: 'Sin etapa específica' };
}

function insertarEnHojaDestinoOptimizada(hojaDestino, datosDestino, nombreHoja) {
  try {
    if (!Array.isArray(datosDestino) || datosDestino.length !== TOTAL_COLUMNAS_DESTINO) {
      return { exito: false, error: `Datos inválidos: ${datosDestino.length} columnas, se requieren ${TOTAL_COLUMNAS_DESTINO}` };
    }

    hojaDestino.insertRows(2, 1);
    const rangoDestino = hojaDestino.getRange(2, 1, 1, TOTAL_COLUMNAS_DESTINO);
    rangoDestino.setValues([datosDestino]);

    aplicarFormatoFilaDestino(hojaDestino, 2, nombreHoja);

    return { exito: true, error: null };

  } catch (error) {
    return { exito: false, error: error.message };
  }
}

function aplicarFormatoFilaDestino(hoja, fila, nombreHoja) {
  try {
    const configuracionFormato = {
      '📞 Llamada 1': { fondo: '#e3f2fd', borde: '#1976d2' },
      '📞 Llamada 2': { fondo: '#f3e5f5', borde: '#7b1fa2' },
      '📞 Llamada 3': { fondo: '#fff3e0', borde: '#f57c00' },
      '📞 Llamada 4': { fondo: '#ffecb3', borde: '#ff8f00' },
      '📞 Llamada 5': { fondo: '#fce4ec', borde: '#c2185b' },
      '✅ Finalizados': { fondo: '#e8f5e8', borde: '#388e3c' }
    };

    const config = configuracionFormato[nombreHoja] || { fondo: '#ffffff', borde: '#dee2e6' };

    const rangoFila = hoja.getRange(fila, 1, 1, TOTAL_COLUMNAS_DESTINO);
    rangoFila.setBackground(config.fondo);
    rangoFila.setBorder(true, true, true, true, true, true, config.borde, SpreadsheetApp.BorderStyle.SOLID);

  } catch (error) {
    console.warn('Error aplicando formato:', error.message);
  }
}

// [CONTINÚA EN SIGUIENTE MENSAJE - El archivo es muy extenso]
// ====================================
// CONTINUACIÓN DEL CÓDIGO - PARTE 2
// ====================================

// ====================================
// FUNCIONES DE IMPORTACIÓN
// ====================================
function importarDatos2024() {
  importarDatosDesdeSheetExternoOptimizado('2024');
}

function importarDatos2025() {
  importarDatosDesdeSheetExternoOptimizado('2025');
}

function importarDatosPersonalizados() {
  const ui = SpreadsheetApp.getUi();
  const respuesta = ui.prompt(
    '📥 Importar Desde Otra Hoja',
    'Escribe el nombre EXACTO de la hoja/pestaña:',
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() === ui.Button.OK) {
    const nombreHoja = respuesta.getResponseText().trim();
    if (nombreHoja) {
      importarDatosDesdeSheetExternoOptimizado(nombreHoja);
    } else {
      ui.alert('❌ Nombre requerido', 'Debes escribir el nombre exacto de la hoja', ui.ButtonSet.OK);
    }
  }
}

function configurarSheetExterno() {
  const ui = SpreadsheetApp.getUi();
  const propiedades = PropertiesService.getScriptProperties();
  const urlActual = propiedades.getProperty('URL_SHEET_EXTERNO') || '';

  const mensaje = `🔗 CONFIGURAR GOOGLE SHEET EXTERNO\n\n` +
    `${urlActual ? `📋 URL actual: ✅ Configurada` : '❌ No hay URL configurada'}\n\n` +
    `¿Continuar con la configuración?`;

  const respuesta = ui.alert('🔗 Google Sheet Externo', mensaje, ui.ButtonSet.YES_NO);

  if (respuesta === ui.Button.YES) {
    const nuevaURL = ui.prompt(
      '📋 URL del Google Sheet',
      'Pega la URL COMPLETA del Google Sheet:',
      ui.ButtonSet.OK_CANCEL
    );

    if (nuevaURL.getSelectedButton() === ui.Button.OK) {
      const url = nuevaURL.getResponseText().trim();
      if (validarURLOptimizada(url)) {
        propiedades.setProperty('URL_SHEET_EXTERNO', url);
        ui.alert('✅ Configuración exitosa!',
          'Google Sheet externo configurado correctamente.',
          ui.ButtonSet.OK);
      } else {
        ui.alert('❌ URL inválida',
          'La URL no es válida. Debe ser una URL completa de Google Sheets.',
          ui.ButtonSet.OK);
      }
    }
  }
}

function validarURLOptimizada(url) {
  if (!url || typeof url !== 'string') return false;
  return url.includes('docs.google.com/spreadsheets') &&
         url.includes('/d/') &&
         url.length > 50;
}

function importarDatosDesdeSheetExternoOptimizado(nombreHoja) {
  const ui = SpreadsheetApp.getUi();

  try {
    const propiedades = PropertiesService.getScriptProperties();
    const url = propiedades.getProperty('URL_SHEET_EXTERNO');

    if (!url) {
      ui.alert('❌ Configuración faltante',
        'Primero configura el Google Sheet externo',
        ui.ButtonSet.OK);
      return;
    }

    const matches = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!matches) {
      ui.alert('❌ URL inválida',
        'La URL configurada no es válida.',
        ui.ButtonSet.OK);
      return;
    }

    const sheetId = matches[1];

    ui.alert('🔄 Importando...',
      `Conectando con Google Sheet externo...\nBuscando hoja: "${nombreHoja}"`,
      ui.ButtonSet.OK);

    const sheetExterno = SpreadsheetApp.openById(sheetId);
    const hojaExterna = sheetExterno.getSheetByName(nombreHoja);

    if (!hojaExterna) {
      const hojasDisponibles = sheetExterno.getSheets()
        .map(h => h.getName())
        .join(', ');

      ui.alert('❌ Hoja no encontrada',
        `No existe la hoja "${nombreHoja}"\n\nHojas disponibles: ${hojasDisponibles}`,
        ui.ButtonSet.OK);
      return;
    }

    const rangoDatos = hojaExterna.getDataRange();
    if (rangoDatos.getNumRows() <= 1) {
      ui.alert('❌ Hoja vacía',
        `La hoja "${nombreHoja}" está vacía o solo tiene encabezados.`,
        ui.ButtonSet.OK);
      return;
    }

    const datosOriginales = rangoDatos.getValues();
    const datosImportados = procesarDatosImportacionOptimizada(datosOriginales, nombreHoja);

    if (datosImportados.length === 0) {
      ui.alert('❌ Sin datos válidos',
        `No se encontraron datos válidos en "${nombreHoja}".`,
        ui.ButtonSet.OK);
      return;
    }

    if (!verificarSistemaCompleto()) {
      ui.alert('❌ Sistema no configurado',
        'Primero configura el sistema local',
        ui.ButtonSet.OK);
      return;
    }

    const registrosInsertados = insertarDatosOptimizado(datosImportados);

    const mensaje = `✅ IMPORTACIÓN EXITOSA\n\n` +
      `📊 ${registrosInsertados} registros importados\n` +
      `📋 Desde: "${sheetExterno.getName()}"\n` +
      `📄 Hoja: "${nombreHoja}"\n` +
      `📞 Sistema configurado para ${LLAMADAS_PARA_FINALIZAR} llamadas`;

    ui.alert('✅ Importación Completada', mensaje, ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de importación',
      `Error al importar desde "${nombreHoja}": ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function insertarDatosOptimizado(datos) {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');

  if (!hojaGeneral) {
    throw new Error('La hoja principal no existe. Configura el sistema primero.');
  }

  if (datos.length === 0) {
    throw new Error('No hay datos para insertar');
  }

  try {
    const filaInicio = 2;
    hojaGeneral.insertRows(filaInicio, datos.length);

    const rango = hojaGeneral.getRange(filaInicio, 1, datos.length, TOTAL_COLUMNAS);
    rango.setValues(datos);

    aplicarValidacionesNuevasFilas(hojaGeneral, filaInicio, datos.length);
    aplicarColorFormacionCorregido();

    return datos.length;

  } catch (error) {
    throw new Error(`Error al insertar datos: ${error.message}`);
  }
}

function aplicarValidacionesNuevasFilas(hoja, filaInicio, cantidadFilas) {
  try {
    const opcionesFormacion = [
      'Barista I', 'Barista II', 'Barista III', 'Barista IV',
      'Barismo', 'Barismo 1', 'Barismo 2', 'Barismo 3', 'Barismo 4', 'Barismo 5',
      'Gastronomía', 'Gastronomía I', 'Gastronomía II', 'Gastronomía III',
      'Panadería', 'Repostería', 'Sommelier', 'Food Manager',
      'Análisis de datos E-commerce', 'SAC', 'Ofimática',
      'Otra'
    ];

    const rangoFormacion = hoja.getRange(filaInicio, COLUMNAS.FORMACION, cantidadFilas, 1);
    const validacionFormacion = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesFormacion)
      .setAllowInvalid(true)
      .build();
    rangoFormacion.setDataValidation(validacionFormacion);

    const opcionesDocumentos = [
      'Salud', 'Manipulación', 'Pulmones', 'Policiacos',
      'Penales', 'Ninguno', 'CV', 'NIT', 'Fotografía', 'DPI'
    ];

    const rangoDocumentos = hoja.getRange(filaInicio, COLUMNAS.DOCUMENTOS, cantidadFilas, 1);
    const validacionDocumentos = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesDocumentos)
      .setAllowInvalid(true)
      .setHelpText('Selecciona documentos faltantes (separados por comas para múltiples)')
      .build();
    rangoDocumentos.setDataValidation(validacionDocumentos);

    const rangoCheckbox = hoja.getRange(filaInicio, COLUMNAS.PROCESAR, cantidadFilas, 1);
    rangoCheckbox.insertCheckboxes();

  } catch (error) {
    console.error('Error aplicando validaciones:', error);
  }
}

function aplicarColorFormacionCorregido() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('📋 Seguimiento General');
  if (!hoja) return;

  const ultimaFila = hoja.getLastRow();
  if (ultimaFila <= 1) return;

  const rangoFormacion = hoja.getRange(2, COLUMNAS.FORMACION, ultimaFila - 1, 1);
  const valoresFormacion = rangoFormacion.getValues();

  const coloresFormacion = {
    'Barista I': '#E3F2FD',
    'Barista II': '#BBDEFB',
    'Barista III': '#90CAF9',
    'Barista IV': '#64B5F6',
    'Barismo': '#BBDEFB',
    'Gastronomía': '#FFECB3',
    'Gastronomía I': '#FFF3E0',
    'Gastronomía II': '#FFE0B2',
    'Gastronomía III': '#FFCC80',
    'Food Manager': '#E8F5E8',
    'Panadería': '#F3E5AB',
    'Repostería': '#F8BBD9',
    'Sommelier': '#E1BEE7',
    'Análisis de datos E-commerce': '#B2EBF2',
    'SAC': '#C5E1A5',
    'Ofimática': '#D1C4E9',
    'Otra': '#E0E0E0'
  };

  const coloresPorDefecto = '#FFFFFF';

  const coloresParaAplicar = valoresFormacion.map(fila => {
    const formacion = fila[0];
    let colorFila = coloresPorDefecto;

    if (formacion) {
      const formacionStr = formacion.toString().trim();

      if (coloresFormacion[formacionStr]) {
        colorFila = coloresFormacion[formacionStr];
      }
      else if (formacionStr.toLowerCase().includes('análisis') || formacionStr.toLowerCase().includes('datos')) {
        colorFila = coloresFormacion['Análisis de datos E-commerce'];
      }
      else if (formacionStr.toLowerCase() === 'sac') {
        colorFila = coloresFormacion['SAC'];
      }
      else if (formacionStr.toLowerCase().includes('ofimatica')) {
        colorFila = coloresFormacion['Ofimática'];
      }
      else if (formacionStr.toLowerCase().includes('food') && formacionStr.toLowerCase().includes('manager')) {
        colorFila = coloresFormacion['Food Manager'];
      }
      else if (formacionStr.toLowerCase().includes('barista') || formacionStr.toLowerCase().includes('barismo')) {
        colorFila = coloresFormacion['Barismo'];
      }
      else if (formacionStr.toLowerCase().includes('gastronomía') || formacionStr.toLowerCase().includes('gastronom')) {
        colorFila = coloresFormacion['Gastronomía'];
      }
    }

    return new Array(TOTAL_COLUMNAS).fill(colorFila);
  });

  const rangoFilas = hoja.getRange(2, 1, ultimaFila - 1, TOTAL_COLUMNAS);
  rangoFilas.setBackgrounds(coloresParaAplicar);
}

// ====================================
// PROCESAMIENTO AUTOMÁTICO
// ====================================
function activarProcesomientoAutomatico() {
  const ui = SpreadsheetApp.getUi();

  const confirmacion = ui.alert('⚡ ACTIVAR Procesamiento Automático',
    `¿Activar procesamiento automático?\n\nLas llamadas se procesarán automáticamente al marcar checkboxes.\n\n🆕 v2.8: Sistema anti-doble procesamiento + columnas de etapas`,
    ui.ButtonSet.YES_NO);

  if (confirmacion !== ui.Button.YES) return;

  try {
    if (!verificarSistemaCompleto()) {
      ui.alert('❌ Sistema no configurado',
        'Primero configura el sistema completamente',
        ui.ButtonSet.OK);
      return;
    }

    const propiedades = PropertiesService.getScriptProperties();

    // Limpiar bloqueos antes de activar
    limpiarBloqueosProcesamiento();

    propiedades.setProperty('PROCESAMIENTO_AUTOMATICO', 'true');
    propiedades.setProperty('FECHA_ACTIVACION_AUTO', new Date().toISOString());

    configurarTriggerOnEdit();

    ui.alert('⚡ Procesamiento Automático Activado',
      'Procesamiento automático activado exitosamente!\n\n✅ Sistema anti-doble procesamiento activo\n📞 Marca checkboxes para procesar llamadas\n📝 Se registrará en columnas de etapas\n\n💡 Espera 1-2 segundos entre cada checkbox',
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Activación',
      `No se pudo activar: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function desactivarProcesomientoAutomatico() {
  const ui = SpreadsheetApp.getUi();

  const confirmacion = ui.alert('🔴 DESACTIVAR Procesamiento Automático',
    '¿Desactivar procesamiento automático?\n\nDeberás usar procesamiento manual.',
    ui.ButtonSet.YES_NO);

  if (confirmacion !== ui.Button.YES) return;

  try {
    const propiedades = PropertiesService.getScriptProperties();
    propiedades.setProperty('PROCESAMIENTO_AUTOMATICO', 'false');
    propiedades.setProperty('FECHA_DESACTIVACION_AUTO', new Date().toISOString());

    removerTriggerOnEdit();

    // Limpiar bloqueos al desactivar
    limpiarBloqueosProcesamiento();

    ui.alert('🔴 Procesamiento Automático Desactivado',
      'Procesamiento automático desactivado.\n\n📋 Bloqueos limpiados\n🔄 Usa "Procesar Llamadas Marcadas" para procesamiento manual.',
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Desactivación',
      `Error al desactivar: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function verificarEstadoProcesomientoAutomatico() {
  try {
    const propiedades = PropertiesService.getScriptProperties();
    const estado = propiedades.getProperty('PROCESAMIENTO_AUTOMATICO') === 'true';

    if (estado) {
      console.log('⚡ Procesamiento automático: ACTIVO');
    } else {
      console.log('📋 Procesamiento automático: INACTIVO');
    }

    return estado;
  } catch (error) {
    return false;
  }
}

function configurarTriggerOnEdit() {
  try {
    removerTriggerOnEdit();

    ScriptApp.newTrigger('onEdit')
      .forSpreadsheet(SpreadsheetApp.getActive())
      .onEdit()
      .create();

    console.log('✅ Trigger onEdit configurado');

  } catch (error) {
    throw new Error(`No se pudo configurar el trigger automático: ${error.message}`);
  }
}

function removerTriggerOnEdit() {
  try {
    const triggers = ScriptApp.getProjectTriggers();
    const triggersOnEdit = triggers.filter(trigger =>
      trigger.getHandlerFunction() === 'onEdit' &&
      trigger.getEventType() === ScriptApp.EventType.ON_EDIT
    );

    triggersOnEdit.forEach(trigger => {
      ScriptApp.deleteTrigger(trigger);
    });

    console.log(`✅ ${triggersOnEdit.length} triggers onEdit removidos`);

  } catch (error) {
    console.error('Error removiendo triggers:', error);
  }
}

function mostrarNotificacionDiscreta(mensaje) {
  try {
    SpreadsheetApp.getActiveSpreadsheet().toast(mensaje, 'Sistema', 4);
  } catch (error) {
    console.log(`Notificación: ${mensaje}`);
  }
}

// ====================================
// PROCESAMIENTO MANUAL
// ====================================
function procesarLlamadasManualesCorregido() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');

  if (!hojaGeneral) {
    ui.alert('❌ Sistema no configurado', 'Primero configura el sistema', ui.ButtonSet.OK);
    return;
  }

  const ultimaFila = hojaGeneral.getLastRow();
  if (ultimaFila <= 1) {
    ui.alert('ℹ️ Sin datos', 'No hay participantes para procesar', ui.ButtonSet.OK);
    return;
  }

  if (!verificarSistemaCompleto()) {
    ui.alert('❌ Sistema incompleto', 'Faltan hojas del sistema', ui.ButtonSet.OK);
    return;
  }

  let procesados = 0;
  let movidos = 0;
  let errores = 0;
  const erroresDetalle = [];

  for (let fila = ultimaFila; fila >= 2; fila--) {
    try {
      const checkbox = obtenerCheckboxSeguro(hojaGeneral, fila);
      if (checkbox === true) {
        const nombre = hojaGeneral.getRange(fila, COLUMNAS.NOMBRE).getValue();
        const resultado = procesarLlamadaOptimizada(hojaGeneral, fila);

        if (resultado.exito) {
          procesados++;
          if (resultado.movido) {
            movidos++;
          }
        } else {
          errores++;
          erroresDetalle.push(`Fila ${fila} (${nombre}): ${resultado.error}`);
        }
      }
    } catch (error) {
      errores++;
      const nombre = hojaGeneral.getRange(fila, COLUMNAS.NOMBRE).getValue() || 'Sin nombre';
      erroresDetalle.push(`Fila ${fila} (${nombre}): ${error.message}`);
    }
  }

  mostrarResultadosProcesamiento(procesados, movidos, errores, erroresDetalle);
}

function mostrarResultadosProcesamiento(procesados, movidos, errores, erroresDetalle) {
  const ui = SpreadsheetApp.getUi();

  if (procesados === 0 && errores === 0) {
    ui.alert('ℹ️ Sin procesamiento',
      'No se encontraron checkboxes marcados para procesar.',
      ui.ButtonSet.OK);
    return;
  }

  let mensaje = `📞 PROCESAMIENTO COMPLETADO v2.8\n\n`;

  if (procesados > 0) {
    mensaje += `✅ ÉXITO: ${procesados} llamadas procesadas\n`;
    if (movidos > 0) {
      mensaje += `🎯 FINALIZADOS: ${movidos} participantes\n`;
    }
  }

  if (errores > 0) {
    mensaje += `❌ ERRORES: ${errores} problemas encontrados\n`;
    mensaje += erroresDetalle.slice(0, 3).join('\n');
    if (erroresDetalle.length > 3) {
      mensaje += `\n... y ${erroresDetalle.length - 3} errores más`;
    }
  }

  ui.alert('📞 Resultado del Procesamiento', mensaje, ui.ButtonSet.OK);
}

function obtenerCheckboxSeguro(hoja, fila) {
  try {
    const valorCheckbox = hoja.getRange(fila, COLUMNAS.PROCESAR).getValue();

    if (valorCheckbox === true ||
        valorCheckbox === 'TRUE' ||
        valorCheckbox === 1 ||
        valorCheckbox === '✓' ||
        valorCheckbox === 'true') {
      return true;
    }

    return false;
  } catch (error) {
    return false;
  }
}

function verificarSistemaCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojasRequeridas = [
    '📋 Seguimiento General',
    '📞 Llamada 1',
    '📞 Llamada 2',
    '📞 Llamada 3',
    '📞 Llamada 4',
    '📞 Llamada 5',
    '✅ Finalizados'
  ];

  for (let nombreHoja of hojasRequeridas) {
    if (!ss.getSheetByName(nombreHoja)) {
      return false;
    }
  }

  return true;
}

// ====================================
// PROCESAMIENTO DE DOCUMENTOS FALTANTES
// ====================================
function procesarSeleccionDocumentos(hoja, fila, valorNuevo, valorAnterior) {
  try {
    console.log('=== PROCESAMIENTO DE DOCUMENTOS ===');
    console.log(`Fila: ${fila}, Nuevo: "${valorNuevo}", Anterior: "${valorAnterior}"`);

    const documentosValidos = [
      'Salud', 'Manipulación', 'Pulmones', 'Policiacos',
      'Penales', 'Ninguno', 'CV', 'NIT', 'Fotografía', 'DPI'
    ];

    if (!valorNuevo || valorNuevo === '') {
      return;
    }

    const valorNuevoStr = valorNuevo.toString().trim();

    if (!documentosValidos.includes(valorNuevoStr)) {
      hoja.getRange(fila, COLUMNAS.DOCUMENTOS).setValue(valorAnterior || 'Ninguno');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `"${valorNuevoStr}" no es válido. Documentos permitidos: ${documentosValidos.join(', ')}`,
        'Documento Inválido',
        6
      );
      return;
    }

    let documentosActuales = [];
    if (valorAnterior && valorAnterior !== '') {
      documentosActuales = valorAnterior.toString()
        .split(',')
        .map(d => d.trim())
        .filter(d => d !== '' && documentosValidos.includes(d));
    }

    if (valorNuevoStr === 'Ninguno') {
      hoja.getRange(fila, COLUMNAS.DOCUMENTOS).setValue('Ninguno');
      SpreadsheetApp.getActiveSpreadsheet().toast(
        'Documentos faltantes: Ninguno',
        'Actualizado',
        3
      );
      return;
    }

    documentosActuales = documentosActuales.filter(d => d !== 'Ninguno');

    if (documentosActuales.includes(valorNuevoStr)) {
      documentosActuales = documentosActuales.filter(d => d !== valorNuevoStr);

      const valorFinal = documentosActuales.length > 0 ?
        documentosActuales.sort().join(', ') : 'Ninguno';

      hoja.getRange(fila, COLUMNAS.DOCUMENTOS).setValue(valorFinal);

      SpreadsheetApp.getActiveSpreadsheet().toast(
        `"${valorNuevoStr}" removido. Actual: ${valorFinal}`,
        'Documento Removido',
        4
      );

    } else {
      documentosActuales.push(valorNuevoStr);

      const valorFinal = documentosActuales.sort().join(', ');

      hoja.getRange(fila, COLUMNAS.DOCUMENTOS).setValue(valorFinal);

      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Documento agregado: ${valorFinal}`,
        'Actualizado',
        4
      );
    }

  } catch (error) {
    console.error('Error en procesarSeleccionDocumentos:', error);
    hoja.getRange(fila, COLUMNAS.DOCUMENTOS).setValue(valorAnterior || 'Ninguno');
    SpreadsheetApp.getActiveSpreadsheet().toast(
      `Error procesando documentos: ${error.message}`,
      'Error',
      5
    );
  }
}
// ====================================
// CONTINUACIÓN DEL CÓDIGO - PARTE 3
// ====================================

// ====================================
// FINALIZACIÓN POR ETAPA
// ====================================
function procesarFinalizacionPorEtapa(hojaGeneral, fila) {
  try {
    const nombre = hojaGeneral.getRange(fila, COLUMNAS.NOMBRE).getValue();

    if (!verificarSistemaCompleto()) {
      throw new Error('Sistema no configurado completamente');
    }

    const rangoDatos = hojaGeneral.getRange(fila, 1, 1, TOTAL_COLUMNAS);
    const datos = rangoDatos.getValues()[0];

    const nombreCompleto = datos[COLUMNAS.NOMBRE - 1] ? datos[COLUMNAS.NOMBRE - 1].toString().trim() : '';
    if (!nombreCompleto) {
      throw new Error('Nombre vacío o inválido');
    }

    const fechaActual = new Date();
    const nuevasLlamadas = LLAMADAS_PARA_FINALIZAR;

    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const hojaDestino = ss.getSheetByName('✅ Finalizados');

    if (!hojaDestino) {
      throw new Error('Hoja "✅ Finalizados" no encontrada');
    }

    const notasFinalizacion = `${datos[COLUMNAS.NOTAS - 1] || ''} | FINALIZADO POR ETAPA - ${LLAMADAS_PARA_FINALIZAR} llamadas completadas automáticamente`.trim();

    const datosDestino = [
      datos[COLUMNAS.ID - 1] || '',
      nombreCompleto,
      datos[COLUMNAS.TELEFONO - 1] || '',
      datos[COLUMNAS.FORMACION - 1] || 'Barismo',
      datos[COLUMNAS.ALIADOS - 1] || '',
      datos[COLUMNAS.PLATAFORMAS - 1] || '',
      datos[COLUMNAS.CONEXION_LABORAL - 1] || '',
      datos[COLUMNAS.POR_SU_CUENTA - 1] || '',
      datos[COLUMNAS.NO_BUSCA_TRABAJAR - 1] || '',
      datos[COLUMNAS.EMPLEADO - 1] || '',
      datos[COLUMNAS.NO_TERMINO_FORMACION - 1] || '',
      'Finalizado',
      'Completado',
      datos[COLUMNAS.DOCUMENTOS - 1] || 'Ninguno',
      Utilities.formatDate(fechaActual, Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
      notasFinalizacion,
      nuevasLlamadas
    ];

    const resultadoInsercion = insertarEnHojaDestinoOptimizada(hojaDestino, datosDestino, '✅ Finalizados');

    if (!resultadoInsercion.exito) {
      throw new Error(`Error insertando en Finalizados: ${resultadoInsercion.error}`);
    }

    hojaGeneral.deleteRow(fila);

    const mensaje = `🎯 ${nombreCompleto} FINALIZADO automáticamente por cambio de etapa`;
    mostrarNotificacionDiscreta(mensaje);

    return { exito: true, error: null, movido: true, totalLlamadas: nuevasLlamadas };

  } catch (error) {
    console.error('❌ Error en procesamiento por cambio de etapa:', error);

    try {
      SpreadsheetApp.getActiveSpreadsheet().toast(
        `Error procesando finalización por etapa: ${error.message}`,
        'Error de Finalización',
        8
      );
    } catch (cleanupError) {
      console.error('❌ Error mostrando notificación:', cleanupError);
    }

    return { exito: false, error: error.message, movido: false, totalLlamadas: 0 };
  }
}

function finalizarParticipantesPorEtapa() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');

  if (!hojaGeneral) {
    ui.alert('❌ Sistema no configurado', 'Primero configura el sistema', ui.ButtonSet.OK);
    return;
  }

  const ultimaFila = hojaGeneral.getLastRow();
  if (ultimaFila <= 1) {
    ui.alert('ℹ️ Sin datos', 'La hoja está vacía', ui.ButtonSet.OK);
    return;
  }

  const participantesFinalizados = [];

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const etapa = hojaGeneral.getRange(fila, COLUMNAS.ETAPA_ACTUAL).getValue();
    const nombre = hojaGeneral.getRange(fila, COLUMNAS.NOMBRE).getValue();

    if (etapa && etapa.toString().toLowerCase().includes('finalizado') && nombre) {
      participantesFinalizados.push({
        fila: fila,
        nombre: nombre.toString().trim()
      });
    }
  }

  if (participantesFinalizados.length === 0) {
    ui.alert('ℹ️ Sin participantes',
      'No hay participantes con etapa "Finalizado" para procesar.',
      ui.ButtonSet.OK);
    return;
  }

  const confirmacion = ui.alert('🎯 Finalizar por Etapa',
    `Se encontraron ${participantesFinalizados.length} participantes con etapa "Finalizado"\n\n¿Continuar?`,
    ui.ButtonSet.YES_NO);

  if (confirmacion !== ui.Button.YES) return;

  let procesados = 0;
  let errores = 0;

  for (let i = participantesFinalizados.length - 1; i >= 0; i--) {
    const participante = participantesFinalizados[i];

    try {
      const resultado = procesarFinalizacionPorEtapa(hojaGeneral, participante.fila);

      if (resultado && resultado.exito) {
        procesados++;
      } else {
        errores++;
      }

      for (let j = 0; j < i; j++) {
        if (participantesFinalizados[j].fila > participante.fila) {
          participantesFinalizados[j].fila--;
        }
      }

    } catch (error) {
      errores++;
    }
  }

  let mensaje = `🎯 FINALIZACIÓN POR ETAPA COMPLETADA\n\n`;

  if (procesados > 0) {
    mensaje += `✅ ÉXITO: ${procesados} participantes finalizados\n`;
  }

  if (errores > 0) {
    mensaje += `❌ ERRORES: ${errores} problemas encontrados\n`;
  }

  ui.alert('🎯 Resultado', mensaje, ui.ButtonSet.OK);
}

// ====================================
// FUNCIONES DE PRUEBA
// ====================================
function procesarUnaLlamadaPrueba() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');

  if (!hojaGeneral) {
    ui.alert('❌ Sistema no configurado', 'Configura el sistema primero', ui.ButtonSet.OK);
    return;
  }

  const ultimaFila = hojaGeneral.getLastRow();
  if (ultimaFila <= 1) {
    ui.alert('ℹ️ Sin datos', 'Importa datos primero', ui.ButtonSet.OK);
    return;
  }

  let filaEncontrada = -1;
  let nombreParticipante = '';

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const checkbox = obtenerCheckboxSeguro(hojaGeneral, fila);
    if (checkbox === true) {
      filaEncontrada = fila;
      nombreParticipante = hojaGeneral.getRange(fila, COLUMNAS.NOMBRE).getValue() || 'Sin nombre';
      break;
    }
  }

  if (filaEncontrada === -1) {
    ui.alert('ℹ️ Sin checkboxes marcados',
      'Marca un checkbox e intenta nuevamente',
      ui.ButtonSet.OK);
    return;
  }

  const resultado = procesarLlamadaOptimizada(hojaGeneral, filaEncontrada);

  if (resultado.exito) {
    const mensaje = `✅ PRUEBA EXITOSA v2.8\n\n` +
      `👤 Participante: ${nombreParticipante}\n` +
      `📊 Resultado: ${resultado.movido ? `Movido a Finalizados` : `Procesado correctamente`}\n` +
      `📞 Llamadas: ${resultado.totalLlamadas}/${LLAMADAS_PARA_FINALIZAR}\n` +
      `📝 Información registrada en columnas de etapas`;

    ui.alert('🧪 Resultado de Prueba', mensaje, ui.ButtonSet.OK);
  } else {
    ui.alert('❌ Error en Prueba',
      `Participante: ${nombreParticipante}\nError: ${resultado.error}`,
      ui.ButtonSet.OK);
  }
}

function probarConexionSheetExterno() {
  const ui = SpreadsheetApp.getUi();

  try {
    const propiedades = PropertiesService.getScriptProperties();
    const url = propiedades.getProperty('URL_SHEET_EXTERNO');

    if (!url) {
      ui.alert('❌ No configurado',
        'Primero configura el Google Sheet externo usando "🔗 Configurar Google Sheet Externo"',
        ui.ButtonSet.OK);
      return;
    }

    const matches = url.match(/\/spreadsheets\/d\/([a-zA-Z0-9-_]+)/);
    if (!matches) {
      ui.alert('❌ URL inválida',
        'La URL configurada no es válida.',
        ui.ButtonSet.OK);
      return;
    }

    const sheetId = matches[1];
    const sheetExterno = SpreadsheetApp.openById(sheetId);

    const hojas = sheetExterno.getSheets();
    const nombresHojas = hojas.map(h => h.getName()).join('\n• ');

    ui.alert('✅ Conexión Exitosa',
      `Conectado a: "${sheetExterno.getName()}"\n\n📋 Hojas disponibles:\n• ${nombresHojas}`,
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Conexión',
      `No se pudo conectar al Google Sheet externo:\n\n${error.message}\n\nVerifica que:\n• La URL sea correcta\n• Tengas permisos de acceso\n• El documento exista`,
      ui.ButtonSet.OK);
  }
}

// ====================================
// REPORTES Y ESTADÍSTICAS
// ====================================
function generarReporte() {
  const ui = SpreadsheetApp.getUi();

  const estadisticas = recopilarEstadisticasDetalladas();
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  const propiedades = PropertiesService.getScriptProperties();
  const procesamientoActivo = propiedades.getProperty('PROCESAMIENTO_AUTOMATICO') === 'true';

  const reporte = `📈 REPORTE COMPLETO DEL SISTEMA v2.8\n` +
    `⏰ Generado el: ${fecha}\n\n` +

    `⚡ PROCESAMIENTO AUTOMÁTICO: ${procesamientoActivo ? '✅ ACTIVO' : '📋 INACTIVO'}\n` +
    `🔒 Sistema anti-doble procesamiento: ✅ ACTIVO\n` +
    `📝 Columnas de etapas: ✅ ACTIVAS\n` +
    `📞 Sistema configurado para ${LLAMADAS_PARA_FINALIZAR} llamadas\n` +
    `🆕 Formaciones: Análisis de datos E-commerce, SAC, Ofimática\n\n` +

    `📊 RESUMEN EJECUTIVO:\n` +
    `👥 Total de participantes: ${estadisticas.total}\n` +
    `📞 En proceso de llamadas: ${estadisticas.enProceso}\n` +
    `✅ Finalizados: ${estadisticas.finalizados}\n` +
    `📈 Tasa de finalización: ${estadisticas.total > 0 ? Math.round((estadisticas.finalizados / estadisticas.total) * 100) : 0}%\n\n` +

    `📞 DISTRIBUCIÓN POR LLAMADAS:\n` +
    `📋 Seguimiento inicial: ${estadisticas.seguimientoGeneral} participantes\n` +
    `📞 Primera llamada: ${estadisticas.llamada1} participantes\n` +
    `📞 Segunda llamada: ${estadisticas.llamada2} participantes\n` +
    `📞 Tercera llamada: ${estadisticas.llamada3} participantes\n` +
    `📞 Cuarta llamada: ${estadisticas.llamada4} participantes\n` +
    `📞 Quinta llamada: ${estadisticas.llamada5} participantes\n` +
    `✅ Completados: ${estadisticas.finalizados} participantes\n\n` +

    `🎓 DISTRIBUCIÓN POR FORMACIÓN:\n` +
    (estadisticas.formaciones.length > 0 ?
      estadisticas.formaciones.map(f => `• ${f.nombre}: ${f.cantidad} participantes`).join('\n') :
      'Sin datos de formación');

  ui.alert('📈 Reporte Completo v2.8', reporte, ui.ButtonSet.OK);
}

function recopilarEstadisticasDetalladas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const estadisticas = {
    total: 0,
    enProceso: 0,
    finalizados: 0,
    seguimientoGeneral: 0,
    llamada1: 0,
    llamada2: 0,
    llamada3: 0,
    llamada4: 0,
    llamada5: 0,
    formaciones: []
  };

  const hojasLlamadas = [
    { nombre: '📋 Seguimiento General', clave: 'seguimientoGeneral' },
    { nombre: '📞 Llamada 1', clave: 'llamada1' },
    { nombre: '📞 Llamada 2', clave: 'llamada2' },
    { nombre: '📞 Llamada 3', clave: 'llamada3' },
    { nombre: '📞 Llamada 4', clave: 'llamada4' },
    { nombre: '📞 Llamada 5', clave: 'llamada5' },
    { nombre: '✅ Finalizados', clave: 'finalizados' }
  ];

  const conteoFormaciones = {};

  hojasLlamadas.forEach(({ nombre, clave }) => {
    const hoja = ss.getSheetByName(nombre);
    if (!hoja) return;

    const ultimaFila = hoja.getLastRow();
    if (ultimaFila <= 1) return;

    const registros = ultimaFila - 1;
    estadisticas[clave] = registros;
    estadisticas.total += registros;

    if (clave !== 'finalizados') {
      estadisticas.enProceso += registros;
    }

    if (nombre === '📋 Seguimiento General') {
      for (let fila = 2; fila <= ultimaFila; fila++) {
        const formacion = hoja.getRange(fila, COLUMNAS.FORMACION).getValue();
        if (formacion) {
          conteoFormaciones[formacion] = (conteoFormaciones[formacion] || 0) + 1;
        }
      }
    }
  });

  estadisticas.formaciones = Object.entries(conteoFormaciones)
    .map(([nombre, cantidad]) => ({ nombre, cantidad }))
    .sort((a, b) => b.cantidad - a.cantidad);

  return estadisticas;
}

function diagnosticarSistema() {
  const ui = SpreadsheetApp.getUi();

  const diagnostico = realizarDiagnosticoCompleto();
  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  let reporte = `🔍 DIAGNÓSTICO COMPLETO v2.8\n⏰ ${fecha}\n\n`;

  reporte += `📊 ESTADO GENERAL:\n`;
  reporte += `• Nivel de salud: ${diagnostico.nivelSalud}\n`;
  reporte += `• Hojas configuradas: ${diagnostico.hojasConfiguradas}/${diagnostico.hojasRequeridas}\n`;
  reporte += `• Total participantes: ${diagnostico.totalParticipantes}\n`;
  reporte += `• Versión: 2.8 (Columnas de etapas + Anti-doble procesamiento)\n\n`;

  reporte += `⚡ PROCESAMIENTO AUTOMÁTICO:\n`;
  reporte += `• Estado: ${diagnostico.procesamientoActivo ? '✅ ACTIVO' : '📋 INACTIVO'}\n`;
  reporte += `• Checkboxes marcados: ${diagnostico.checkboxesMarcados}\n\n`;

  if (diagnostico.problemas.length > 0) {
    reporte += `❌ PROBLEMAS DETECTADOS:\n`;
    diagnostico.problemas.forEach(problema => {
      reporte += `• ${problema}\n`;
    });
    reporte += `\n`;
  }

  if (diagnostico.recomendaciones.length > 0) {
    reporte += `💡 RECOMENDACIONES:\n`;
    diagnostico.recomendaciones.forEach(recomendacion => {
      reporte += `• ${recomendacion}\n`;
    });
  }

  if (diagnostico.problemas.length === 0) {
    reporte += `✅ SISTEMA SALUDABLE\nNo se detectaron problemas críticos`;
  }

  ui.alert('🔍 Diagnóstico v2.8', reporte, ui.ButtonSet.OK);
}

function realizarDiagnosticoCompleto() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const propiedades = PropertiesService.getScriptProperties();

  const diagnostico = {
    nivelSalud: 'Desconocido',
    hojasConfiguradas: 0,
    hojasRequeridas: 7,
    totalParticipantes: 0,
    procesamientoActivo: false,
    checkboxesMarcados: 0,
    problemas: [],
    recomendaciones: []
  };

  try {
    const hojasRequeridas = [
      '📋 Seguimiento General',
      '📞 Llamada 1', '📞 Llamada 2', '📞 Llamada 3',
      '📞 Llamada 4', '📞 Llamada 5',
      '✅ Finalizados'
    ];

    let hojasExistentes = 0;
    hojasRequeridas.forEach(nombre => {
      if (ss.getSheetByName(nombre)) {
        hojasExistentes++;
      } else {
        diagnostico.problemas.push(`Falta hoja: ${nombre}`);
      }
    });

    diagnostico.hojasConfiguradas = hojasExistentes;
    diagnostico.procesamientoActivo = propiedades.getProperty('PROCESAMIENTO_AUTOMATICO') === 'true';

    const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');
    if (hojaGeneral) {
      const ultimaFila = hojaGeneral.getLastRow();
      diagnostico.totalParticipantes = Math.max(0, ultimaFila - 1);

      // Verificar si tiene las nuevas columnas
      if (ultimaFila > 0) {
        const encabezados = hojaGeneral.getRange(1, 1, 1, hojaGeneral.getLastColumn()).getValues()[0];
        if (encabezados.length < TOTAL_COLUMNAS || encabezados[4] !== 'Aliados') {
          diagnostico.problemas.push('Estructura desactualizada: Faltan columnas de etapas');
          diagnostico.recomendaciones.push('Ejecutar "ACTUALIZAR Estructura de Tabla"');
        }
      }

      for (let fila = 2; fila <= ultimaFila; fila++) {
        try {
          if (obtenerCheckboxSeguro(hojaGeneral, fila)) {
            diagnostico.checkboxesMarcados++;
          }
        } catch (error) {
          diagnostico.problemas.push('Error verificando checkboxes');
          break;
        }
      }
    } else {
      diagnostico.problemas.push('Hoja principal no existe');
    }

    const urlExterno = propiedades.getProperty('URL_SHEET_EXTERNO');
    if (!urlExterno) {
      diagnostico.recomendaciones.push('Configurar Google Sheet externo');
    }

    if (diagnostico.problemas.length === 0) {
      diagnostico.nivelSalud = '✅ EXCELENTE';
    } else if (diagnostico.problemas.length <= 2) {
      diagnostico.nivelSalud = '⚠️ BUENO CON OBSERVACIONES';
    } else {
      diagnostico.nivelSalud = '❌ REQUIERE ATENCIÓN';
      diagnostico.recomendaciones.push('Ejecutar reparación del sistema');
    }

    if (!diagnostico.procesamientoActivo && diagnostico.checkboxesMarcados > 0) {
      diagnostico.recomendaciones.push('Activar procesamiento automático');
    }

    if (diagnostico.totalParticipantes === 0) {
      diagnostico.recomendaciones.push('Importar datos de participantes');
    }

  } catch (error) {
    diagnostico.problemas.push(`Error en diagnóstico: ${error.message}`);
    diagnostico.nivelSalud = '❌ ERROR CRÍTICO';
  }

  return diagnostico;
}

function verificarCheckboxes() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');

  if (!hojaGeneral) {
    ui.alert('❌ Hoja no encontrada', 'Configura el sistema primero', ui.ButtonSet.OK);
    return;
  }

  const ultimaFila = hojaGeneral.getLastRow();
  if (ultimaFila <= 1) {
    ui.alert('ℹ️ Sin datos', 'No hay participantes', ui.ButtonSet.OK);
    return;
  }

  let checkboxesMarcados = 0;
  let totalParticipantes = 0;
  const participantesMarcados = [];

  const propiedades = PropertiesService.getScriptProperties();
  const procesamientoActivo = propiedades.getProperty('PROCESAMIENTO_AUTOMATICO') === 'true';

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const nombre = hojaGeneral.getRange(fila, COLUMNAS.NOMBRE).getValue();

    if (nombre && nombre.toString().trim() !== '') {
      totalParticipantes++;
      const checkbox = obtenerCheckboxSeguro(hojaGeneral, fila);
      const llamadas = hojaGeneral.getRange(fila, COLUMNAS.TOTAL_LLAMADAS).getValue() || 0;

      if (checkbox === true) {
        checkboxesMarcados++;
        participantesMarcados.push({
          fila: fila,
          nombre: nombre.toString().trim(),
          llamadas: llamadas
        });
      }
    }
  }

  let mensaje = `🔍 VERIFICACIÓN DE CHECKBOXES v2.8\n\n`;
  mensaje += `⚡ PROCESAMIENTO AUTOMÁTICO: ${procesamientoActivo ? '✅ ACTIVO' : '📋 INACTIVO'}\n`;
  mensaje += `🔒 Sistema anti-doble procesamiento: ✅ ACTIVO\n`;
  mensaje += `📝 Columnas de etapas: ✅ ACTIVAS\n`;
  mensaje += `📞 Sistema configurado para ${LLAMADAS_PARA_FINALIZAR} llamadas\n\n`;

  mensaje += `📊 RESUMEN:\n`;
  mensaje += `👥 Total participantes: ${totalParticipantes}\n`;
  mensaje += `✅ Checkboxes marcados: ${checkboxesMarcados}\n`;
  mensaje += `⚪ Sin marcar: ${totalParticipantes - checkboxesMarcados}\n\n`;

  if (checkboxesMarcados > 0) {
    mensaje += `📋 LISTOS PARA PROCESAR:\n`;
    participantesMarcados.slice(0, 10).forEach(p => {
      mensaje += `• Fila ${p.fila}: ${p.nombre} (${p.llamadas}/${LLAMADAS_PARA_FINALIZAR} llamadas)\n`;
    });

    if (participantesMarcados.length > 10) {
      mensaje += `... y ${participantesMarcados.length - 10} participantes más\n`;
    }
  } else {
    mensaje += `ℹ️ No hay checkboxes marcados para procesar`;
  }

  ui.alert('🔍 Estado de Checkboxes v2.8', mensaje, ui.ButtonSet.OK);
}

// Continúa en siguiente archivo...
// ====================================
// CONTINUACIÓN DEL CÓDIGO - PARTE 4 (FINAL)
// ====================================

// ====================================
// MANTENIMIENTO
// ====================================
function repararSistema() {
  const ui = SpreadsheetApp.getUi();

  const confirmacion = ui.alert('🔧 Reparar Sistema',
    'Esta función reparará el sistema completo. ¿Continuar?',
    ui.ButtonSet.YES_NO);

  if (confirmacion !== ui.Button.YES) return;

  try {
    limpiarBloqueosProcesamiento();
    configurarHojasCorregido();
    aplicarFormato();

    ui.alert('✅ Reparación Completada',
      'Sistema v2.8 reparado exitosamente!\n\n🔒 Bloqueos limpiados\n✅ Sistema anti-doble procesamiento activo\n📝 Columnas de etapas configuradas',
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Reparación',
      `Error durante la reparación: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function limpiarDatosVacios() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const confirmacion = ui.alert('🧹 Limpiar Datos Vacíos',
    'Esta función eliminará filas sin datos válidos. ¿Continuar?',
    ui.ButtonSet.YES_NO);

  if (confirmacion !== ui.Button.YES) return;

  try {
    let totalLimpiados = 0;

    const hojasALimpiar = [
      '📋 Seguimiento General',
      '📞 Llamada 1', '📞 Llamada 2', '📞 Llamada 3',
      '📞 Llamada 4', '📞 Llamada 5',
      '✅ Finalizados'
    ];

    hojasALimpiar.forEach(nombreHoja => {
      const hoja = ss.getSheetByName(nombreHoja);
      if (!hoja) return;

      const ultimaFila = hoja.getLastRow();
      if (ultimaFila <= 1) return;

      let filasEliminadas = 0;

      for (let fila = ultimaFila; fila >= 2; fila--) {
        const nombre = hoja.getRange(fila, COLUMNAS.NOMBRE).getValue();

        if (!nombre || nombre.toString().trim() === '' || nombre.toString().trim().length < 2) {
          hoja.deleteRow(fila);
          filasEliminadas++;
        }
      }

      totalLimpiados += filasEliminadas;
    });

    ui.alert('🧹 Limpieza Completada',
      `${totalLimpiados} registros vacíos eliminados`,
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Limpieza',
      `Error durante la limpieza: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function crearRespaldo() {
  const ui = SpreadsheetApp.getUi();

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();
    const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'yyyy-MM-dd_HH-mm');
    const nombreRespaldo = `RESPALDO_v2.8_${ss.getName()}_${fecha}`;

    ui.alert('💾 Creando Respaldo...',
      'Creando copia de seguridad...',
      ui.ButtonSet.OK);

    const hojasRespaldo = [];
    const hojasRequeridas = [
      '📋 Seguimiento General',
      '📞 Llamada 1', '📞 Llamada 2', '📞 Llamada 3',
      '📞 Llamada 4', '📞 Llamada 5',
      '✅ Finalizados'
    ];

    hojasRequeridas.forEach(nombre => {
      const hoja = ss.getSheetByName(nombre);
      if (hoja) {
        hojasRespaldo.push({
          nombre: nombre,
          datos: hoja.getDataRange().getValues()
        });
      }
    });

    const nuevoSS = SpreadsheetApp.create(nombreRespaldo);

    hojasRespaldo.forEach(({ nombre, datos }) => {
      const nuevaHoja = nuevoSS.insertSheet(nombre);
      if (datos.length > 0) {
        nuevaHoja.getRange(1, 1, datos.length, datos[0].length).setValues(datos);
      }
    });

    const hojaDefecto = nuevoSS.getSheetByName('Hoja 1');
    if (hojaDefecto && nuevoSS.getSheets().length > 1) {
      nuevoSS.deleteSheet(hojaDefecto);
    }

    const urlRespaldo = nuevoSS.getUrl();

    ui.alert('💾 Respaldo Creado',
      `✅ Respaldo v2.8 creado exitosamente!\n\nNombre: ${nombreRespaldo}\nURL: ${urlRespaldo}`,
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Respaldo',
      `Error creando respaldo: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function resetearSistema() {
  const ui = SpreadsheetApp.getUi();

  const confirmacion = ui.alert('🔄 RESETEAR Sistema Completo',
    '⚠️ ADVERTENCIA: Esta acción eliminará TODOS los datos. ¿Continuar?',
    ui.ButtonSet.YES_NO);

  if (confirmacion !== ui.Button.YES) return;

  const confirmaFinal = ui.alert('🔴 CONFIRMACIÓN FINAL',
    'TODOS los datos se eliminarán permanentemente. ¿Confirmar RESET?',
    ui.ButtonSet.YES_NO);

  if (confirmaFinal !== ui.Button.YES) return;

  try {
    const ss = SpreadsheetApp.getActiveSpreadsheet();

    const hojasAEliminar = [
      '📋 Seguimiento General',
      '📞 Llamada 1', '📞 Llamada 2', '📞 Llamada 3',
      '📞 Llamada 4', '📞 Llamada 5',
      '✅ Finalizados'
    ];

    hojasAEliminar.forEach(nombre => {
      const hoja = ss.getSheetByName(nombre);
      if (hoja) {
        ss.deleteSheet(hoja);
      }
    });

    const propiedades = PropertiesService.getScriptProperties();
    propiedades.deleteProperty('PROCESAMIENTO_AUTOMATICO');
    propiedades.deleteProperty('URL_SHEET_EXTERNO');
    propiedades.deleteProperty('FECHA_ACTIVACION_AUTO');
    propiedades.deleteProperty('FECHA_DESACTIVACION_AUTO');

    limpiarBloqueosProcesamiento();
    removerTriggerOnEdit();
    configurarHojasCorregido();

    ui.alert('🔄 Reset Completado',
      '✅ Sistema v2.8 reseteado completamente!\n\nSistema listo para nueva configuración.\n🔒 Bloqueos limpiados\n📝 Columnas de etapas configuradas\n🆕 Formaciones: Análisis de datos E-commerce, SAC, Ofimática',
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Reset',
      `Error durante el reset: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function limpiarValidacionesPorRangos(hoja) {
  try {
    const rangosConValidaciones = [
      hoja.getRange(2, COLUMNAS.FORMACION, 998, 1),
      hoja.getRange(2, COLUMNAS.ETAPA_ACTUAL, 998, 1),
      hoja.getRange(2, COLUMNAS.DOCUMENTOS, 998, 1),
      hoja.getRange(2, COLUMNAS.PROCESAR, 998, 1)
    ];

    rangosConValidaciones.forEach(rango => {
      try {
        rango.clearDataValidations();
      } catch (e) {
        rango.setDataValidation(null);
      }
    });

  } catch (error) {
    console.warn('Error limpiando validaciones:', error.message);
  }
}

// ====================================
// FORMATO Y PRESENTACIÓN
// ====================================
function aplicarFormato() {
  const ui = SpreadsheetApp.getUi();

  try {
    ui.alert('🎨 Aplicando Diseño...',
      'Aplicando formato profesional v2.8...',
      ui.ButtonSet.OK);

    aplicarFormatoHojaPrincipal();
    aplicarFormatoHojasLlamadas();

    ui.alert('🎨 Diseño Aplicado',
      '✅ Formato profesional v2.8 aplicado exitosamente!',
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Formato',
      `Error aplicando formato: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function aplicarFormatoHojaPrincipal() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hoja = ss.getSheetByName('📋 Seguimiento General');

  if (!hoja) return;

  const rangoEncabezado = hoja.getRange(1, 1, 1, TOTAL_COLUMNAS);
  rangoEncabezado.setBackground('#1f4e79');
  rangoEncabezado.setFontColor('#ffffff');
  rangoEncabezado.setFontWeight('bold');
  rangoEncabezado.setFontSize(11);
  rangoEncabezado.setHorizontalAlignment('center');

  // Ajustar anchos de columnas
  hoja.setColumnWidth(COLUMNAS.ID, 100);
  hoja.setColumnWidth(COLUMNAS.NOMBRE, 180);
  hoja.setColumnWidth(COLUMNAS.TELEFONO, 120);
  hoja.setColumnWidth(COLUMNAS.FORMACION, 140);
  // Columnas de etapas
  hoja.setColumnWidth(COLUMNAS.ALIADOS, 150);
  hoja.setColumnWidth(COLUMNAS.PLATAFORMAS, 150);
  hoja.setColumnWidth(COLUMNAS.CONEXION_LABORAL, 150);
  hoja.setColumnWidth(COLUMNAS.POR_SU_CUENTA, 150);
  hoja.setColumnWidth(COLUMNAS.NO_BUSCA_TRABAJAR, 150);
  hoja.setColumnWidth(COLUMNAS.EMPLEADO, 150);
  hoja.setColumnWidth(COLUMNAS.NO_TERMINO_FORMACION, 150);
  // Columnas restantes
  hoja.setColumnWidth(COLUMNAS.ETAPA_ACTUAL, 120);
  hoja.setColumnWidth(COLUMNAS.RESULTADOS, 150);
  hoja.setColumnWidth(COLUMNAS.DOCUMENTOS, 150);
  hoja.setColumnWidth(COLUMNAS.FECHA, 120);
  hoja.setColumnWidth(COLUMNAS.NOTAS, 200);
  hoja.setColumnWidth(COLUMNAS.TOTAL_LLAMADAS, 80);
  hoja.setColumnWidth(COLUMNAS.PROCESAR, 80);

  hoja.setFrozenRows(1);
}

function aplicarFormatoHojasLlamadas() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const configuracionColores = {
    '📞 Llamada 1': { fondo: '#1976d2', texto: '#ffffff' },
    '📞 Llamada 2': { fondo: '#7b1fa2', texto: '#ffffff' },
    '📞 Llamada 3': { fondo: '#f57c00', texto: '#ffffff' },
    '📞 Llamada 4': { fondo: '#ff8f00', texto: '#ffffff' },
    '📞 Llamada 5': { fondo: '#c2185b', texto: '#ffffff' },
    '✅ Finalizados': { fondo: '#388e3c', texto: '#ffffff' }
  };

  Object.entries(configuracionColores).forEach(([nombreHoja, colores]) => {
    const hoja = ss.getSheetByName(nombreHoja);
    if (!hoja) return;

    const rangoEncabezado = hoja.getRange(1, 1, 1, TOTAL_COLUMNAS_DESTINO);
    rangoEncabezado.setBackground(colores.fondo);
    rangoEncabezado.setFontColor(colores.texto);
    rangoEncabezado.setFontWeight('bold');
    rangoEncabezado.setFontSize(11);
    rangoEncabezado.setHorizontalAlignment('center');

    // Ajustar anchos
    hoja.setColumnWidth(1, 100); // ID
    hoja.setColumnWidth(2, 180); // Nombre
    hoja.setColumnWidth(3, 120); // Teléfono
    hoja.setColumnWidth(4, 140); // Formación
    // Columnas de etapas
    for (let col = 5; col <= 11; col++) {
      hoja.setColumnWidth(col, 150);
    }
    // Columnas restantes
    hoja.setColumnWidth(12, 120); // Etapa Actual
    hoja.setColumnWidth(13, 140); // Resultados
    hoja.setColumnWidth(14, 150); // Documentos
    hoja.setColumnWidth(15, 140); // Fecha
    hoja.setColumnWidth(16, 200); // Notas
    hoja.setColumnWidth(17, 100); // Total Llamadas

    hoja.setFrozenRows(1);
  });
}

// ====================================
// GESTIÓN DE ETAPAS
// ====================================
function configurarDesplegableEtapaActual() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');

  if (!hojaGeneral) {
    ui.alert('❌ Hoja no encontrada', 'Configura el sistema primero', ui.ButtonSet.OK);
    return;
  }

  try {
    const ultimaFila = hojaGeneral.getLastRow();
    if (ultimaFila <= 1) {
      ui.alert('ℹ️ Sin datos', 'No hay participantes', ui.ButtonSet.OK);
      return;
    }

    const opcionesEtapa = [
      'Inicial', 'Aliados', 'Plataformas', 'Conexión Laboral',
      'Por su cuenta', 'No busca trabajar', 'Empleado',
      'No termino la formación', 'Finalizado'
    ];

    const rangoEtapa = hojaGeneral.getRange(2, COLUMNAS.ETAPA_ACTUAL, ultimaFila - 1, 1);
    const validacionEtapa = SpreadsheetApp.newDataValidation()
      .requireValueInList(opcionesEtapa)
      .setAllowInvalid(true)
      .build();

    rangoEtapa.setDataValidation(validacionEtapa);

    ui.alert('✅ Etapas Configuradas',
      'Desplegable de etapas configurado correctamente',
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Configuración',
      `Error: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function actualizarEtapaSegunLlamadas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const confirmacion = ui.alert('🔄 Actualizar Etapas',
    'Esta función actualizará las etapas basándose en el número de llamadas. ¿Continuar?',
    ui.ButtonSet.YES_NO);

  if (confirmacion !== ui.Button.YES) return;

  try {
    const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');
    if (!hojaGeneral) {
      ui.alert('❌ Hoja no encontrada', 'Sistema no configurado', ui.ButtonSet.OK);
      return;
    }

    const ultimaFila = hojaGeneral.getLastRow();
    if (ultimaFila <= 1) {
      ui.alert('ℹ️ Sin datos', 'No hay participantes para actualizar', ui.ButtonSet.OK);
      return;
    }

    let actualizados = 0;

    for (let fila = 2; fila <= ultimaFila; fila++) {
      const llamadas = parseInt(hojaGeneral.getRange(fila, COLUMNAS.TOTAL_LLAMADAS).getValue()) || 0;
      let nuevaEtapa = '';

      if (llamadas === 0) {
        nuevaEtapa = 'Inicial';
      } else if (llamadas <= 2) {
        nuevaEtapa = 'En proceso';
      } else if (llamadas <= 4) {
        nuevaEtapa = 'Conexión Laboral';
      } else if (llamadas < LLAMADAS_PARA_FINALIZAR) {
        nuevaEtapa = 'Por su cuenta';
      } else {
        nuevaEtapa = 'Finalizado';
      }

      const etapaActual = hojaGeneral.getRange(fila, COLUMNAS.ETAPA_ACTUAL).getValue();
      if (etapaActual !== nuevaEtapa) {
        hojaGeneral.getRange(fila, COLUMNAS.ETAPA_ACTUAL).setValue(nuevaEtapa);
        actualizados++;
      }
    }

    ui.alert('✅ Etapas Actualizadas',
      `Actualización completada!\n\n🔄 ${actualizados} participantes actualizados\n📊 Etapas basadas en número de llamadas`,
      ui.ButtonSet.OK);

  } catch (error) {
    ui.alert('❌ Error de Actualización',
      `Error actualizando etapas: ${error.message}`,
      ui.ButtonSet.OK);
  }
}

function generarReporteEtapas() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');
  if (!hojaGeneral) {
    ui.alert('❌ Sistema no configurado', 'Configura el sistema primero', ui.ButtonSet.OK);
    return;
  }

  const ultimaFila = hojaGeneral.getLastRow();
  if (ultimaFila <= 1) {
    ui.alert('ℹ️ Sin datos', 'No hay participantes para analizar', ui.ButtonSet.OK);
    return;
  }

  const conteoEtapas = {};
  let totalParticipantes = 0;

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const etapa = hojaGeneral.getRange(fila, COLUMNAS.ETAPA_ACTUAL).getValue() || 'Sin etapa';
    conteoEtapas[etapa] = (conteoEtapas[etapa] || 0) + 1;
    totalParticipantes++;
  }

  const etapasOrdenadas = Object.entries(conteoEtapas)
    .sort((a, b) => b[1] - a[1]);

  const fecha = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');

  let reporte = `📊 REPORTE DE ETAPAS v2.8\n⏰ ${fecha}\n\n`;
  reporte += `👥 Total participantes: ${totalParticipantes}\n\n`;
  reporte += `📈 DISTRIBUCIÓN POR ETAPAS:\n`;

  etapasOrdenadas.forEach(([etapa, cantidad]) => {
    const porcentaje = Math.round((cantidad / totalParticipantes) * 100);
    reporte += `• ${etapa}: ${cantidad} participantes (${porcentaje}%)\n`;
  });

  reporte += `\n💡 ANÁLISIS:\n`;

  if (conteoEtapas['Inicial'] > totalParticipantes * 0.4) {
    reporte += `• Alto número de participantes en etapa inicial\n`;
  }

  if (conteoEtapas['Finalizado'] > totalParticipantes * 0.3) {
    reporte += `• Buena tasa de finalización\n`;
  }

  if (conteoEtapas['Conexión Laboral'] > totalParticipantes * 0.2) {
    reporte += `• Buen progreso en conexión laboral\n`;
  }

  ui.alert('📊 Reporte de Etapas', reporte, ui.ButtonSet.OK);
}

function buscarPorEtapa() {
  const ui = SpreadsheetApp.getUi();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const respuesta = ui.prompt(
    '🔍 Buscar por Etapa',
    'Escribe la etapa a buscar:',
    ui.ButtonSet.OK_CANCEL
  );

  if (respuesta.getSelectedButton() !== ui.Button.OK) return;

  const etapaBuscada = respuesta.getResponseText().trim();
  if (!etapaBuscada) {
    ui.alert('❌ Etapa requerida', 'Debes especificar una etapa', ui.ButtonSet.OK);
    return;
  }

  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');
  if (!hojaGeneral) {
    ui.alert('❌ Sistema no configurado', 'Configura el sistema primero', ui.ButtonSet.OK);
    return;
  }

  const ultimaFila = hojaGeneral.getLastRow();
  if (ultimaFila <= 1) {
    ui.alert('ℹ️ Sin datos', 'No hay participantes', ui.ButtonSet.OK);
    return;
  }

  const participantesEncontrados = [];

  for (let fila = 2; fila <= ultimaFila; fila++) {
    const etapa = hojaGeneral.getRange(fila, COLUMNAS.ETAPA_ACTUAL).getValue();
    const nombre = hojaGeneral.getRange(fila, COLUMNAS.NOMBRE).getValue();
    const llamadas = hojaGeneral.getRange(fila, COLUMNAS.TOTAL_LLAMADAS).getValue() || 0;

    if (etapa && etapa.toString().toLowerCase().includes(etapaBuscada.toLowerCase())) {
      participantesEncontrados.push({
        fila: fila,
        nombre: nombre,
        etapa: etapa,
        llamadas: llamadas
      });
    }
  }

  if (participantesEncontrados.length === 0) {
    ui.alert('🔍 Sin resultados',
      `No se encontraron participantes en etapa: "${etapaBuscada}"`,
      ui.ButtonSet.OK);
    return;
  }

  let resultado = `🔍 RESULTADOS DE BÚSQUEDA\n`;
  resultado += `🎯 Etapa: "${etapaBuscada}"\n`;
  resultado += `📊 Encontrados: ${participantesEncontrados.length} participantes\n\n`;

  participantesEncontrados.slice(0, 15).forEach(p => {
    resultado += `• Fila ${p.fila}: ${p.nombre} (${p.llamadas} llamadas)\n`;
  });

  if (participantesEncontrados.length > 15) {
    resultado += `... y ${participantesEncontrados.length - 15} participantes más\n`;
  }

  ui.alert('🔍 Resultados de Búsqueda', resultado, ui.ButtonSet.OK);
}

// ====================================
// CONFIGURACIÓN Y AYUDA
// ====================================
function verConfiguracion() {
  const ui = SpreadsheetApp.getUi();
  const propiedades = PropertiesService.getScriptProperties();
  const ss = SpreadsheetApp.getActiveSpreadsheet();

  const urlExterno = propiedades.getProperty('URL_SHEET_EXTERNO') || 'No configurada';
  const procesamientoActivo = propiedades.getProperty('PROCESAMIENTO_AUTOMATICO') === 'true';
  const fechaActivacion = propiedades.getProperty('FECHA_ACTIVACION_AUTO');

  const hojasRequeridas = [
    '📋 Seguimiento General',
    '📞 Llamada 1', '📞 Llamada 2', '📞 Llamada 3',
    '📞 Llamada 4', '📞 Llamada 5',
    '✅ Finalizados'
  ];

  const hojasExistentes = [];
  const hojasFaltantes = [];

  hojasRequeridas.forEach(nombre => {
    if (ss.getSheetByName(nombre)) {
      hojasExistentes.push(nombre);
    } else {
      hojasFaltantes.push(nombre);
    }
  });

  const hojaGeneral = ss.getSheetByName('📋 Seguimiento General');
  const totalParticipantes = hojaGeneral ? Math.max(0, hojaGeneral.getLastRow() - 1) : 0;

  let configuracion = `⚙️ CONFIGURACIÓN ACTUAL v2.8\n`;
  configuracion += `⏰ ${Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm')}\n\n`;

  configuracion += `📊 ESTADO DEL SISTEMA:\n`;
  configuracion += `• Versión: 2.8 (Columnas de etapas + Anti-doble procesamiento)\n`;
  configuracion += `• Hojas configuradas: ${hojasExistentes.length}/${hojasRequeridas.length}\n`;
  configuracion += `• Participantes registrados: ${totalParticipantes}\n`;
  configuracion += `• Llamadas configuradas: ${LLAMADAS_PARA_FINALIZAR}\n\n`;

  configuracion += `🆕 NUEVAS FORMACIONES v2.8:\n`;
  configuracion += `• Análisis de datos E-commerce ✅\n`;
  configuracion += `• SAC ✅\n`;
  configuracion += `• Ofimática ✅\n\n`;

  configuracion += `📝 COLUMNAS DE ETAPAS v2.8:\n`;
  configuracion += `• Aliados ✅\n`;
  configuracion += `• Plataformas ✅\n`;
  configuracion += `• Conexión laboral ✅\n`;
  configuracion += `• Por su cuenta ✅\n`;
  configuracion += `• No busca trabajar ✅\n`;
  configuracion += `• Empleado ✅\n`;
  configuracion += `• No terminó la formación ✅\n\n`;

  configuracion += `⚡ PROCESAMIENTO AUTOMÁTICO:\n`;
  configuracion += `• Estado: ${procesamientoActivo ? '✅ ACTIVO' : '📋 INACTIVO'}\n`;
  configuracion += `• Sistema anti-doble procesamiento: ✅ ACTIVO\n`;
  if (fechaActivacion) {
    configuracion += `• Última activación: ${new Date(fechaActivacion).toLocaleDateString()}\n`;
  }

  configuracion += `\n🔗 GOOGLE SHEET EXTERNO:\n`;
  if (urlExterno === 'No configurada') {
    configuracion += `• ❌ No configurado\n`;
  } else {
    configuracion += `• ✅ Configurado\n`;
  }

  if (hojasFaltantes.length > 0) {
    configuracion += `\n❌ HOJAS FALTANTES:\n`;
    hojasFaltantes.forEach(nombre => {
      configuracion += `• ${nombre}\n`;
    });
    configuracion += `\n💡 Usa "🔧 Configurar Sistema" para crearlas\n`;
  }

  configuracion += `\n🎯 ESTADO GENERAL:\n`;
  if (hojasFaltantes.length === 0 && urlExterno !== 'No configurada') {
    configuracion += `✅ Sistema completamente configurado\n`;
    configuracion += `🚀 Listo para uso completo\n`;
  } else {
    configuracion += `⚠️ Configuración incompleta\n`;
    configuracion += `🔧 Requiere configuración adicional\n`;
  }

  ui.alert('⚙️ Configuración Actual v2.8', configuracion, ui.ButtonSet.OK);
}

function mostrarHojaInstrucciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('📘 Instrucciones');

  if (!hoja) {
    crearHojaInstrucciones();
    hoja = ss.getSheetByName('📘 Instrucciones');
  }

  ss.setActiveSheet(hoja);
}

function crearHojaInstrucciones() {
  const ss = SpreadsheetApp.getActiveSpreadsheet();
  let hoja = ss.getSheetByName('📘 Instrucciones');

  if (!hoja) {
    hoja = ss.insertSheet('📘 Instrucciones');
  } else {
    hoja.clear();
  }

  const instruccionesTexto = [
    ['🚀 SISTEMA DE SEGUIMIENTO DE PRÁCTICAS v2.8'],
    [''],
    ['Bienvenido al sistema avanzado para gestionar el seguimiento de participantes.'],
    [''],
    ['✨ FUNCIONALIDADES PRINCIPALES:'],
    [`  • Procesamiento AUTOMÁTICO al marcar checkboxes (hasta ${LLAMADAS_PARA_FINALIZAR} llamadas)`],
    ['  • 🔒 Sistema anti-doble procesamiento ACTIVO'],
    ['  • 📝 Columnas de etapas para registro detallado'],
    ['  • Importación inteligente desde Google Sheets externos'],
    ['  • Detección automática de formaciones'],
    ['  • 🆕 NUEVAS FORMACIONES: Análisis de datos E-commerce, SAC, Ofimática'],
    ['  • Reportes detallados y análisis en tiempo real'],
    ['  • Gestión dinámica de etapas y estados'],
    ['  • Selección múltiple de documentos faltantes'],
    [''],
    ['📋 PASOS PARA INICIAR:'],
    ['  1️⃣ Ejecuta "⚙️ Configuración Inicial"'],
    ['  2️⃣ Si ya tienes datos, ejecuta "🔄 ACTUALIZAR Estructura de Tabla"'],
    ['  3️⃣ Configura el Google Sheet externo con "🔗 Configurar Google Sheet Externo"'],
    ['  4️⃣ Importa datos con "📥 Importar Datos"'],
    ['  5️⃣ Activa procesamiento automático con "⚡ ACTIVAR Procesamiento Automático"'],
    ['  6️⃣ Marca checkboxes para procesar llamadas automáticamente'],
    ['  7️⃣ Espera 1-2 segundos entre cada checkbox (sistema anti-doble procesamiento)'],
    [''],
    ['🆕 NOVEDADES v2.8:'],
    ['  • 📝 7 COLUMNAS NUEVAS de etapas para registro detallado'],
    ['  • Escritura automática en columnas según etapa del participante'],
    ['  • Función "ACTUALIZAR Estructura de Tabla" sin pérdida de datos'],
    ['  • 🔒 Sistema anti-doble procesamiento mejorado'],
    ['  • Análisis de datos E-commerce, SAC y Ofimática como formaciones'],
    [''],
    ['📝 COLUMNAS DE ETAPAS:'],
    ['  Al procesar cada llamada, se registra información en:'],
    ['  • Aliados - Registro de conexiones con aliados'],
    ['  • Plataformas - Información sobre plataformas de empleo'],
    ['  • Conexión laboral - Detalles de conexiones laborales'],
    ['  • Por su cuenta - Información de búsqueda independiente'],
    ['  • No busca trabajar - Registro si no busca empleo'],
    ['  • Empleado - Información de empleos conseguidos'],
    ['  • No terminó la formación - Registro de no completados'],
    [''],
    ['🎓 FORMACIONES DISPONIBLES:'],
    ['  • Barista (I, II, III, IV)'],
    ['  • Barismo (1-10)'],
    ['  • Gastronomía (I-V o 1-5)'],
    ['  • Food Manager'],
    ['  • Panadería, Repostería, Sommelier'],
    ['  • 🆕 Análisis de datos E-commerce'],
    ['  • 🆕 SAC (Servicio al Cliente)'],
    ['  • 🆕 Ofimática'],
    [''],
    ['⚠️ IMPORTANTE - EVITAR DOBLE PROCESAMIENTO:'],
    ['  • Espera 1-2 segundos entre cada checkbox marcado'],
    ['  • Si ves doble procesamiento, usa "Diagnóstico de Doble Procesamiento"'],
    ['  • Limpia bloqueos con "Limpiar Bloqueos de Procesamiento"'],
    ['  • Si persiste, desactiva y reactiva el procesamiento automático'],
    [''],
    ['💡 CONSEJOS:'],
    ['  • Usa el procesamiento automático para mayor eficiencia'],
    ['  • Las columnas de etapas se llenan automáticamente según la etapa actual'],
    ['  • Realiza respaldos periódicos'],
    ['  • Revisa reportes regularmente'],
    ['  • Las nuevas formaciones se detectan automáticamente al importar'],
    [''],
    ['🔧 SOLUCIÓN DE PROBLEMAS:'],
    ['  • Falta estructura nueva: Usa "ACTUALIZAR Estructura de Tabla"'],
    ['  • Doble procesamiento: Usa "Diagnóstico de Doble Procesamiento"'],
    ['  • Bloqueos atorados: Usa "Limpiar Bloqueos de Procesamiento"'],
    ['  • Sistema lento: Desactiva y reactiva procesamiento automático'],
    ['  • Errores persistentes: Usa "Reparar Sistema"'],
    [''],
    ['🎉 ¡Sistema v2.8 con columnas de etapas y protección anti-doble procesamiento!']
  ];

  hoja.getRange(1, 1, instruccionesTexto.length, 1).setValues(instruccionesTexto);
  hoja.setColumnWidth(1, 800);

  const rangoCompleto = hoja.getRange(1, 1, instruccionesTexto.length, 1);
  rangoCompleto.setFontFamily('Calibri');
  rangoCompleto.setFontSize(11);
  rangoCompleto.setWrap(true);

  const titulo = hoja.getRange(1, 1);
  titulo.setFontWeight('bold').setFontSize(18).setFontColor('#00bcd4');
  titulo.setHorizontalAlignment('center');

  hoja.setFrozenRows(1);
}

// ====================================
// MENSAJE DE CARGA DEL SISTEMA
// ====================================
console.log('===============================================');
console.log('✅ Sistema de Seguimiento de Prácticas v2.8');
console.log('===============================================');
console.log(`📞 Configurado para ${LLAMADAS_PARA_FINALIZAR} llamadas por participante`);
console.log('🔒 SISTEMA ANTI-DOBLE PROCESAMIENTO: ACTIVO');
console.log('📝 COLUMNAS DE ETAPAS: ACTIVAS');
console.log('');
console.log('🆕 NUEVAS FORMACIONES:');
console.log('   • Análisis de datos E-commerce');
console.log('   • SAC (Servicio al Cliente)');
console.log('   • Ofimática');
console.log('');
console.log('📝 COLUMNAS DE ETAPAS:');
console.log('   • Aliados');
console.log('   • Plataformas');
console.log('   • Conexión laboral');
console.log('   • Por su cuenta');
console.log('   • No busca trabajar');
console.log('   • Empleado');
console.log('   • No terminó la formación');
console.log('');
console.log('🔧 MEJORAS v2.8:');
console.log('   • Escritura automática en columnas de etapas');
console.log('   • Actualización de estructura sin pérdida de datos');
console.log('   • Bloqueo por fila para evitar duplicados');
console.log('   • Diagnóstico de doble procesamiento');
console.log('   • Registro detallado en consola');
console.log('');
console.log('🚀 Sistema listo para uso completo');
console.log('💡 Espera 1-2 segundos entre checkboxes');
console.log('===============================================');
