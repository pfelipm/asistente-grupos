/**
 * Crea un activador (trigger) para ejecutar la sincronización
 * periódica de todas las hojas de administración de grupos.
 * 
 * @param   {Object}  ajustes    Objeto: { cadaPeriodo, periodo ('hora' | 'dia' | 'semana' ), hora (núm 0-23) }
 * @returns {Object}  Contiene propiedad para indicar icono, color y mensaje de estado a mostrar usando M.toast() en el cuadro de diálogo.
 */
function crearActivador2(ajustes) {

  if (!verificarAdmin(false)) return { estilo: PARAMS.mToast.excepcion, mensaje: 'Debes ser administrador.' };

  // Eliminar activadores previos
  eliminarActivador(false);

  try {
    const diasMapping = { 
      'Lunes': ScriptApp.WeekDay.MONDAY, 
      'Martes': ScriptApp.WeekDay.TUESDAY, 
      'Miércoles': ScriptApp.WeekDay.WEDNESDAY, 
      'Jueves': ScriptApp.WeekDay.THURSDAY, 
      'Viernes': ScriptApp.WeekDay.FRIDAY, 
      'Sábado': ScriptApp.WeekDay.SATURDAY, 
      'Domingo': ScriptApp.WeekDay.SUNDAY 
    };

    switch (ajustes.periodo) {
      case 'hora':
        ScriptApp.newTrigger('sincronizarGrupos').timeBased().everyHours(Number(ajustes.cadaPeriodo)).create();
        break;
      case 'dia':
        ScriptApp.newTrigger('sincronizarGrupos').timeBased().everyDays(Number(ajustes.cadaPeriodo)).atHour(Number(ajustes.hora)).create();
        break;
      case 'semana':
        // ajustes.cadaPeriodo es ahora un array en el caso de semana
        const diasSeleccionados = Array.isArray(ajustes.cadaPeriodo) ? ajustes.cadaPeriodo : [ajustes.cadaPeriodo];
        if (diasSeleccionados.length === 0) throw new Error('Debes seleccionar al menos un día.');
        
        diasSeleccionados.forEach(dia => {
          ScriptApp.newTrigger('sincronizarGrupos')
            .timeBased()
            .onWeekDay(diasMapping[dia])
            .atHour(Number(ajustes.hora))
            .create();
        });
        break;
      default:
        throw new Error('Periodo no válido');
    }

    PropertiesService.getScriptProperties().setProperty(PARAMS.clavePropActivador, JSON.stringify(ajustes));
    onOpen(); // Regenerar menú

    let descPeriodo = '';
    if (ajustes.periodo === 'semana') {
      const dias = Array.isArray(ajustes.cadaPeriodo) ? ajustes.cadaPeriodo : [ajustes.cadaPeriodo];
      descPeriodo = dias.join(', ').toLowerCase();
    } else {
      descPeriodo = ajustes.cadaPeriodo == 1 
        ? (ajustes.periodo == 'hora' ? 'hora' : 'día') 
        : `${ajustes.cadaPeriodo} ${ajustes.periodo == 'hora' ? 'horas' : 'días'}`;
    }

    const mensaje = `Los grupos seleccionados se sincronizarán cada <b>${descPeriodo}</b>` +
        (ajustes.periodo == 'hora'
          ? '.'
          : ` entre las <b>${String(ajustes.hora).padStart(2, '0')}:00h</b> y las <b>${String((Number(ajustes.hora) + 1) % 24).padStart(2, '0')}:00h.</b>`);

    return { estilo: PARAMS.mToast.ok, mensaje };
    
  } catch (e) {
    console.error(e);
    return { estilo: PARAMS.mToast.excepcion, mensaje: `Error al crear el activador: ${e.message}` };
  }

}

function mostrarDialogoActivador() {

  const scriptProperties = PropertiesService.getScriptProperties();
  const ajustesActuales = scriptProperties.getProperty(PARAMS.clavePropActivador);
  
  // Título dinámico según si existe programación previa
  const tituloModal = ajustesActuales ? 'Modificar programación automática' : 'Activar sincronización automática';

  const htmlOutput = HtmlService.createTemplateFromFile('DialogoActivador').evaluate().setHeight(350);
  SpreadsheetApp.getUi().showModalDialog(htmlOutput, `${PARAMS.icono} ${tituloModal}`);

}

function mostrarDialogoDetener() {
  const template = HtmlService.createTemplateFromFile('DialogoConfirmacion');
  template.mensaje = '¿Seguro que deseas detener la sincronización automática de grupos?';
  template.funcionExito = 'confirmarEliminarActivador';
  template.txtNo = 'MANTENER';
  template.txtSi = 'DETENER';
  template.colorNo = 'green darken-3';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setHeight(230).setWidth(500), `${PARAMS.icono} Detener sincronización`);
}

function mostrarDialogoReparar() {
  const template = HtmlService.createTemplateFromFile('DialogoConfirmacion');
  template.mensaje = 'Se eliminarán todos los activadores instalados por este script y la sincronización se detendrá por completo. ¿Deseas continuar?';
  template.funcionExito = 'confirmarRepararActivador';
  template.txtNo = 'CANCELAR';
  template.txtSi = 'REPARAR';
  template.colorNo = 'grey lighten-1';
  SpreadsheetApp.getUi().showModalDialog(template.evaluate().setHeight(250).setWidth(420), `${PARAMS.icono} Reparar sistema`);
}

function confirmarEliminarActivador() {
  return eliminarActivador(false); // No usar toast nativo
}

function confirmarRepararActivador() {
  eliminarActivador(false);
  return { estilo: PARAMS.mToast.ok, mensaje: 'Sincronización desactivada y sistema reparado.' };
}

/**
 * Elimina todos los activadores de ejecución periódica existentes
 * 
 * @param {boolean} [usarToastNativo] Indica si se debe mostrar el toast nativo de la hdc.
 * @returns {Object}                  Objeto con estado para feedback visual en el cliente.
 */
function eliminarActivador(usarToastNativo = true) {

  if (verificarAdmin(usarToastNativo)) {

    // Eliminar todos los activadores instalados por este script
    const activadores = ScriptApp.getProjectTriggers();
    activadores.forEach(trigger => ScriptApp.deleteTrigger(trigger));

    // Actualizar propiedad, regenerar menú e informar al usuario
    PropertiesService.getScriptProperties().deleteProperty(PARAMS.clavePropActivador);
    onOpen();
    
    const resultado = { estilo: PARAMS.mToast.atencion, mensaje: 'Se ha desactivado la sincronización de grupos.' };

    if (usarToastNativo) {
      SpreadsheetApp.getActive().toast(resultado.mensaje, PARAMS.tituloMensajes);
    }

    return resultado;

  }

}

/**
 * Repara el sistema de sincronización automática.
 * Ahora redirigido a mostrarDialogoReparar desde el menú.
 */
function repararActivador() {
  mostrarDialogoReparar();
}

