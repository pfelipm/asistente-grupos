/**
 * Un script para sincronizar listas de usuarios en una hoja de cálculo
 * con los miembros del grupos del servicio de Grupos de Google
 * 
 * NOTA:  Parece haber un retraso de 10 - 15 segundos en que la API obtenga
 *        la lista actualizada de miembros de un grupo tras una operación
 *        de adición o eliminación de miembros.
 *         
 * V1.2 Noviembre 2024 (corrección de errores, nuevas verificaciones de usuario activo, mejoras en los mensajes emitidos)
 * 
 * @OnlyCurrentDoc
 */

const PARAMS = {
  version: 'Version: 2.0 | Mayo 2026',
  nombreApp: 'Asistente de Grupos',
  icono: '🐙',
  tituloMensajes: '🐙 Asistente de Grupos',
  hojaDirectorios: 'Directorio',
  tablaDirectorioUsuarios: 'Directorio!A2:C',
  tablaDirectorioGrupos: 'Directorio!H2:I',
  tablaMiembros: {
    rangoChkActiva: 'A4',
    rangoEmailGrupo: 'B1:C1', // Etiqueta de texto + email de grupo seleccionado
    txtMarcadorHoja: 'Grupo de Google', // Identifica la hoja como contenedora de una lista de miembros
    filEncabezado: 4,
    colEmail: 3,
    colResultado: 4
  },
  tablaRegistro: {
    hoja: 'Registro',
    filDatos: 3, // antes 2
    checkRegistro: 'C1',
    checkNotas: 'E1'
  },
  prefijoIntervaloConNombre: 'AG_',
  // Clave utilizada para almacenar en las propiedades del script los ajustes del activador { cadaPeriodo, periodo, hora}
  clavePropActivador: 'ajustesActivador', 
  // Clave utilizada para almacenar en las propiedades del usuario si es o no admin
  clavePropEsAdmin: 'esAdmin', // almacenada en UserProperties
  // Aspecto del M.toast()
  mToast: {
    ok: { icono: 'check_circle', color: 'green darken-4' },
    atencion: { icono: 'error', color: 'orange darken-4' },
    excepcion: { icono: 'warning', color: 'red darken-4' },
  }
};

/**
 * Construye el menú de acciones del script
 */
function onOpen() {

  // Verificar si se ha determinado ya que el usuario actual es o no un admin del dominio,
  // se utiliza el almacén de propiedades del usuario para que la comprobación sea por cuenta
  // que accede a la hoja de cálculo.
  if (!JSON.parse(PropertiesService.getUserProperties().getProperty(PARAMS.clavePropEsAdmin))) {

    // Mostrar solo el comando de activación que verifica si el usuario es admin
    SpreadsheetApp.getUi().createMenu(`${PARAMS.icono} ${PARAMS.nombreApp}`)
      .addItem('Activar Asistente de Grupos', 'activarFunciones')
      .addSeparator()
      .addItem('Acerca de Asistente de Grupos', 'acercaDe')
      .addToUi()

  } else {

    // Mostrar el menú completo
    const ajustesActivador = PropertiesService.getScriptProperties().getProperty(PARAMS.clavePropActivador);

    const menu = SpreadsheetApp.getUi().createMenu(`${PARAMS.icono} ${PARAMS.nombreApp}`)
      .addItem('🔄 Sincronizar hoja activa', 'sincronizarUsuariosGrupo')
      .addItem('🔃 Sincronizar las hojas marcadas', 'sincronizarGrupos')
      .addSeparator()
      .addItem('👤 Descargar usuarios del directorio', 'descargarDirectorioUsuarios')
      .addItem('👥 Descargar grupos del directorio', 'descargarDirectorioGrupos')
      .addSeparator();

    if (!ajustesActivador) {
      menu.addItem('🟢 Programar sincronización', 'mostrarDialogoActivador');
    } else {
      menu.addItem('⚙️ Consultar / modificar programación', 'mostrarDialogoActivador')
          .addItem('🟠 Detener sincronización', 'mostrarDialogoDetener');
    }

    menu.addSeparator()
      .addItem('⚠️ Reparar sistema de sincronización', 'repararActivador')
      .addSeparator()
      .addItem('ℹ️ Acerca de Asistente de Grupos', 'acercaDe')
      .addToUi();

  }

}

/**
 * Muestra información sobre el script
 */
function acercaDe() {

  const panel = HtmlService.createTemplateFromFile('Acerca de');
  panel.version = PARAMS.version;
  panel.appName = PARAMS.nombreApp;
  SpreadsheetApp.getUi().showModalDialog(panel.evaluate().setWidth(470).setHeight(420), `${PARAMS.icono} ¿Qué es ${PARAMS.nombreApp}?`);

}

/**
 * Comprueba si el usuario actual es administrador y, en su caso,
 * actualiza la propiedad interna y regenera el menú para que aparezca
 * con los comandos (funciones) habituales.
 */
function activarFunciones() {

  const ui = SpreadsheetApp.getUi();

  if (esGWSAdmin()) {
    PropertiesService.getUserProperties().setProperty(PARAMS.clavePropEsAdmin, JSON.stringify(true));
    onOpen();
    ui.alert(
      PARAMS.tituloMensajes,
      `🟢 Parece que eres un Administrador, se han activado todas las funciones de ${PARAMS.nombreApp}.`,
      ui.ButtonSet.OK
    );
  } else {
    PropertiesService.getUserProperties().setProperty(PARAMS.clavePropEsAdmin, JSON.stringify(false));
    ui.alert(
      PARAMS.tituloMensajes,
      `🔴 ${PARAMS.nombreApp} debe ser utilizado por un Administrador del dominio.`,
      ui.ButtonSet.OK
    );
  }

}

/**
 * Verifica que el usuario que lo ejecuta es admin, en caso contrario
 * muestra una alerta. Usado como comprobación previa en todos los
 * comandos del menú que con funcionalidad operativa, pero solo como
 * precaución adicional (en onOpen se comprueba siempre si el usuario
 * actual ha activado el script), el usuario prodría haber activado el
 * script siendo admin pero haber sido degradado posteriormente.
 * 
 * @param   {boolean} [usarUi]  Indica si deben emitirse alertas usando la clase Ui, verdadero por defecto.  
 * @returns {boolean}           Devuelve true si es admin, false en caso contrario.
 */
function verificarAdmin(usarUi = true) {


  const esAdmin = esGWSAdmin();
  if (!esAdmin) {

    // ¡Hemos cazado a un usuario que (ya) no es administrador, devolvemos el menú a su estado no inicializado
    PropertiesService.getUserProperties().setProperty(PARAMS.clavePropEsAdmin, JSON.stringify(false));
    onOpen();
    if (usarUi) {
      const ui = SpreadsheetApp.getUi();
      ui.alert(PARAMS.tituloMensajes, `🔴 ${PARAMS.nombreApp} debe ser utilizado por un Administrador del dominio.`, ui.ButtonSet.OKusarUi);
    }

  }
  return esAdmin;

}

/**
 * Indica si el usuario cuyo email se facilita como parámetro es
 * administrador de un dominio Google Workspace
 * 
 * Caso 1: El email es personal (gmail) → FALSO
 * Caso 2: El usuario que ejecuta script es Admin → lo que diga AdminDirectory.Users
 * Caso 3: El usuario que ejecuta script *no* es Admin y el email es del usuario que ejecuta → FALSO
 * Caso 4: Cualquier otra situación → NULL (indeterminado, no es posible comprobarlo)
 * 
 * @param   {string}        [email] Email o 'me' para indicar cuenta bajo cuya autoridad se ejecuta el script.
 * @returns {boolean|null}  Se devuelve null en caso de que no se haya podido comprobar.
 */
function esGWSAdmin(email = 'me') {

  let esAdmin;
  const testEmail = email == 'me' ? Session.getActiveUser().getEmail() : email;

  // ¿Se trata de una cuenta personal?
  if (testEmail.slice(-9) == 'gmail.com') {
    esAdmin = false;
  } else {
    try {
      // Si un usuario no Admin del dominio invoca este método y el email no existe se devuelve 404
      // Si un usuario no Admin del dominio invoca este método y el email sí existe (incluso si es el propio) se devuelve 403 (¿es esto seguro?)
      esAdmin = AdminDirectory.Users.get(testEmail).isAdmin;
    } catch (e) {
      console.info(e.details.code);
      switch (e.details.code) {
        case 404: // no existe testEmail
          esAdmin = false
          break;
        case 403: // operación prohíbida
          esAdmin = null;
          break;
        default:  // para todo lo demás...
          esAdmin = null;
      }
    }
  }
  return esAdmin;

}

/**
 * Los intervalos con nombre no se expanden cuando se añaden filas por su parte inferior,
 * y además no es posible utilizar intervalos abiertos en su definición. Asistente de Grupos
 * utiliza 4 intervalos con nombre para hacer más inteligibles sus fórmulas y alimentar 
 * los desplegables que permiten seleccionar el email del grupo a sincronizar y los nombres
 * de sus miembros:
 *   - [AG_nombreGruposDirectorio]: lista de nombres de los grupos del dominio.
 *   - [AG_emailsGruposDirectorio]: lista de direcciones de email de los grupos del domininio.
 *   - [AG_nombresCompleto]: lista de nombres unificada, incluye usuarios y grupos de dominio + externos añadidos manualmente.
 *   - [AG_emailsCompleto]: lista de emails unificada, incluye usuarios y grupos de dominio + externos añadidos manualmente.
 *  Al descargar la lista de usuarios y grupos del directorio (dominio), es posible que se añadan nuevas filas por debajo. Esta función
 *  se encarga de ajustarlos para que no se queden elementos fuera (más abajo) de ellos.
 */
function actualizarIntervalosConNombre() {

  const hdc = SpreadsheetApp.getActive();

  // Obtiene los intervalos con nombre que utiliza Asistente de Grupos
  const intervalosAG = hdc.getNamedRanges().filter(intervalo => intervalo.getName().startsWith(PARAMS.prefijoIntervaloConNombre));

  // Se utiliza como referencia segura la última fila de la hoja con datos, aunque no es exacto dado que las nuevas celdas podrían
  // no estar dentro del intervalo con nombre
  const ultimaFila = hdc.getSheetByName(PARAMS.hojaDirectorios).getLastRow();
  intervalosAG.forEach(intervalo => {
    const rangoActual = intervalo.getRange();
    const numFilas = Math.max(1, ultimaFila - rangoActual.getRow() + 1);
    intervalo.setRange(rangoActual.offset(0, 0, numFilas));
  });

}