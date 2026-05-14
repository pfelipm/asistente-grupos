
/**
 * Muestra el diálogo para elegir qué usuarios descargar del directorio.
 */
function mostrarDialogoSeleccionDirectorio() {
  if (verificarAdmin()) {
    const html = HtmlService.createHtmlOutputFromFile('DialogoSeleccionDirectorio')
      .setWidth(450)
      .setHeight(400)
      .setTitle(`${PARAMS.icono} Configurar descarga`);
    SpreadsheetApp.getUi().showModalDialog(html, `${PARAMS.icono} Descarga de directorio`);
  }
}

/**
 * Obtiene los dominios y las UOs de la consola para el diálogo.
 */
function obtenerOpcionesParaDirectorio() {
  try {
    const customer = 'my_customer';
    
    // Obtener Dominios
    const resDominios = AdminDirectory.Domains.list(customer);
    const dominios = resDominios.domains ? resDominios.domains.map(d => d.domainName).sort() : [];
    
    // Obtener UOs
    const resUos = AdminDirectory.Orgunits.list(customer, { type: 'all' });
    const uos = resUos.organizationUnits ? resUos.organizationUnits.map(uo => uo.orgUnitPath).sort() : [];
    if (!uos.includes('/')) uos.unshift('/');

    return { dominios, uos };
  } catch (e) {
    throw new Error(`Error al obtener estructura de la consola: ${e.message}`);
  }
}

/**
 * Ejecuta la descarga de usuarios según la configuración elegida en el diálogo.
 * @param {Object} config { tipo: 'total'|'dominio'|'uo', valor: string }
 */
function ejecutarDescargaDirectorio(config) {
  try {
    return descargarDirectorioUsuarios(config);
  } catch (e) {
    throw new Error(e.message);
  }
}

/**
 * Descarga y escribe en la tabla de usuarios internos de la hdc.
 * @param {Object} [config] Opciones de filtrado.
 */
function descargarDirectorioUsuarios(config = { tipo: 'total' }) {

  if (verificarAdmin()) {

    let usuarios = [];
    let nextPageToken;
    const hdc = SpreadsheetApp.getActive();

    // Configuración de la petición según el filtro
    const opcionesPeticion = {
      customer: 'my_customer',
      viewType: 'admin_view',
      maxResults: 500
    };

    let mensajeLog = 'Descargando usuarios de toda la organización...';
    if (config.tipo === 'dominio') {
      opcionesPeticion.domain = config.valor;
      mensajeLog = `Descargando usuarios del dominio ${config.valor}...`;
    } else if (config.tipo === 'uo') {
      opcionesPeticion.orgUnitPath = config.valor;
      mensajeLog = `Descargando usuarios de la UO ${config.valor}...`;
    }

    console.info(mensajeLog);

    try {
      do {
        opcionesPeticion.pageToken = nextPageToken;
        const respuesta = AdminDirectory.Users.list(opcionesPeticion);
        
        if (respuesta.users) {
          usuarios = usuarios.concat(
            respuesta.users.map(u => [u.name.fullName, u.primaryEmail, u.suspended])
          );
        }
        nextPageToken = respuesta.nextPageToken;

      } while (nextPageToken);
    } catch (e) {
      console.error(`Error en la descarga: ${e.message}`);
      throw new Error(`Fallo al conectar con Admin SDK: ${e.message}`);
    }

    if (usuarios.length === 0) {
      return 'No se han encontrado usuarios con el filtro seleccionado.';
    }

    const rangoUsuarios = hdc.getRange(PARAMS.tablaDirectorioUsuarios);
    rangoUsuarios.clearContent();
    
    // Escribir datos
    rangoUsuarios.offset(0, 0, usuarios.length, usuarios[0].length)
      .setValues(
        usuarios.sort(((u1, u2) => u1[0].localeCompare(u2[0])))
      );

    // Asegurar que los datos se asientan antes de actualizar intervalos
    SpreadsheetApp.flush();
    actualizarIntervalosConNombre(hdc);
    
    const msgFinal = `Proceso completado. ${usuarios.length} usuarios descargados.`;
    return msgFinal;

  }
}


/**
 * Descarga y escribe en la tabla de grupos internos de la hdc
 * la lista de grupos del dominio.
 */
function descargarDirectorioGrupos() {

  if (verificarAdmin()) {

    let grupos = [];
    let nextPageToken;
    const hdc = SpreadsheetApp.getActive();
    
    // Para grupos, el Admin SDK requiere domain o customer. 
    // Usamos customer para obtener TODOS los grupos de la consola (multi-dominio).
    const opcionesPeticion = {
      customer: 'my_customer',
      maxResults: 500
    };

    hdc.toast('Descargando todos los grupos de la organización...', PARAMS.tituloMensajes, -1);

    try {
      do {
        opcionesPeticion.pageToken = nextPageToken;
        const respuesta = AdminDirectory.Groups.list(opcionesPeticion);
        
        if (respuesta.groups) {
          grupos = grupos.concat(
            respuesta.groups.map(g => [g.name, g.email])
          );
        }
        nextPageToken = respuesta.nextPageToken;

      } while (nextPageToken);
    } catch (e) {
      hdc.toast(`Error en la descarga de grupos: ${e.message}`, PARAMS.tituloMensajes);
      throw new Error(`Fallo al descargar grupos: ${e.message}`);
    }

    if (grupos.length === 0) {
      hdc.toast('No se han encontrado grupos en la organización.', PARAMS.tituloMensajes);
      return;
    }

    const rangoGrupos = hdc.getRange(PARAMS.tablaDirectorioGrupos);
    rangoGrupos.clearContent();
    rangoGrupos.offset(0, 0, grupos.length, grupos[0].length)
      .setValues(
        grupos.sort(((u1, u2) => u1[0].localeCompare(u2[0])))
      );

    SpreadsheetApp.flush();
    actualizarIntervalosConNombre(hdc);
    return `Proceso completado. ${grupos.length} grupos descargados.`;

  }

}