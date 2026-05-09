
/**
 * Descarga y escribe en la tabla de usuarios internos de la hdc
 * la lista de usuarios del dominio: [fullName, email, suspended].
 */
function descargarDirectorioUsuarios() {

  if (verificarAdmin()) {

    let usuarios = [];
    let nextPageToken;

    const hdc = SpreadsheetApp.getActive();

    hdc.toast('Descargando directorio de usuarios del dominio...', PARAMS.tituloMensajes, -1);
    const domain = Session.getEffectiveUser().getEmail().match(/@(.+)/)[1];

    do {

      const respuesta = AdminDirectory.Users.list({
        domain,
        customer: 'my_customer',
        pagToken: nextPageToken,
        viewType: 'admin_view',
      });
      if (respuesta.users) usuarios = usuarios.concat(
        // Extrae los campos deseados
        respuesta.users.map(usuario => [usuario.name.fullName, usuario.primaryEmail, usuario.suspended])
      );
      nextPageToken = respuesta.nextPageToken;

    } while (nextPageToken);

    // console.info(usuarios);

    const rangoUsuarios = hdc.getRange(PARAMS.tablaDirectorioUsuarios);
    rangoUsuarios.clearContent();
    rangoUsuarios.offset(0, 0, usuarios.length, usuarios[0].length)
      .setValues(
        usuarios.sort(((u1, u2) => u1[0].localeCompare(u2[0])))
      );

    // Ajusta el tamaño (nº de filas de los intervalos con nombre utilizados por Asistente de Grupos)
    actualizarIntervalosConNombre();

    hdc.toast(`Proceso completado, revisa la lista de grupos en la hoja «${rangoUsuarios.getSheet().getName()}».`, PARAMS.tituloMensajes);

  }

}


/**
 * Descarga y escribe en la tabla de grupos internos de la hdc
 * la lista de grupos del dominio: [name, email].
 */
function descargarDirectorioGrupos() {

  if (verificarAdmin()) {

    let grupos = [];
    let nextPageToken;
    const domain = Session.getEffectiveUser().getEmail().match(/@(.+)/)[1];


    const hdc = SpreadsheetApp.getActive();

    hdc.toast('Descargando directorio de GRUPOS del dominio...', PARAMS.tituloMensajes, -1);

    do {

      const respuesta = AdminDirectory.Groups.list({
        domain,
        customer: 'my_customer',
        pagToken: nextPageToken,
        sortOrder: 'ASCENDING',
      });
      if (respuesta.groups) grupos = grupos.concat(
        // Extrae los campos deseados
        respuesta.groups.map(grupo => [grupo.name, grupo.email])
      );
      nextPageToken = respuesta.nextPageToken;

    } while (nextPageToken);

    // console.info(usuarios);

    const rangoGrupos = hdc.getRange(PARAMS.tablaDirectorioGrupos);
    rangoGrupos.clearContent();
    rangoGrupos.offset(0, 0, grupos.length, grupos[0].length)
      .setValues(
        grupos.sort(((u1, u2) => u1[0].localeCompare(u2[0])))
      );

    // Ajusta el tamaño (nº de filas de los intervalos con nombre utilizados por Asistente de Grupos)
    actualizarIntervalosConNombre();

    hdc.toast(`Proceso completado, revisa la lista de grupos en la hoja «${rangoGrupos.getSheet().getName()}».`, PARAMS.tituloMensajes);

  }

}