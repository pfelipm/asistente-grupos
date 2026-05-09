/**
 * Sincroniza todas las hojas de administración de grupos, esta función puede ser invocada por:
 *   - La acción del menú.
 *   - El activador por tiempo que instala este script.
 * 
 * @param {Object} objetoEvento  Objeto que caracteriza el evento que se ha producido cuando la función se ejectua en el contexto de un activador.
 */
function sincronizarGrupos(objetoEvento) {

  // Solo se emitirán alertas y mensajes cuando se ejecuta esta función desde la acción del menú:
  // https://www.linkedin.com/posts/pfelipm_context-aware-code-standard-vs-trigger-execution-activity-7257081666941112322-iipq
  const muestraMensajes = typeof objetoEvento?.triggerUid == 'undefined';

  if (verificarAdmin(muestraMensajes)) {

    const hdc = SpreadsheetApp.getActive();

    // Obtener las hojas que parecen contener una lista de miembros de un grupo
    const hojasGrupos = hdc.getSheets().filter(hojaGrupo => {

      // ["Grupo de Google", emailGrupo]
      const marcadorEmailGrupo = hojaGrupo.getRange(PARAMS.tablaMiembros.rangoEmailGrupo).getValues();

      // Casilla de verificación que determina si la hoja se debe procesar en una operación de sincronización completa,
      // manual o automática, tanto da.
      const checkActiva = hojaGrupo.getRange(PARAMS.tablaMiembros.rangoChkActiva).getValue();
      return marcadorEmailGrupo[0][0] == PARAMS.tablaMiembros.txtMarcadorHoja && marcadorEmailGrupo[0][1] && checkActiva;

    });

    // Sincronizar hojas
    let totalAgregados = 0;
    let totalEliminados = 0;
    let totalErrores = 0;

    hojasGrupos.forEach(hojaGrupo => {
      if (muestraMensajes) hdc.toast(`Sincronizando usuarios de hoja «${hojaGrupo.getName()}»...`, PARAMS.tituloMensajes, -1);
      const resultado = sincronizarUsuariosGrupo(hojaGrupo);
      if (resultado) {
        totalAgregados += resultado.agregados;
        totalEliminados += resultado.eliminados;
        totalErrores += resultado.errores;
      }
    });

    if (muestraMensajes) {
      const resumen = `Proceso completado.\nGrupos: ${hojasGrupos.length}\nAgregados: ${totalAgregados}\nEliminados: ${totalEliminados}\nErrores: ${totalErrores}`;
      hdc.toast(resumen, PARAMS.tituloMensajes);
      SpreadsheetApp.getUi().alert(PARAMS.tituloMensajes, resumen, SpreadsheetApp.getUi().ButtonSet.OK);
    }

  }

}

/**
 * Sincroniza la lista de miembros directos del grupo representado
 * por la hoja de administración que se pasa como parámetro, o de
 * la hoja activa si se invoca sin indicarlo.
 * 
 * @param {GoogleAppsScript.Spreadsheet.Sheet}  [hojaActual]  Hoja de administración de grupo.
 * @returns {Object|void}                                    Objeto con resumen de la operación.
 */
function sincronizarUsuariosGrupo(hojaActual) {

  const hdc = SpreadsheetApp.getActive();
  let muestraMensajes = false;
  if (!hojaActual) {
    hojaActual = hdc.getActiveSheet();
    muestraMensajes = true;
  }

  if (verificarAdmin(muestraMensajes)) {

    const marcadorEmailGrupo = hojaActual.getRange(PARAMS.tablaMiembros.rangoEmailGrupo).getValues();
    const emailGrupo = marcadorEmailGrupo[0][1];
    
    if (marcadorEmailGrupo[0][0] != PARAMS.tablaMiembros.txtMarcadorHoja || !emailGrupo) {
      if (muestraMensajes) hdc.toast(`Esta hoja no parece contener una lista de usuarios de grupo o falta el email de grupo.`);
      return;
    }

    if (muestraMensajes) hdc.toast(`Sincronizando usuarios de hoja «${hojaActual.getName()}»...`, PARAMS.tituloMensajes, -1);

    const tablaDatos = hojaActual.getDataRange().getValues();
    let miembrosGrupo = [];
    let nextPageToken;
    let numErrores = 0;
    let miembrosAgregados = 0;
    let miembrosEliminados = 0;
    const emailsAgregados = [];
    const emailsEliminados = [];
    let mensajeMasInformacion = '';

    if (!existeGrupo(emailGrupo)) {
      mensajeMasInformacion = 'El grupo no existe o no tienes permisos.';
      numErrores = 1;
    } else {

      try {
        do {
          const respuesta = AdminDirectory.Members.list(emailGrupo, {
            includeDerivedMembership: false,
            pageToken: nextPageToken,
          });
          if (respuesta.members) miembrosGrupo = miembrosGrupo.concat(respuesta.members.map(miembro => miembro.email.toLowerCase()));
          nextPageToken = respuesta.nextPageToken;
        } while (nextPageToken);
      } catch (e) {
        mensajeMasInformacion = `Error al listar miembros: ${e.message}`;
        numErrores = 1;
      }

      if (numErrores === 0) {
        const tablaMiembros = tablaDatos.slice(PARAMS.tablaMiembros.filEncabezado);

        // [2] Añade nuevos miembros al grupo
        tablaMiembros.forEach((miembro, fila) => {
          const emailMiembro = String(miembro[PARAMS.tablaMiembros.colEmail - 1]).trim().toLowerCase();
          
          if (emailMiembro) {
            if (!miembrosGrupo.includes(emailMiembro)) {
              try {
                AdminDirectory.Members.insert(
                  { kind: 'admin#directory#member', email: emailMiembro, role: 'MEMBER', type: 'USER' },
                  emailGrupo
                );
                miembrosAgregados++;
                emailsAgregados.push(emailMiembro);
                tablaMiembros[fila][PARAMS.tablaMiembros.colResultado - 1] = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm');
              } catch (e) {
                tablaMiembros[fila][PARAMS.tablaMiembros.colResultado - 1] = `¡Error: ${e.message}!`;
                numErrores++;
              }
            } else {
              if (!(tablaMiembros[fila][PARAMS.tablaMiembros.colResultado - 1] instanceof Date)) {
                tablaMiembros[fila][PARAMS.tablaMiembros.colResultado - 1] = 'Ya es miembro';
              }
            }
          }
        });

        // [3] Elimina miembros ya no incluidos en el grupo
        miembrosGrupo.forEach((miembroGrupo) => {
          const existeEnTabla = tablaMiembros.some(miembroTabla => 
            String(miembroTabla[PARAMS.tablaMiembros.colEmail - 1]).trim().toLowerCase() === miembroGrupo
          );

          if (!existeEnTabla) {
            try {
              AdminDirectory.Members.remove(emailGrupo, miembroGrupo);
              miembrosEliminados++;
              emailsEliminados.push(miembroGrupo);
            } catch (e) {
              console.error(`Error al eliminar ${miembroGrupo}: ${e.message}`);
              numErrores++;
            }
          }
        });

        // [4] Escribe columna de registro del resultado en la hoja de cálculo
        tablaMiembros.forEach((miembro, fila) => {
          if (!String(miembro[PARAMS.tablaMiembros.colEmail - 1]).trim()) tablaMiembros[fila][PARAMS.tablaMiembros.colResultado - 1] = '';
        });

        hojaActual.getRange(
          PARAMS.tablaMiembros.filEncabezado + 1,
          PARAMS.tablaMiembros.colResultado,
          tablaMiembros.length,
          1,
        ).setValues(tablaMiembros.map(fila => [fila[PARAMS.tablaMiembros.colResultado - 1]]));
      }
    }

    // [5] Escribe registro de la operación global
    const hojaRegistro = hdc.getSheetByName(PARAMS.tablaRegistro.hoja);
    let registrarResumen = false;
    let registrarNotas = false;
    
    if (hojaRegistro) {
      const valoresChecks = hojaRegistro.getRange('C1:E1').getValues()[0];
      registrarResumen = valoresChecks[0]; // C1
      registrarNotas = valoresChecks[2];   // E1
    }

    if (hojaRegistro && registrarResumen) {
      const registro = [
        Utilities.formatDate(new Date(), Session.getScriptTimeZone(), 'dd/MM/yyyy HH:mm'),
        emailGrupo,
        miembrosAgregados,
        miembrosEliminados,
        numErrores,
        mensajeMasInformacion
      ];
      if (hojaRegistro.getLastRow() >= PARAMS.tablaRegistro.filDatos) hojaRegistro.insertRowBefore(PARAMS.tablaRegistro.filDatos);
      const rangoFila = hojaRegistro.getRange(PARAMS.tablaRegistro.filDatos, 1, 1, registro.length);
      rangoFila.setValues([registro]);

      // Añadir notas con el detalle de emails solo si el segundo check (D1) está activo
      if (registrarNotas) {
        if (emailsAgregados.length > 0) {
          hojaRegistro.getRange(PARAMS.tablaRegistro.filDatos, 3).setNote('Emails añadidos:\n' + emailsAgregados.join(',\n'));
        }
        if (emailsEliminados.length > 0) {
          hojaRegistro.getRange(PARAMS.tablaRegistro.filDatos, 4).setNote('Emails eliminados:\n' + emailsEliminados.join(',\n'));
        }
      }
    }

    if (muestraMensajes) hdc.toast(`Proceso completado: +${miembrosAgregados}, -${miembrosEliminados}, ⚠️${numErrores}`, PARAMS.tituloMensajes);
    
    return { agregados: miembrosAgregados, eliminados: miembrosEliminados, errores: numErrores };

  }

}

/**
 * Determina si el Grupo que se pasa como parámetro
 * existe en el dominio Google Workspace.
 * 
 * @param   {string}  emailGrupo  Email del grupo a comprobar
 * @returns {boolean}       
 */
function existeGrupo(emailGrupo) { try { AdminDirectory.Groups.get(emailGrupo); return true; } catch (e) { return false; } }
