function DB_OBS_insert(obs) {

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName("OBSERVACIONES");

  sh.appendRow([
    obs.ID_OBSERVACION,
    obs.Fecha,
    obs.QA,
    obs.DEV_asignado,
    obs.Validación_relacionada,
    obs.Módulo,
    obs.Descripción_corta,
    obs.Pasos_reproducir,
    obs.Capturas,
    obs.Estado,
    obs.Prioridad,
    obs.DEV_comentario,
    obs.Fecha_corrección,
    obs.Fecha_cierre
  ]);

  _OBS_CACHE_clear();
}

function _OBS_CACHE_clear(){

  CacheService
    .getScriptCache()
    .remove("OBS_ALL_V1");
}

function DB_OBS_getAll() {

  const cache = CacheService.getScriptCache();
  const key = "OBS_ALL_V1";
  const cached = cache.get(key);

  if(cached){
    try{
      return JSON.parse(cached);
    }catch(error){
      Logger.log(error);
    }
  }

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName("OBSERVACIONES");

  const data =
    sh.getDataRange().getValues();

  const headers = data[0];

  const result = [];

  for(let i = 1; i < data.length; i++){

    const obj =
      mapRow(headers, data[i]);

    // =========================
    // CONVERTIR FECHAS
    // =========================

    [
      "FECHA",
      "FECHA_CORRECCION",
      "FECHA_CIERRE",
      "FECHA_ASIGNACION"
    ].forEach(campo => {

      if(obj[campo] instanceof Date){

        obj[campo] =
          Utilities.formatDate(
            obj[campo],
            Session.getScriptTimeZone(),
            "dd/MM/yyyy HH:mm:ss"
          );
      }
    });

    result.push(obj);
  }

  try{
    cache.put(
      key,
      JSON.stringify(result),
      20
    );
  }catch(error){
    Logger.log(error);
  }

  return result;
}


function dbGetObservacionById(id){

  const sh =
    SpreadsheetApp
      .openById(CONFIG.ID_SHEET)
      .getSheetByName("OBSERVACIONES");

  const headers =
    sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const rowIndex = findRowIndexById(sh, 1, id);

  if(rowIndex === -1){
    return null;
  }

  const row =
    sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];

  let obj = {};

  headers.forEach((h,j)=>{

    const value = row[j];

    obj[h] =

      value instanceof Date

      ?

      value.toISOString()

      :

      value;

  });

  return obj;
}

function actualizarObservacion(payload, userFromClient){

  console.log(payload);

    const user =
      getSession(userFromClient || payload.user || payload.__user);

    if(!user){

      throw new Error(
        "No autenticado"
      );
    }

    return OBS_SERVICE_update(
      payload,
      user
    );
}


function DB_OBS_updateEstado(id, newEstado){

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName("OBSERVACIONES");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = {};

  headers.forEach((h, i) => {
    col[String(h).trim().toUpperCase()] = i + 1;
  });

  const rowIndex = findRowIndexById(sh, col["ID_OBSERVACION"], id);

  if(rowIndex === -1){
    throw new Error("No se encontró la observación: " + id);
  }

  sh.getRange(rowIndex, col["ESTADO"]).setValue(newEstado);

  if(newEstado === "CERRADO" && col["FECHA_CIERRE"]){
    sh.getRange(rowIndex, col["FECHA_CIERRE"]).setValue(new Date());
  }

  if(newEstado === "EN_PROCESO" && col["FECHA_ASIGNACION"]){
    const currentAssign = sh.getRange(rowIndex, col["FECHA_ASIGNACION"]).getValue();
    if(!currentAssign){
      sh.getRange(rowIndex, col["FECHA_ASIGNACION"]).setValue(new Date());
    }
  }

  _OBS_CACHE_clear();

  return {ok: true};
}

function prueba(id){
  id="OBS_1779264162758"
  Logger.log(dbGetObservacionById(id))
}
