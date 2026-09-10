function DB_TASK_getAll() {

  const cache = CacheService.getScriptCache();
  const key = "TASK_ALL_V1";
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
    .getSheetByName("TAREAS");

  const data =
    sh.getDataRange().getValues();

  const headers = data[0];

  const result = [];

  for(let i = 1; i < data.length; i++){

    const obj =
      mapRow(
        headers,
        data[i]
      );

    // =========================
    // CONVERTIR FECHAS
    // =========================

    [
      "FECHA",
      "FECHA_CORRECCION",
      "FECHA_CIERRE",
      "FECHA_ASIGNACION"
    ].forEach(campo => {

      if(
        obj[campo] instanceof Date
      ){

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

function _TASK_CACHE_clear(){

  CacheService
    .getScriptCache()
    .remove("TASK_ALL_V1");
}

function pruebita333(){
 Logger.log(DB_TASK_getAll())
}

function DB_TASK_insert(
  task
){

  const sh =
    SpreadsheetApp
      .getActive()
      .getSheetByName(
        "TAREAS"
      );

  const headers =
    sh
      .getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(h => String(h || "").trim());

  const taskByHeader = {};

  Object.keys(task || {}).forEach(key => {
    taskByHeader[
      String(key)
        .trim()
        .toUpperCase()
    ] = task[key];
  });

  const row = headers.map(header => {
    const normalized =
      String(header)
        .trim()
        .toUpperCase();

    return Object.prototype.hasOwnProperty.call(taskByHeader, normalized)
      ? taskByHeader[normalized]
      : "";
  });

  sh.appendRow(row);

  _TASK_CACHE_clear();
}



function dbGetTareaById(id){

  const sh =
    SpreadsheetApp
      .openById(CONFIG.ID_SHEET)
      .getSheetByName("TAREAS");

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

    const value =
      row[j];

    obj[h] =

      value instanceof Date

      ?

      value.toISOString()

      :

      value;

  });

  return obj;
}


function DB_TASK_update(
  payload
){

  const sh =
    SpreadsheetApp
      .getActive()
      .getSheetByName(
        "TAREAS"
      );

  const headers =
    sh.getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0];

  const col = {};

  headers.forEach((h,i) => {

    col[
      String(h)
        .trim()
        .toUpperCase()
    ] = i + 1;

  });

  const rowIndex = findRowIndexById(
    sh,
    col["ID_OBSERVACION"],
    payload.ID_OBSERVACION
  );

  if(rowIndex === -1){
    throw new Error(
      "No se encontró la tarea: " +
      payload.ID_OBSERVACION
    );
  }

  const row =
    sh.getRange(rowIndex, 1, 1, sh.getLastColumn())
      .getValues()[0];

  sh.getRange(
    rowIndex,
    col["MODULO"]
  ).setValue(
    payload.MODULO
  );

  sh.getRange(
    rowIndex,
    col["DESCRIPCION_CORTA"]
  ).setValue(
    payload.DESCRIPCION_CORTA
  );

  sh.getRange(
    rowIndex,
    col["PASOS_REPRODUCIR"]
  ).setValue(
    payload.PASOS_REPRODUCIR
  );

  sh.getRange(
    rowIndex,
    col["PRIORIDAD"]
  ).setValue(
    payload.PRIORIDAD
  );

  sh.getRange(
    rowIndex,
    col["USUARIOS_ASIGNADOS"]
  ).setValue(
    payload.USUARIOS_ASIGNADOS
  );

  // Solo permite registrar captura en edicion si aun no existe una.
  if(col["CAPTURAS"]){

    const capturaActual =
      String(
        row[
          col["CAPTURAS"] - 1
        ] || ""
      )
      .trim();

    const imagenNueva =
      String(payload.Imagen || "")
        .trim();

    if(!capturaActual && imagenNueva){

      const imageUrl =
        uploadImage(
          imagenNueva,
          payload.ID_OBSERVACION
        );

      sh.getRange(
        rowIndex,
        col["CAPTURAS"]
      ).setValue(
        imageUrl
      );
    }
  }


  // =========================
  // CAMBIAR A EN_PROCESO
  // SI EXISTE ALGUN ASIGNADO
  // =========================

  let tieneAsignados = false;

  try{

    const asignados =

      JSON.parse(
        payload.USUARIOS_ASIGNADOS || "{}"
      );

    tieneAsignados =

      Object.values(asignados)

        .some(lista =>

          Array.isArray(lista)

          &&

          lista.length > 0

        );

  }catch(error){

    Logger.log(error);

  }

  if(tieneAsignados){

    sh.getRange(
      rowIndex,
      col["ESTADO"]
    ).setValue(
      "EN_PROCESO"
    );

  }

  // =========================
  // FECHA ASIGNACION
  // SOLO LA PRIMERA VEZ
  // =========================

  const fechaAsignacionActual =

    row[
      col["FECHA_ASIGNACION"] - 1
    ];

  if(

    !fechaAsignacionActual ||

    !(fechaAsignacionActual instanceof Date)

  ){

    sh.getRange(
      rowIndex,
      col["FECHA_ASIGNACION"]
    ).setValue(
      new Date()
    );

  }

  _TASK_CACHE_clear();

  return {
    ok:true
  };

}

function DB_TASK_updateEstado(id, newEstado){

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName("TAREAS");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = {};

  headers.forEach((h, i) => {
    col[String(h).trim().toUpperCase()] = i + 1;
  });

  const rowIndex = findRowIndexById(sh, col["ID_OBSERVACION"], id);

  if(rowIndex === -1){
    throw new Error("No se encontró la tarea: " + id);
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

  _TASK_CACHE_clear();

  return {ok: true};
}