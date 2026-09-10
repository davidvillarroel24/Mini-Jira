function OBS_SERVICE_create(payload, user) {

  const id =
    generateId("OBS");

  let imageUrl = "";

  if(payload.Imagen){

    imageUrl = uploadImage(
      payload.Imagen,
      id
    );
  }

  const observacion = {

    ID_OBSERVACION: id,

    Fecha: new Date(),

    QA: payload.qa, //user.nombre,

    DEV_asignado: "",

    Validación_relacionada: "",

    Módulo: payload.Modulo,

    Descripción_corta:
      payload.Descripcion,

    Pasos_reproducir:
      payload.Pasos,

    Capturas: imageUrl,

    Estado: "ABIERTO",

    Prioridad:
      payload.Prioridad,

    DEV_comentario: "",

    Fecha_corrección: "",

    Fecha_cierre: ""
  };

  DB_OBS_insert(observacion);

  _OBS_CACHE_clear();

  return {
    ok:true,
    id:id
  };
}


function isDeleted(obs){

  return String(obs.ELIMINADO || "")
    .trim()
    .toUpperCase() === "TRUE";
}

function OBS_SERVICE_list(
  user,
  filters = {}
) {

  let data =
    DB_OBS_getAll();

  // =========================
  // FILTRO DE VISTA
  // =========================

  switch(filters.view){

    case "ELIMINADAS":

      data =
        data.filter(
          obs => isDeleted(obs)
        );

      break;

    default:

      data =
        data.filter(
          obs => !isDeleted(obs)
        );
  }

  // =========================
  // PM Y LIDERES VEN TODO
  // =========================

  if(PERM_isManager(user)){

    return data;
  }

  // =========================
  // QA VE SOLO SUS OBS
  // =========================

  if(PERM_isRole(user, "QA")){

    return data.filter(x =>

      String(x.QA)
        .trim()
        .toUpperCase()

      ===

      String(user.nombre)
        .trim()
        .toUpperCase()
    );
  }

  // =========================
  // DEV VE SOLO ASIGNADAS
  // =========================

  if(PERM_isRole(user, "DEV")){

    return data.filter(x =>

      String(x.DEV_ASIGNADO)
        .trim()
        .toUpperCase()

      ===

      String(user.nombre)
        .trim()
        .toUpperCase()
    );
  }

  // =========================
  // DEFAULT
  // =========================

  return [];
}

function OBS_SERVICE_countAssignedOpen(userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  const userName = String(user.nombre || "").trim().toUpperCase();

  if(!userName){
    return {count: 0};
  }

  const data = DB_OBS_getAll() || [];

  const count = data.filter(obs => {

    if(isDeleted(obs)){
      return false;
    }

    const estado = String(obs.ESTADO || "").trim().toUpperCase();

    if(estado === "CERRADO"){
      return false;
    }

    const devAsignado = String(obs.DEV_ASIGNADO || "").trim().toUpperCase();

    if(devAsignado === userName){
      return true;
    }

    try{
      const asignados = JSON.parse(String(obs.USUARIOS_ASIGNADOS || "{}"));
      const devList = Array.isArray(asignados && asignados.DEV) ? asignados.DEV : [];

      return devList.some(name => String(name || "").trim().toUpperCase() === userName);
    }catch(error){
      return false;
    }
  }).length;

  return {count: count};
}

function getObservacionById(id){

  return dbGetObservacionById(id);

}



function OBS_SERVICE_update(
  payload,
  user
){

  Logger.log(
    JSON.stringify(payload)
  );

  const sheet =
    SpreadsheetApp
      .openById(
        CONFIG.ID_SHEET
      )
      .getSheetByName(
        "OBSERVACIONES"
      );

  const headers =
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const idCol =
    headers.indexOf(
      "ID_OBSERVACION"
    );

  const moduloCol =
    headers.indexOf(
      "MODULO"
    );

  const descCol =
    headers.indexOf(
      "DESCRIPCION_CORTA"
    );

  const pasosCol =
    headers.indexOf(
      "PASOS_REPRODUCIR"
    );

  const prioridadCol =
    headers.indexOf(
      "PRIORIDAD"
    );

  const capturasCol =
    headers.indexOf(
      "CAPTURAS"
    );

  const rowIndex = findRowIndexById(
    sheet,
    idCol + 1,
    payload.ID_OBSERVACION
  );

  if(rowIndex !== -1){

    const row =
      sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

    // =====================
    // TEXTO
    // =====================

    row[moduloCol] =

      payload.MODULO;

    row[descCol] =

      payload
        .DESCRIPCION_CORTA;

    row[pasosCol] =

      payload
        .PASOS_REPRODUCIR;

    row[prioridadCol] =

      payload.PRIORIDAD;

    // =====================
    // IMAGEN
    // =====================

    const tieneImagen =

      row[capturasCol];

    if(

      !tieneImagen &&

      payload.Imagen

    ){

      Logger.log(
        "Subiendo imagen..."
      );

      const imageUrl =

        uploadImage(

          payload.Imagen,

          payload
            .ID_OBSERVACION

        );

      Logger.log(
        imageUrl
      );

      row[capturasCol] =
        imageUrl;
    }

    sheet
      .getRange(rowIndex, 1, 1, row.length)
      .setValues([row]);
  }

  _OBS_CACHE_clear();

  return true;
}

function OBS_SERVICE_delete(id){

  const sheet =
    SpreadsheetApp
      .openById(
        CONFIG.ID_SHEET
      )
      .getSheetByName(
        "OBSERVACIONES"
      );

  const rowIndex = findRowIndexById(sheet, 1, id);

  if(rowIndex !== -1){

    // Columna O
    sheet.getRange(rowIndex, 15)
      .setValue(true);

    _OBS_CACHE_clear();

    return {
      success:true
    };
  }

  throw new Error(
    "Observación no encontrada"
  );
}


function OBS_SERVICE_restore(id) {

  const sheet =
    SpreadsheetApp
      .openById(
        CONFIG.ID_SHEET
      )
      .getSheetByName(
        "OBSERVACIONES"
      );

  const rowIndex = findRowIndexById(sheet, 1, id);

  if(rowIndex !== -1){

    // Columna O = 15

    sheet
      .getRange(
        rowIndex,
        15
      )
      .setValue("FALSE");

    _OBS_CACHE_clear();

    return {
      success:true
    };
  }

  throw new Error(
    "Observación no encontrada"
  );
}

function USUARIOS_SERVICE_getDevs(){

  return obtenerUsuariosActivos()

    .filter(
      user => user.rol === "DEV"
    );

  
}

function pruebita(){
  Logger.log(USUARIOS_SERVICE_getDevs())
}


function OBS_SERVICE_assignDev(
  payload
){

  const sheet =
    SpreadsheetApp
      .openById(
        CONFIG.ID_SHEET
      )
      .getSheetByName(
        "OBSERVACIONES"
      );

  const headers =
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = {};

  headers.forEach((h, i) => {
    col[
      String(h)
        .trim()
        .toUpperCase()
    ] = i;
  });

  const rowIndex = findRowIndexById(sheet, 1, payload.id);

  if(rowIndex !== -1){

    const row =
      sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

    // ====================
    // PRIORIDAD
    // ====================

    row[col["PRIORIDAD"]] =
      payload.prioridad;

    // ====================
    // DEV ASIGNADO
    // ====================

    row[col["DEV_ASIGNADO"]] =
      payload.devAsignado || "";

    // ====================
    // SIN DEV
    // ====================

    if(
      !payload.devAsignado
    ){

      row[col["ESTADO"]] =
        "ABIERTO";

      row[col["FECHA_ASIGNACION"]] =
        "";

    }

    // ====================
    // CON DEV
    // ====================

    else{

      row[col["ESTADO"]] =
        "EN_PROCESO";

      if(
        !row[col["FECHA_ASIGNACION"]]
      ){

        row[col["FECHA_ASIGNACION"]] =
          new Date();
      }

    }

    // ====================
    // GUARDAR FILA
    // ====================

    sheet
      .getRange(
        rowIndex,
        1,
        1,
        row.length
      )
      .setValues([
        row
      ]);

    _OBS_CACHE_clear();

    return {
      success:true
    };
  }

  throw new Error(
    "Observación no encontrada"
  );
}

function OBS_SERVICE_addComment(
  payload
){

  const sheet =
    SpreadsheetApp
      .openById(
        CONFIG.ID_SHEET
      )
      .getSheetByName(
        "OBSERVACIONES"
      );

  const headers =
    sheet.getRange(1, 1, 1, sheet.getLastColumn()).getValues()[0];

  const col = {};

  headers.forEach((h, i) => {

    col[
      String(h)
        .trim()
        .toUpperCase()
    ] = i;

  });

  // =====================
  // VALIDACIONES
  // =====================

  if(

    payload.marcarRevision

    &&

    payload.rol !== "DEV"

  ){

    throw new Error(
      "Solo el DEV puede marcar una observación para revisión"
    );
  }

  if(

    payload.aprobarCorreccion

    &&

    ![

      "QA",
      "LIDER_QA",
      "LIDER_DEV",
      "PM"

    ].includes(
      payload.rol
    )

  ){

    throw new Error(
      "No tiene permisos para aprobar correcciones"
    );
  }

  if(

    payload.rechazarCorreccion

    &&

    ![

      "QA",
      "LIDER_QA",
      "LIDER_DEV",
      "PM"

    ].includes(
      payload.rol
    )

  ){

    throw new Error(
      "No tiene permisos para rechazar correcciones"
    );
  }

  const acciones =

    Number(
      !!payload.marcarRevision
    )

    +

    Number(
      !!payload.aprobarCorreccion
    )

    +

    Number(
      !!payload.rechazarCorreccion
    );

  if(
    acciones > 1
  ){

    throw new Error(
      "Solo puede realizar una acción por comentario"
    );
  }

  // =====================
  // BUSCAR OBSERVACION
  // =====================

  const rowIndex = findRowIndexById(sheet, col["ID_OBSERVACION"] + 1, payload.id);

  if(rowIndex === -1){
    throw new Error(
      "Observación no encontrada"
    );
  }

  const row =
    sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  let comentarios = [];

  try{

    comentarios =
      JSON.parse(
        row[
          col["COMENTARIO_JSON"]
        ] || "[]"
      );

  }
  catch(e){

    comentarios = [];
  }

  comentarios.push({

    usuario:
      payload.usuario,

    rol:
      payload.rol,

    fecha:
      Utilities.formatDate(

        new Date(),

        Session
          .getScriptTimeZone(),

        "yyyy-MM-dd HH:mm:ss"

      ),

    mensaje:
      payload.mensaje,

    accion:

      payload.marcarRevision
        ? "EN_REVISION"

      :

      payload.aprobarCorreccion
        ? "APROBADO"

      :

      payload.rechazarCorreccion
        ? "RECHAZADO"

      :

      "COMENTARIO"

  });

  // =====================
  // COMENTARIO_JSON
  // =====================

  row[
    col["COMENTARIO_JSON"]
  ] =
    JSON.stringify(
      comentarios
    );

  // =====================
  // DEV ENVIA
  // =====================

  if(
    payload.rol === "DEV"
  ){

    row[
      col["FECHA_CORRECCION"]
    ] =
      new Date();
  }

  // =====================
  // ENVIA A REVISION
  // =====================

  if(
    payload.marcarRevision
  ){

    row[
      col["ESTADO"]
    ] =
      "EN_REVISION";
  }

  // =====================
  // APRUEBA
  // =====================

  if(
    payload.aprobarCorreccion
  ){

    row[
      col["ESTADO"]
    ] =
      "CERRADO";

    row[
      col["FECHA_CIERRE"]
    ] =
      new Date();
  }

  // =====================
  // RECHAZA
  // =====================

  if(
    payload.rechazarCorreccion
  ){

    row[
      col["ESTADO"]
    ] =
      "EN_PROCESO";

    row[
      col["FECHA_CIERRE"]
    ] =
      "";
  }

  // =====================
  // GUARDAR
  // =====================

  sheet
    .getRange(
      rowIndex,
      1,
      1,
      row.length
    )
    .setValues([
      row
    ]);

  _OBS_CACHE_clear();

  return {
    success:true
  };
}