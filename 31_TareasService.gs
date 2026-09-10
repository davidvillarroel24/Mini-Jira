function TASK_SERVICE_list(
  user,
  filters = {}
){

  let data =
    DB_TASK_getAll();

  // =========================
  // FILTRO DE VISTA
  // =========================

  switch(filters.view){

    case "ELIMINADAS":

      data =
        data.filter(
          task => isDeleted(task)
        );

      break;

    default:

      data =
        data.filter(
          task => !isDeleted(task)
        );
  }

  // =========================
  // PM Y LIDERES VEN TODO
  // =========================

  if(PERM_isManager(user)){

    return data;
  }

if(

  PERM_isRole(user, "QA")

  ||

  PERM_isRole(user, "DEV")

){

  return data.filter(task =>

    usuarioAsignadoATarea(

      task,

      user.nombre,

      user.rol

    )

  );

}
}


//helper para las tareas asignadas

function usuarioAsignadoATarea(
  tarea,
  nombre,
  rol
){

  try{

    const asignados =

      JSON.parse(
        tarea.USUARIOS_ASIGNADOS || "{}"
      );

    const lista =

      asignados[rol] || [];

    return lista.some(x =>

      String(x)
        .trim()
        .toUpperCase()

      ===

      String(nombre)
        .trim()
        .toUpperCase()

    );

  }
  catch(error){

    Logger.log(error);

    return false;

  }

}

//fin de helper


function TASK_SERVICE_create(
  payload,
  user
){

  const id =
    generateId("TASK");

  let imageUrl = "";

  if(payload.Imagen){

    imageUrl =
      uploadImage(
        payload.Imagen,
        id
      );
  }

  const tarea = {

    ID_OBSERVACION:
      id,

    FECHA:
      new Date(),

    QA:
      payload.creadoPor,

    DEV_ASIGNADO:
      "",

    VALIDACION_RELACIONADA:
      "",

    MODULO:
      payload.Modulo,

    DESCRIPCION_CORTA:
      payload.Descripcion,

    PASOS_REPRODUCIR:
      payload.Pasos,

    CAPTURAS:
      imageUrl,

    ESTADO:
      "ABIERTO",

    PRIORIDAD:
      payload.Prioridad,

    COMENTARIO_JSON:
      "[]",

    FECHA_CORRECCION:
      "",

    FECHA_CIERRE:
      "",

    ELIMINADO:
      "",

    TIPO:
      "TAREA",

    FECHA_ASIGNACION:
      ""
  };

  DB_TASK_insert(
    tarea
  );

  _TASK_CACHE_clear();

  return {

    ok:true,

    id:id
  };
}

function getTareaById(id){

  return dbGetTareaById(id);

}


function obtenerUsuariosAsignables(){

  return DB_USUARIOS_getActivos();

}


function guardarEdicionTarea(
  payload
){

  return TASK_SERVICE_update(
    payload
  );

}

function TASK_SERVICE_update(
  payload
){

  DB_TASK_update(
    payload
  );

  return {
    ok:true
  };

}

//function TASK_SERVICE_delete(id){
//
//  return DB_TASK_delete(id);
//
//}

function TASK_SERVICE_delete(id){

  const sh =
    SpreadsheetApp
      .openById(
        CONFIG.ID_SHEET
      )
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
    id
  );

  if(rowIndex !== -1){

    sh.getRange(
      rowIndex,
      col["ELIMINADO"]
    ).setValue(
      true
    );

    _TASK_CACHE_clear();

    return {
      success:true
    };

  }

  throw new Error(
    "Tarea no encontrada"
  );

}

function TASK_SERVICE_addComment(
  payload
){

  const sheet =
    SpreadsheetApp
      .openById(
        CONFIG.ID_SHEET
      )
      .getSheetByName(
        "TAREAS"
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

  _TASK_CACHE_clear();

  return {
    success:true
  };
}


function TASK_SERVICE_restore(id){

  const sh =

    SpreadsheetApp

      .openById(
        CONFIG.ID_SHEET
      )

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
    id
  );

  if(rowIndex !== -1){

    sh.getRange(
      rowIndex,
      col["ELIMINADO"]
    ).setValue(
      false
    );

    _TASK_CACHE_clear();

    return {
      success:true
    };

  }

  throw new Error(
    "Tarea no encontrada"
  );

}