function obtenerUsuariosActivos() {

  const result = DB_USERS_getAll()

    .filter(obj =>

      String(obj.Activo)
        .trim()
        .toUpperCase()

      ===

      "TRUE"
    )

    .map(obj => ({
      id: obj.ID,
      nombre: obj.Nombre,
      rol: obj.Rol_principal
    }));

  Logger.log(result);
  return result;
}

function obtenerDevs() {

  return obtenerUsuariosActivos()

    .filter(x =>

      String(
        x.ROL_PRINCIPAL
      ).toUpperCase()

      ===

      "DEV"
    )

    .filter(x =>

      String(
        x.ACTIVO
      ).toUpperCase()

      ===

      "TRUE"
    );
}


function DB_USUARIOS_getActivos(){

  return DB_USERS_getAll().filter(obj =>

    String(obj.Activo)
      .toUpperCase()

    ===

    "TRUE"
  );
}

function prueba123(){
  Logger.log(DB_USUARIOS_getActivos())
}

function _normalizarHeaderUsers(header){

  return String(header || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

function ADMIN_USERS_toggleActive(payload, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("No autenticado");
  }

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName("USUARIOS");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const colId = headers.findIndex(h => _normalizarHeaderUsers(h) === "ID");
  const colNombre = headers.findIndex(h => _normalizarHeaderUsers(h) === "NOMBRE");
  const colActivo = headers.findIndex(h => _normalizarHeaderUsers(h) === "ACTIVO");

  if(colActivo === -1){
    throw new Error("No existe la columna Activo en USUARIOS");
  }

  const targetId = String(payload && (payload.id || payload.ID) || "").trim();
  const targetNombre = String(payload && (payload.nombre || payload.Nombre) || "").trim();
  const nextActivo = !!(payload && payload.activo);

  const rowIndex = (targetId && colId >= 0)
    ? findRowIndexById(sh, colId + 1, targetId)
    : ((targetNombre && colNombre >= 0) ? findRowIndexById(sh, colNombre + 1, targetNombre) : -1);

  if(rowIndex === -1){
    throw new Error("Usuario no encontrado");
  }

  const row = sh.getRange(rowIndex, 1, 1, sh.getLastColumn()).getValues()[0];
  const rowId = colId >= 0 ? String(row[colId] || "").trim() : "";
  const rowNombre = colNombre >= 0 ? String(row[colNombre] || "").trim().toUpperCase() : "";

  sh.getRange(rowIndex, colActivo + 1).setValue(nextActivo);

  _USERS_CACHE_clear();

  return {
    ok: true,
    id: rowId,
    nombre: rowNombre,
    activo: nextActivo
  };
}

function USER_PREF_getDashboard(contexto, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("No autenticado");
  }

  const targetNombre = String(user.nombre || "").trim().toUpperCase();

  const found = DB_USERS_getAll().find(obj =>
    String(obj.Nombre || "").trim().toUpperCase() === targetNombre
  );

  if(!found){
    throw new Error("Usuario no encontrado");
  }

  let prefRaw = String(found.Preferencias || "").trim();
  let prefs = {};

  if(prefRaw){
    try{
      prefs = JSON.parse(prefRaw);
    }catch(error){
      Logger.log(error);
      prefs = {};
    }
  }

  const dashboard = prefs.DASHBOARD || {};

  return dashboard[contexto] || {
    visibleColumns: [],
    estado: "TODOS",
    prioridad: "TODOS"
  };
}

function USER_PREF_saveDashboard(contexto, payload, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("No autenticado");
  }

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName("USUARIOS");

  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const colPref = headers.findIndex(h =>
    _normalizarHeaderUsers(h) === "PREFERENCIAS"
  );

  const colNombre = headers.findIndex(h =>
    _normalizarHeaderUsers(h) === "NOMBRE"
  );

  if(colPref === -1){
    throw new Error("No existe la columna Preferencias en USUARIOS");
  }

  if(colNombre === -1){
    throw new Error("No existe la columna Nombre en USUARIOS");
  }

  const targetNombre = String(user.nombre || "").trim();
  const rowIndex = findRowIndexById(sh, colNombre + 1, targetNombre);

  if(rowIndex === -1){
    throw new Error("Usuario no encontrado");
  }

  let prefRaw = String(sh.getRange(rowIndex, colPref + 1).getValue() || "").trim();
  let prefs = {};

  if(prefRaw){
    try{
      prefs = JSON.parse(prefRaw);
    }catch(error){
      Logger.log(error);
      prefs = {};
    }
  }

  if(!prefs.DASHBOARD){
    prefs.DASHBOARD = {};
  }

  prefs.DASHBOARD[contexto] = {
    visibleColumns: Array.isArray(payload.visibleColumns) ? payload.visibleColumns : [],
    estado: payload.estado || "TODOS",
    prioridad: payload.prioridad || "TODOS",
    updatedAt: Utilities.formatDate(
      new Date(),
      Session.getScriptTimeZone(),
      "yyyy-MM-dd HH:mm:ss"
    )
  };

  sh.getRange(rowIndex, colPref + 1).setValue(
    JSON.stringify(prefs)
  );

  _USERS_CACHE_clear();

  return {
    ok: true
  };
}
