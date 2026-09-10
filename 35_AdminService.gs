function _ADMIN_requireUser(userFromClient){
  const user = getSession(userFromClient);
  if(!user){
    throw new Error("AUTH_REQUIRED");
  }
  return user;
}

function _ADMIN_normHeader(header){
  return String(header || "")
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

function _ADMIN_headerMap(headers){
  const map = {};

  (headers || []).forEach((h, i) => {
    map[_ADMIN_normHeader(h)] = i;
  });

  return map;
}

function _ADMIN_buildRowObject(headers, row, mapping){
  const result = {};

  Object.keys(mapping || {}).forEach(key => {
    const headerNorm = mapping[key];
    const idx = headers.indexOf(headerNorm);

    if(idx >= 0){
      result[key] = row[idx];
    }
  });

  return result;
}

function _ADMIN_getColumnOrderFromHeaders(headers, mapping){

  const normalizedToKey = {};

  Object.keys(mapping || {}).forEach(key => {
    normalizedToKey[_ADMIN_normHeader(mapping[key])] = key;
  });

  return (headers || [])
    .map(h => _ADMIN_normHeader(h))
    .filter(h => normalizedToKey.hasOwnProperty(h))
    .map(h => normalizedToKey[h]);
}

function _ADMIN_getValue(row, headersMap, headerName){
  const idx = headersMap[_ADMIN_normHeader(headerName)];

  if(idx === undefined){
    return "";
  }

  return row[idx];
}

function _ADMIN_pickPayloadValue(payload, keys){
  const source = payload && typeof payload === "object" ? payload : {};
  const lookup = Array.isArray(keys) ? keys : [keys];

  for(let i = 0; i < lookup.length; i++){
    const key = lookup[i];

    if(key !== null && key !== undefined && Object.prototype.hasOwnProperty.call(source, key)){
      return source[key];
    }
  }

  const normalizedTargets = lookup
    .filter(key => key !== null && key !== undefined)
    .map(key => _ADMIN_normHeader(key));
  const sourceKeys = Object.keys(source);

  for(let i = 0; i < sourceKeys.length; i++){
    const sourceKey = sourceKeys[i];

    if(normalizedTargets.includes(_ADMIN_normHeader(sourceKey))){
      return source[sourceKey];
    }
  }

  return "";
}

function _ADMIN_ensureSheet(name, headers){

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);

  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }

  const currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  if(currentHeaders.length === 0 || String(currentHeaders[0] || "").trim() === ""){
    sh.clear();
    sh.appendRow(headers);
  }
  //Logger.log(currentHeaders)
  return sh;
}
function prueba001(){
  _ADMIN_ensureSheet ("USUARIOS",["ID", "Nombre", "Rol_principal", "Roles_permitidos", "QA_asignado", "Preferencias", "Activo"])
}

function _ADMIN_toIsoDateTimeLocal(value){

  if(!value){
    return "";
  }

  let dt = value;

  if(!(dt instanceof Date)){
    dt = new Date(value);
  }

  if(isNaN(dt.getTime())){
    return "";
  }

  const y = dt.getFullYear();
  const m = String(dt.getMonth() + 1).padStart(2, "0");
  const d = String(dt.getDate()).padStart(2, "0");
  const hh = String(dt.getHours()).padStart(2, "0");
  const mm = String(dt.getMinutes()).padStart(2, "0");

  return `${y}-${m}-${d}T${hh}:${mm}`;
}

function _ADMIN_isDeleted(value){
  const normalized = String(value === null || value === undefined ? "" : value).trim().toLowerCase();

  return normalized === "true"
    || normalized === "1"
    || normalized === "si"
    || normalized === "sí"
    || normalized === "x";
}

function _ADMIN_resolveFiltersAndUser(firstArg, secondArg){
  const defaultFilters = {view: "ACTIVAS"};

  if(firstArg && typeof firstArg === "object" && (Object.prototype.hasOwnProperty.call(firstArg, "view") || Object.prototype.hasOwnProperty.call(firstArg, "includeDeleted"))){
    return {
      filters: Object.assign({}, defaultFilters, firstArg),
      userFromClient: secondArg
    };
  }

  return {
    filters: defaultFilters,
    userFromClient: firstArg
  };
}

function ADMIN_USERS_list(){
//userFromClient
  //_ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_ensureSheet(
    "USUARIOS",
    ["ID", "Nombre", "Rol_principal", "Roles_permitidos", "QA_asignado", "Preferencias", "Activo"]
  );

  const data = cacheGetOrSet("USUARIOS_RAW_GRID_V1", 20, () => sh.getDataRange().getValues());

  if (data.length === 0) {
    return {
      data: [],
      columnOrder: []
    };
  }

  const headers = data[0];
  //Logger.log(headers);
  //Logger.log(headers.length);
  //Logger.log(JSON.stringify(headers));

  const col = _ADMIN_headerMap(headers);

  //Logger.log(JSON.stringify(col));

  const map = {
    ID: "ID",
    Nombre: "Nombre",
    Rol_principal: "Rol_principal",
    Roles_permitidos: "Roles_permitidos",
    QA_asignado: "QA_asignado",
    Preferencias: "Preferencias",
    Activo: "Activo"
  };

  const columnOrder = _ADMIN_getColumnOrderFromHeaders(headers, map);

  Logger.log(JSON.stringify(columnOrder));

  const result = [];

  for (let i = 1; i < data.length; i++) {

    const row = data[i];

    if (!String(_ADMIN_getValue(row, col, "Nombre") || "").trim()) {
      continue;
    }

    result.push({
      ID: String(_ADMIN_getValue(row, col, "ID") || "").trim(),
      Nombre: String(_ADMIN_getValue(row, col, "Nombre") || "").trim(),
      Rol_principal: String(_ADMIN_getValue(row, col, "Rol_principal") || "").trim(),
      Roles_permitidos: String(_ADMIN_getValue(row, col, "Roles_permitidos") || "").trim(),
      QA_asignado: String(_ADMIN_getValue(row, col, "QA_asignado") || "").trim(),
      Preferencias: String(_ADMIN_getValue(row, col, "Preferencias") || "").trim(),
      Activo: _ADMIN_getValue(row, col, "Activo")
    });

  }
  Logger.log(headers)
  Logger.log(columnOrder)
  return {
    data: result,
    columnOrder: columnOrder
  };
}

function prueba1(){
  ADMIN_USERS_list()
}
function ADMIN_USERS_saveSprintRoles(payload, userFromClient){

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_ensureSheet("USUARIOS", ["ID", "Nombre", "Rol_principal", "Roles_permitidos", "QA_asignado", "Preferencias", "Activo"]);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];

  const col = _ADMIN_headerMap(headers);

  const nombre = String(_ADMIN_pickPayloadValue(payload, ["Nombre", "nombre"]) || "").trim();

  if(!nombre){
    throw new Error("VALIDATION_ERROR: nombre requerido");
  }

  let sprintRoles = {};

  try{
    sprintRoles = JSON.parse(String(payload && payload.sprintRolesRaw || "{}"));
  }catch(error){
    throw new Error("VALIDATION_ERROR: JSON invalido en roles por sprint");
  }

  if(col.NOMBRE === undefined){
    throw new Error("USUARIOS_SCHEMA_ERROR: falta columna Nombre en USUARIOS");
  }

  if(col.PREFERENCIAS === undefined){
    throw new Error("USUARIOS_SCHEMA_ERROR: falta columna Preferencias en USUARIOS");
  }

  const rowIndex = findRowIndexById(sh, col.NOMBRE + 1, nombre);

  if(rowIndex === -1){
    throw new Error("NOT_FOUND: usuario no encontrado");
  }

  const prefCell = String(sh.getRange(rowIndex, col.PREFERENCIAS + 1).getValue() || "").trim();
  let prefs = {};

  if(prefCell){
    try{
      prefs = JSON.parse(prefCell);
    }catch(error){
      Logger.log(error);
    }
  }

  prefs.SPRINT_ROLES = sprintRoles;

  sh.getRange(rowIndex, col.PREFERENCIAS + 1).setValue(JSON.stringify(prefs));

  _USERS_CACHE_clear();

  return {ok: true};
}

function BACKLOG_list(firstArg, secondArg){

  const input = _ADMIN_resolveFiltersAndUser(firstArg, secondArg);
  const filters = input.filters || {view: "ACTIVAS"};
  const userFromClient = input.userFromClient;

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_ensureSheet("BACKLOG", ["ID", "Descripción", "Sprint", "Prioridad", "Esfuerzo", "Verificación", "DEV", "QA", "Eliminado"]);
  const data = cacheGetOrSet("BACKLOG_ALL_V1", 20, () => sh.getDataRange().getValues());
  const headers = (data[0] || []).map(h => _ADMIN_normHeader(h));

  const map = {
    ID: "ID",
    Descripción: "Descripción",
    Sprint: "Sprint",
    Prioridad: "Prioridad",
    Esfuerzo: "Esfuerzo",
    Verificación: "Verificación",
    DEV: "DEV",
    QA: "QA",
    Eliminado: "Eliminado"
  };

  const colOrder = _ADMIN_getColumnOrderFromHeaders(data[0] || [], map);
  const col = _ADMIN_headerMap(headers);

  const result = [];

  for(let i = 1; i < data.length; i++){

    const id = col.ID >= 0 ? String(data[i][col.ID] || "").trim() : "";

    if(!id){
      continue;
    }

    const isDeleted = col.ELIMINADO !== undefined ? _ADMIN_isDeleted(data[i][col.ELIMINADO]) : false;

    if(filters.view === "ELIMINADAS" ? !isDeleted : isDeleted){
      continue;
    }

    result.push({
      ID: id,
      Descripción: col.DESCRIPCION >= 0 ? String(data[i][col.DESCRIPCION] || "").trim() : "",
      Sprint: col.SPRINT >= 0 ? String(data[i][col.SPRINT] || "").trim() : "",
      Prioridad: col.PRIORIDAD >= 0 ? String(data[i][col.PRIORIDAD] || "").trim() : "",
      Esfuerzo: col.ESFUERZO >= 0 ? String(data[i][col.ESFUERZO] || "").trim() : "",
      Verificación: col.VERIFICACION >= 0 ? String(data[i][col.VERIFICACION] || "").trim() : "",
      DEV: col.DEV >= 0 ? String(data[i][col.DEV] || "").trim() : "",
      QA: col.QA >= 0 ? String(data[i][col.QA] || "").trim() : "",
      eliminado: isDeleted
    });
  }

  return result
    .sort((a, b) => String(b.ID).localeCompare(String(a.ID)))
    .map(item => {
      const filtered = {};

      colOrder.forEach(k => {
        filtered[k] = item[k];
      });

      filtered.ID = item.ID;
      filtered.Descripción = item.Descripción;
      filtered.Sprint = item.Sprint;
      filtered.Prioridad = item.Prioridad;
      filtered.Esfuerzo = item.Esfuerzo;
      filtered.Verificación = item.Verificación;
      filtered.DEV = item.DEV;
      filtered.QA = item.QA;
      filtered.eliminado = item.eliminado;

      return filtered;
    });
}

function BACKLOG_loadData(firstArg, secondArg){

  const input = _ADMIN_resolveFiltersAndUser(firstArg, secondArg);
  const filters = input.filters || {view: "ACTIVAS"};
  const userFromClient = input.userFromClient;

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_ensureSheet("BACKLOG", ["ID", "Descripción", "Sprint", "Prioridad", "Esfuerzo", "Verificación", "DEV", "QA", "Eliminado"]);
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];

  const map = {
    ID: "ID",
    Descripción: "Descripción",
    Sprint: "Sprint",
    Prioridad: "Prioridad",
    Esfuerzo: "Esfuerzo",
    Verificación: "Verificación",
    DEV: "DEV",
    QA: "QA",
    Eliminado: "Eliminado"
  };

  const columnOrder = _ADMIN_getColumnOrderFromHeaders(headers, map);

  return {
    data: BACKLOG_list(filters, userFromClient),
    columnOrder: columnOrder
  };
}

function BACKLOG_save(payload, userFromClient){

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_ensureSheet("BACKLOG", ["ID", "Descripción", "Sprint", "Prioridad", "Esfuerzo", "Verificación", "DEV", "QA", "Eliminado"]);
  const headers = (sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || []).map(h => _ADMIN_normHeader(h));
  const col = _ADMIN_headerMap(headers);

  const id = String(_ADMIN_pickPayloadValue(payload, ["ID", "id"]) || "").trim() || generateId("BACKLOG");
  const descripcion = String(_ADMIN_pickPayloadValue(payload, ["Descripción", "descripcion", "nombre"]) || "").trim();
  const sprint = String(_ADMIN_pickPayloadValue(payload, ["Sprint", "sprint", "estado"]) || "").trim();
  const prioridad = String(_ADMIN_pickPayloadValue(payload, ["Prioridad", "prioridad"]) || "").trim();
  const esfuerzo = String(_ADMIN_pickPayloadValue(payload, ["Esfuerzo", "esfuerzo"]) || "").trim();
  const verificacion = String(_ADMIN_pickPayloadValue(payload, ["Verificación", "verificacion", "objetivo"]) || "").trim();
  const dev = String(_ADMIN_pickPayloadValue(payload, ["DEV", "dev"]) || "").trim();
  const qa = String(_ADMIN_pickPayloadValue(payload, ["QA", "qa"]) || "").trim();

  if(col.DESCRIPCION === undefined){
    throw new Error("BACKLOG_SCHEMA_ERROR: falta columna Descripción en BACKLOG");
  }

  if(!descripcion){
    throw new Error("VALIDATION_ERROR: descripción requerida");
  }

  const rowIndex = col.ID !== undefined ? findRowIndexById(sh, col.ID + 1, id) : -1;

  if(rowIndex !== -1){

    if(col.DESCRIPCION !== undefined){
      sh.getRange(rowIndex, col.DESCRIPCION + 1).setValue(descripcion);
    }
    if(col.SPRINT !== undefined){
      sh.getRange(rowIndex, col.SPRINT + 1).setValue(sprint);
    }
    if(col.PRIORIDAD !== undefined){
      sh.getRange(rowIndex, col.PRIORIDAD + 1).setValue(prioridad);
    }
    if(col.ESFUERZO !== undefined){
      sh.getRange(rowIndex, col.ESFUERZO + 1).setValue(esfuerzo);
    }
    if(col.VERIFICACION !== undefined){
      sh.getRange(rowIndex, col.VERIFICACION + 1).setValue(verificacion);
    }
    if(col.DEV !== undefined){
      sh.getRange(rowIndex, col.DEV + 1).setValue(dev);
    }
    if(col.QA !== undefined){
      sh.getRange(rowIndex, col.QA + 1).setValue(qa);
    }
    if(col.ELIMINADO !== undefined){
      sh.getRange(rowIndex, col.ELIMINADO + 1).setValue(false);
    }

    cacheClear("BACKLOG_ALL_V1");

    return {ok: true, id: id};
  }

  const row = new Array(headers.length).fill("");

  if(col.ID !== undefined){
    row[col.ID] = id;
  }
  if(col.ID !== undefined){
    row[col.ID] = id;
  }
  if(col.DESCRIPCION !== undefined){
    row[col.DESCRIPCION] = descripcion;
  }
  if(col.SPRINT !== undefined){
    row[col.SPRINT] = sprint;
  }
  if(col.PRIORIDAD !== undefined){
    row[col.PRIORIDAD] = prioridad;
  }
  if(col.ESFUERZO !== undefined){
    row[col.ESFUERZO] = esfuerzo;
  }
  if(col.VERIFICACION !== undefined){
    row[col.VERIFICACION] = verificacion;
  }

  if(col.DEV !== undefined){
    row[col.DEV] = dev;
  }

  if(col.QA !== undefined){
    row[col.QA] = qa;
  }

  if(col.ELIMINADO !== undefined){
    row[col.ELIMINADO] = false;
  }

  sh.appendRow(row);

  cacheClear("BACKLOG_ALL_V1");

  return {ok: true, id: id};
}

function BACKLOG_delete(id, userFromClient){

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_ensureSheet("BACKLOG", ["ID", "Descripción", "Sprint", "Prioridad", "Esfuerzo", "Verificación", "DEV", "QA", "Eliminado"]);
  const headers = (sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || []).map(h => _ADMIN_normHeader(h));
  const col = _ADMIN_headerMap(headers);
  const targetId = String(id || "").trim();

  if(!targetId){
    throw new Error("VALIDATION_ERROR: id requerido");
  }

  if(col.ID === undefined){
    throw new Error("BACKLOG_SCHEMA_ERROR: falta columna ID en BACKLOG");
  }

  if(col.ELIMINADO === undefined){
    throw new Error("BACKLOG_SCHEMA_ERROR: falta columna ELIMINADO en BACKLOG");
  }

  const rowIndex = findRowIndexById(sh, col.ID + 1, targetId);

  if(rowIndex === -1){
    throw new Error("NOT_FOUND: sprint no encontrado");
  }

  sh.getRange(rowIndex, col.ELIMINADO + 1).setValue(true);

  cacheClear("BACKLOG_ALL_V1");

  return {ok: true, id: targetId};
}

function _ADMIN_getCalendarSheet(){

  const headers = ["Módulo", "Actividades (RQ)", "Sprint", "Inicio", "Fin", "Duración", "Eliminado"];
  return _ADMIN_ensureSheet("CALENDARIO", headers);
}

function CALENDAR_list(firstArg, secondArg){

  const input = _ADMIN_resolveFiltersAndUser(firstArg, secondArg);
  const filters = input.filters || {view: "ACTIVAS"};
  const userFromClient = input.userFromClient;

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_getCalendarSheet();
  const data = cacheGetOrSet("CALENDARIO_ALL_V1", 20, () => sh.getDataRange().getValues());
  const headers = (data[0] || []).map(h => _ADMIN_normHeader(h));

  const map = {
    Módulo: "Módulo",
    "Actividades (RQ)": "Actividades (RQ)",
    Sprint: "Sprint",
    Inicio: "Inicio",
    Fin: "Fin",
    Duración: "Duración",
    Eliminado: "Eliminado"
  };

  const colOrder = _ADMIN_getColumnOrderFromHeaders(data[0] || [], map);
  const col = _ADMIN_headerMap(headers);

  const result = [];

  for(let i = 1; i < data.length; i++){

    const isDeleted = col.ELIMINADO !== undefined ? _ADMIN_isDeleted(data[i][col.ELIMINADO]) : false;

    if(filters.view === "ELIMINADAS" ? !isDeleted : isDeleted){
      continue;
    }

    result.push({
      __rowNumber: i + 1,
      "Módulo": col.MODULO >= 0 ? String(data[i][col.MODULO] || "").trim() : "",
      "Actividades (RQ)": col["ACTIVIDADES (RQ)"] >= 0 ? String(data[i][col["ACTIVIDADES (RQ)"]] || "").trim() : "",
      Sprint: col.SPRINT >= 0 ? String(data[i][col.SPRINT] || "").trim() : "",
      Inicio: col.INICIO >= 0 && data[i][col.INICIO]
        ? Utilities.formatDate(new Date(data[i][col.INICIO]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
        : "",
      Fin: col.FIN >= 0 && data[i][col.FIN]
        ? Utilities.formatDate(new Date(data[i][col.FIN]), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm")
        : "",
      "Duración": col.DURACION >= 0 ? String(data[i][col.DURACION] || "").trim() : "",
      eliminado: isDeleted
    });
  }

  return result
    .sort((a, b) => String(b["Inicio"]).localeCompare(String(a["Inicio"])))
    .map(item => {
      const filtered = {};

      colOrder.forEach(k => {
        filtered[k] = item[k];
      });

      filtered["Módulo"] = item["Módulo"];
      filtered["Actividades (RQ)"] = item["Actividades (RQ)"];
      filtered.Sprint = item.Sprint;
      filtered.Inicio = item.Inicio;
      filtered.Fin = item.Fin;
      filtered["Duración"] = item["Duración"];
      filtered.__rowNumber = item.__rowNumber;
      filtered.eliminado = item.eliminado;

      return filtered;
    });
}

function CALENDAR_loadData(firstArg, secondArg){

  const input = _ADMIN_resolveFiltersAndUser(firstArg, secondArg);
  const filters = input.filters || {view: "ACTIVAS"};
  const userFromClient = input.userFromClient;

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_getCalendarSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];

  const map = {
    Módulo: "Módulo",
    "Actividades (RQ)": "Actividades (RQ)",
    Sprint: "Sprint",
    Inicio: "Inicio",
    Fin: "Fin",
    Duración: "Duración",
    Eliminado: "Eliminado"
  };

  const columnOrder = _ADMIN_getColumnOrderFromHeaders(headers, map);

  return {
    data: CALENDAR_list(filters, userFromClient),
    columnOrder: columnOrder
  };
}

function CALENDAR_save(payload, userFromClient){

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_getCalendarSheet();
  const lastRow = sh.getLastRow();
  const headers = (sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || []).map(h => _ADMIN_normHeader(h));
  const col = _ADMIN_headerMap(headers);

  const rowNumber = Number(_ADMIN_pickPayloadValue(payload, ["__rowNumber", "rowNumber", "id"]) || 0);
  const modulo = String(_ADMIN_pickPayloadValue(payload, ["Módulo", "modulo", "titulo"]) || "").trim();
  const actividades = String(_ADMIN_pickPayloadValue(payload, ["Actividades (RQ)", "actividades", "tipo", "detalle"]) || "").trim();
  const sprint = String(_ADMIN_pickPayloadValue(payload, ["Sprint", "sprint"]) || "").trim();
  const inicio = _ADMIN_pickPayloadValue(payload, ["Inicio", "inicio"]);
  const fin = _ADMIN_pickPayloadValue(payload, ["Fin", "fin"]);
  const duracion = String(_ADMIN_pickPayloadValue(payload, ["Duración", "duracion"]) || "").trim();

  if(col.MODULO === undefined){
    throw new Error("CALENDAR_SCHEMA_ERROR: falta columna Módulo en CALENDARIO");
  }

  if(!modulo){
    throw new Error("VALIDATION_ERROR: módulo requerido");
  }

  if(rowNumber > 1 && rowNumber <= lastRow){
    if(col.MODULO !== undefined){
      sh.getRange(rowNumber, col.MODULO + 1).setValue(modulo);
    }
    if(col["ACTIVIDADES (RQ)"] !== undefined){
      sh.getRange(rowNumber, col["ACTIVIDADES (RQ)"] + 1).setValue(actividades);
    }
    if(col.SPRINT !== undefined){
      sh.getRange(rowNumber, col.SPRINT + 1).setValue(sprint);
    }
    if(col.INICIO !== undefined){
      sh.getRange(rowNumber, col.INICIO + 1).setValue(inicio || "");
    }
    if(col.FIN !== undefined){
      sh.getRange(rowNumber, col.FIN + 1).setValue(fin || "");
    }
    if(col.DURACION !== undefined){
      sh.getRange(rowNumber, col.DURACION + 1).setValue(duracion);
    }
    if(col.ELIMINADO !== undefined){
      sh.getRange(rowNumber, col.ELIMINADO + 1).setValue(false);
    }

    cacheClear("CALENDARIO_ALL_V1");

    return {ok: true, id: String(rowNumber)};
  }

  if(rowNumber > 1 && rowNumber <= lastRow){
    if(col.MODULO !== undefined){
      sh.getRange(rowNumber, col.MODULO + 1).setValue(modulo);
    }
    if(col["ACTIVIDADES (RQ)"] !== undefined){
      sh.getRange(rowNumber, col["ACTIVIDADES (RQ)"] + 1).setValue(actividades);
    }
    if(col.SPRINT !== undefined){
      sh.getRange(rowNumber, col.SPRINT + 1).setValue(sprint);
    }
    if(col.INICIO !== undefined){
      sh.getRange(rowNumber, col.INICIO + 1).setValue(inicio || "");
    }
    if(col.FIN !== undefined){
      sh.getRange(rowNumber, col.FIN + 1).setValue(fin || "");
    }
    if(col.DURACION !== undefined){
      sh.getRange(rowNumber, col.DURACION + 1).setValue(duracion);
    }
    if(col.ELIMINADO !== undefined){
      sh.getRange(rowNumber, col.ELIMINADO + 1).setValue(false);
    }

    return {ok: true, id: String(rowNumber)};
  }

  const row = new Array(headers.length).fill("");

  if(col.MODULO !== undefined){
    row[col.MODULO] = modulo;
  }
  if(col["ACTIVIDADES (RQ)"] !== undefined){
    row[col["ACTIVIDADES (RQ)"]] = actividades;
  }

  if(col.SPRINT !== undefined){
    row[col.SPRINT] = sprint;
  }

  if(col.INICIO !== undefined){
    row[col.INICIO] = inicio || "";
  }

  if(col.FIN !== undefined){
    row[col.FIN] = fin || "";
  }

  if(col.DURACION !== undefined){
    row[col.DURACION] = duracion;
  }

  if(col.ELIMINADO !== undefined){
    row[col.ELIMINADO] = false;
  }

  sh.appendRow(row);

  cacheClear("CALENDARIO_ALL_V1");

  return {ok: true, id: String(lastRow + 1)};
}

function CALENDAR_delete(id, userFromClient){

  _ADMIN_requireUser(userFromClient);

  const sh = _ADMIN_getCalendarSheet();
  const headers = (sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || []).map(h => _ADMIN_normHeader(h));
  const col = _ADMIN_headerMap(headers);
  const targetRow = Number(id || 0);

  if(!targetRow){
    throw new Error("VALIDATION_ERROR: fila requerida");
  }

  if(targetRow < 2 || targetRow > sh.getLastRow()){
    throw new Error("NOT_FOUND: evento no encontrado");
  }

  if(col.ELIMINADO === undefined){
    throw new Error("CALENDAR_SCHEMA_ERROR: falta columna Eliminado en CALENDARIO");
  }

  sh.getRange(targetRow, col.ELIMINADO + 1).setValue(true);

  cacheClear("CALENDARIO_ALL_V1");

  return {ok: true, id: String(targetRow)};
}
