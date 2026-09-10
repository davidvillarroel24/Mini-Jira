const _BROADCAST_HEADERS = [
  "ID",
  "TITULO",
  "MENSAJE",
  "NIVEL",
  "ALCANCE",
  "ROLES",
  "USUARIOS",
  "ESTADO",
  "ACTIVO",
  "ELIMINADO",
  "CREADO_POR",
  "FECHA_CREACION",
  "FECHA_PUBLICACION"
];

function _BROAD_norm(value){
  return String(value === null || value === undefined ? "" : value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

function _BROAD_bool(value){
  const text = _BROAD_norm(value);
  return text === "TRUE" || text === "1" || text === "SI" || text === "X";
}

function _BROAD_requireUser(userFromClient){
  const user = getSession(userFromClient);

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  return user;
}

function _BROAD_headerMap(headers){
  const map = {};

  (headers || []).forEach((header, idx) => {
    map[_BROAD_norm(header)] = idx;
  });

  return map;
}

function _BROAD_ensureSheet(){

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("BROADCAST");

  if(!sh){
    sh = ss.insertSheet("BROADCAST");
    sh.appendRow(_BROADCAST_HEADERS);
  }

  const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];

  if(current.length === 0 || String(current[0] || "").trim() === ""){
    sh.clear();
    sh.appendRow(_BROADCAST_HEADERS);
  }

  return sh;
}

function _BROAD_parseList(value){

  if(Array.isArray(value)){
    return value.map(x => String(x || "").trim()).filter(Boolean);
  }

  const raw = String(value || "").trim();

  if(!raw){
    return [];
  }

  try{
    const parsed = JSON.parse(raw);

    if(Array.isArray(parsed)){
      return parsed.map(x => String(x || "").trim()).filter(Boolean);
    }
  }catch(error){
    // fallback
  }

  return raw.split(/[;,]/).map(x => String(x || "").trim()).filter(Boolean);
}

function _BROAD_toJsonList(value){
  return JSON.stringify(_BROAD_parseList(value));
}

function _BROAD_getRawGrid(){

  const sh = _BROAD_ensureSheet();
  return cacheGetOrSet("BROADCAST_ALL_V1", 20, () => sh.getDataRange().getValues());
}

function _BROAD_readAll(){

  const sh = _BROAD_ensureSheet();
  const data = _BROAD_getRawGrid();
  const headers = data[0] || [];
  const col = _BROAD_headerMap(headers);

  return {
    sh: sh,
    data: data,
    headers: headers,
    col: col
  };
}

function _BROAD_cell(row, col, key){

  const idx = col[_BROAD_norm(key)];

  if(idx === undefined){
    return "";
  }

  return row[idx];
}

function _BROAD_setCell(row, col, key, value){

  const idx = col[_BROAD_norm(key)];

  if(idx === undefined){
    return;
  }

  row[idx] = value;
}

function _BROAD_isOwner(row, col, user){
  return _BROAD_norm(_BROAD_cell(row, col, "CREADO_POR")) === _BROAD_norm(user && user.nombre || "");
}

function _BROAD_scopeLabel(alcance, roles, users){
  const norm = _BROAD_norm(alcance || "TODOS");

  if(norm === "ROL" || norm === "ROLES"){
    return `ROL (${roles.join(", ") || "-"})`;
  }

  if(norm === "USUARIOS"){
    return `USUARIOS (${users.length})`;
  }

  return "TODOS";
}

function _BROAD_findRowIndexById(data, col, id){

  const target = String(id || "").trim();

  for(let i = 1; i < data.length; i++){
    if(String(_BROAD_cell(data[i], col, "ID") || "").trim() === target){
      return i;
    }
  }

  return -1;
}

function _BROAD_validatePublishConflicts(targetRowIndex, data, col, alcance, roles, users){

  const targetId = String(_BROAD_cell(data[targetRowIndex], col, "ID") || "").trim();

  if(alcance === "ROL"){
    const wanted = new Set((roles || []).map(x => _BROAD_norm(x)));
    const conflicts = [];

    for(let i = 1; i < data.length; i++){
      if(i === targetRowIndex){
        continue;
      }

      const row = data[i];
      const rowId = String(_BROAD_cell(row, col, "ID") || "").trim();

      if(!rowId || rowId === targetId){
        continue;
      }

      if(_BROAD_bool(_BROAD_cell(row, col, "ELIMINADO"))){
        continue;
      }

      if(!_BROAD_bool(_BROAD_cell(row, col, "ACTIVO"))){
        continue;
      }

      if(_BROAD_norm(_BROAD_cell(row, col, "ESTADO")) !== "PUBLICADO"){
        continue;
      }

      const rowScope = _BROAD_norm(_BROAD_cell(row, col, "ALCANCE"));

      if(rowScope !== "ROL" && rowScope !== "ROLES"){
        continue;
      }

      const rowRoles = _BROAD_parseList(_BROAD_cell(row, col, "ROLES")).map(_BROAD_norm);

      rowRoles.forEach(role => {
        if(wanted.has(role)){
          conflicts.push(role);
        }
      });
    }

    if(conflicts.length){
      const uniq = Array.from(new Set(conflicts));
      throw new Error(`CONFLICTO_PUBLICACION: ya existe broadcast activo para rol(es): ${uniq.join(", ")}`);
    }
  }

  if(alcance === "USUARIOS"){
    const wantedUsers = new Set((users || []).map(x => _BROAD_norm(x)));
    const conflicts = [];

    for(let i = 1; i < data.length; i++){
      if(i === targetRowIndex){
        continue;
      }

      const row = data[i];
      const rowId = String(_BROAD_cell(row, col, "ID") || "").trim();

      if(!rowId || rowId === targetId){
        continue;
      }

      if(_BROAD_bool(_BROAD_cell(row, col, "ELIMINADO"))){
        continue;
      }

      if(!_BROAD_bool(_BROAD_cell(row, col, "ACTIVO"))){
        continue;
      }

      if(_BROAD_norm(_BROAD_cell(row, col, "ESTADO")) !== "PUBLICADO"){
        continue;
      }

      const rowScope = _BROAD_norm(_BROAD_cell(row, col, "ALCANCE"));

      if(rowScope !== "USUARIOS"){
        continue;
      }

      const rowUsers = _BROAD_parseList(_BROAD_cell(row, col, "USUARIOS")).map(_BROAD_norm);

      rowUsers.forEach(userName => {
        if(wantedUsers.has(userName)){
          conflicts.push(userName);
        }
      });
    }

    if(conflicts.length){
      const uniq = Array.from(new Set(conflicts));
      throw new Error(`CONFLICTO_PUBLICACION: ya existe broadcast activo para usuario(s): ${uniq.join(", ")}`);
    }
  }
}

function BROADCAST_loadData(filters, userFromClient){

  const user = _BROAD_requireUser(userFromClient);
  const view = filters && filters.view === "ELIMINADAS" ? "ELIMINADAS" : "ACTIVAS";

  const source = _BROAD_readAll();
  const data = source.data;
  const col = source.col;

  const result = [];

  for(let i = 1; i < data.length; i++){
    const row = data[i];
    const id = String(_BROAD_cell(row, col, "ID") || "").trim();

    if(!id){
      continue;
    }

    const eliminado = _BROAD_bool(_BROAD_cell(row, col, "ELIMINADO"));

    if(view === "ELIMINADAS" ? !eliminado : eliminado){
      continue;
    }

    const roles = _BROAD_parseList(_BROAD_cell(row, col, "ROLES"));
    const usuarios = _BROAD_parseList(_BROAD_cell(row, col, "USUARIOS"));
    const alcance = String(_BROAD_cell(row, col, "ALCANCE") || "TODOS").trim().toUpperCase();
    const canManage = _BROAD_isOwner(row, col, user);

    result.push({
      id: id,
      titulo: String(_BROAD_cell(row, col, "TITULO") || "").trim(),
      mensaje: String(_BROAD_cell(row, col, "MENSAJE") || "").trim(),
      nivel: String(_BROAD_cell(row, col, "NIVEL") || "INFO").trim(),
      alcance: alcance,
      alcanceLabel: _BROAD_scopeLabel(alcance, roles, usuarios),
      roles: roles,
      usuarios: usuarios,
      estado: String(_BROAD_cell(row, col, "ESTADO") || "BORRADOR").trim(),
      activo: _BROAD_bool(_BROAD_cell(row, col, "ACTIVO")),
      creadoPor: String(_BROAD_cell(row, col, "CREADO_POR") || "").trim(),
      fechaCreacion: String(_BROAD_cell(row, col, "FECHA_CREACION") || "").trim(),
      fechaPublicacion: String(_BROAD_cell(row, col, "FECHA_PUBLICACION") || "").trim(),
      canManage: canManage
    });
  }

  const users = (obtenerUsuariosActivos() || []).map(userRow => ({
    nombre: String(userRow && userRow.nombre || "").trim(),
    rol: String(userRow && userRow.rol || "").trim()
  }));

  return {
    data: result,
    users: users
  };
}

function BROADCAST_save(payload, userFromClient){

  const user = _BROAD_requireUser(userFromClient);
  const sh = _BROAD_ensureSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = _BROAD_headerMap(headers);

  const id = String(payload && payload.id || "").trim();
  const titulo = String(payload && payload.titulo || "").trim();
  const mensaje = String(payload && payload.mensaje || "").trim();
  const nivel = String(payload && payload.nivel || "INFO").trim().toUpperCase();
  const alcance = String(payload && payload.alcance || "TODOS").trim().toUpperCase();
  const roles = _BROAD_parseList(payload && payload.roles || []);
  const usuarios = _BROAD_parseList(payload && payload.usuarios || []);

  if(!titulo || !mensaje){
    throw new Error("VALIDATION_ERROR: titulo y mensaje son obligatorios");
  }

  if(alcance === "ROL" && roles.length === 0){
    throw new Error("VALIDATION_ERROR: debe seleccionar al menos un rol");
  }

  if(alcance === "USUARIOS" && usuarios.length === 0){
    throw new Error("VALIDATION_ERROR: debe seleccionar al menos un usuario");
  }

  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  if(!id){
    const newId = generateId("BRC");
    const row = new Array(headers.length).fill("");

    _BROAD_setCell(row, col, "ID", newId);
    _BROAD_setCell(row, col, "TITULO", titulo);
    _BROAD_setCell(row, col, "MENSAJE", mensaje);
    _BROAD_setCell(row, col, "NIVEL", nivel);
    _BROAD_setCell(row, col, "ALCANCE", alcance);
    _BROAD_setCell(row, col, "ROLES", _BROAD_toJsonList(roles));
    _BROAD_setCell(row, col, "USUARIOS", _BROAD_toJsonList(usuarios));
    _BROAD_setCell(row, col, "ESTADO", "BORRADOR");
    _BROAD_setCell(row, col, "ACTIVO", false);
    _BROAD_setCell(row, col, "ELIMINADO", false);
    _BROAD_setCell(row, col, "CREADO_POR", String(user.nombre || "").trim());
    _BROAD_setCell(row, col, "FECHA_CREACION", now);
    _BROAD_setCell(row, col, "FECHA_PUBLICACION", "");

    sh.appendRow(row);

    cacheClear("BROADCAST_ALL_V1");

    return {ok: true, id: newId};
  }

  const rowIndex = findRowIndexById(sh, col.ID + 1, id);

  if(rowIndex === -1){
    throw new Error("NOT_FOUND: broadcast no encontrado");
  }

  const row = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  if(!_BROAD_isOwner(row, col, user)){
    throw new Error("AUTH_FORBIDDEN: solo el creador puede editar");
  }

  _BROAD_setCell(row, col, "TITULO", titulo);
  _BROAD_setCell(row, col, "MENSAJE", mensaje);
  _BROAD_setCell(row, col, "NIVEL", nivel);
  _BROAD_setCell(row, col, "ALCANCE", alcance);
  _BROAD_setCell(row, col, "ROLES", _BROAD_toJsonList(roles));
  _BROAD_setCell(row, col, "USUARIOS", _BROAD_toJsonList(usuarios));

  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);

  cacheClear("BROADCAST_ALL_V1");

  return {ok: true, id: id};
}

function BROADCAST_publish(payload, userFromClient){

  const user = _BROAD_requireUser(userFromClient);
  const sh = _BROAD_ensureSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = _BROAD_headerMap(headers);

  const id = String(payload && payload.id || "").trim();

  if(!id){
    throw new Error("VALIDATION_ERROR: id requerido");
  }

  const rowIndex = findRowIndexById(sh, col.ID + 1, id);

  if(rowIndex === -1){
    throw new Error("NOT_FOUND: broadcast no encontrado");
  }

  const row = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  if(!_BROAD_isOwner(row, col, user)){
    throw new Error("AUTH_FORBIDDEN: solo el creador puede publicar");
  }

  if(_BROAD_bool(_BROAD_cell(row, col, "ELIMINADO"))){
    throw new Error("VALIDATION_ERROR: no puede publicar un broadcast eliminado");
  }

  const alcance = _BROAD_norm(_BROAD_cell(row, col, "ALCANCE") || "TODOS");
  const roles = _BROAD_parseList(_BROAD_cell(row, col, "ROLES"));
  const usuarios = _BROAD_parseList(_BROAD_cell(row, col, "USUARIOS"));

  // El chequeo de conflictos cruza todas las filas: se apoya en la lista
  // cacheada (misma que usa la vista de lista), no en una lectura fresca.
  const allData = _BROAD_getRawGrid();
  _BROAD_validatePublishConflicts(rowIndex - 1, allData, col, alcance, roles, usuarios);

  _BROAD_setCell(row, col, "ESTADO", "PUBLICADO");
  _BROAD_setCell(row, col, "ACTIVO", true);
  _BROAD_setCell(row, col, "FECHA_PUBLICACION", Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss"));

  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);

  cacheClear("BROADCAST_ALL_V1");

  return {ok: true, id: id};
}

function BROADCAST_delete(id, userFromClient){

  const user = _BROAD_requireUser(userFromClient);
  const sh = _BROAD_ensureSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = _BROAD_headerMap(headers);

  const rowIndex = findRowIndexById(sh, col.ID + 1, id);

  if(rowIndex === -1){
    throw new Error("NOT_FOUND: broadcast no encontrado");
  }

  const row = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  if(!_BROAD_isOwner(row, col, user)){
    throw new Error("AUTH_FORBIDDEN: solo el creador puede eliminar");
  }

  _BROAD_setCell(row, col, "ELIMINADO", true);
  _BROAD_setCell(row, col, "ACTIVO", false);

  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);

  cacheClear("BROADCAST_ALL_V1");

  return {ok: true};
}

function BROADCAST_restore(id, userFromClient){

  const user = _BROAD_requireUser(userFromClient);
  const sh = _BROAD_ensureSheet();
  const headers = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0];
  const col = _BROAD_headerMap(headers);

  const rowIndex = findRowIndexById(sh, col.ID + 1, id);

  if(rowIndex === -1){
    throw new Error("NOT_FOUND: broadcast no encontrado");
  }

  const row = sh.getRange(rowIndex, 1, 1, headers.length).getValues()[0];

  if(!_BROAD_isOwner(row, col, user)){
    throw new Error("AUTH_FORBIDDEN: solo el creador puede restaurar");
  }

  _BROAD_setCell(row, col, "ELIMINADO", false);

  sh.getRange(rowIndex, 1, 1, row.length).setValues([row]);

  cacheClear("BROADCAST_ALL_V1");

  return {ok: true};
}
