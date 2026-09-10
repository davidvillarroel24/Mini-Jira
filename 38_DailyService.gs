const _DAILY_HEADERS = [
  "ID",
  "FECHA",
  "USUARIO",
  "ROL",
  "AVANCE_HOY",
  "PROXIMO_PASO",
  "BLOQUEO",
  "NECESITO_AYUDA_DE",
  "CREADO_AT",
  "ACTUALIZADO_AT"
];

function _DAILY_norm(value){
  return String(value === null || value === undefined ? "" : value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

function _DAILY_requireUser(userFromClient){
  const user = getSession(userFromClient);

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  return user;
}

function _DAILY_ensureSheet(){
  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName("DAILY_SCRUM");

  if(!sh){
    sh = ss.insertSheet("DAILY_SCRUM");
    sh.appendRow(_DAILY_HEADERS);
  }

  const current = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];

  if(current.length === 0 || String(current[0] || "").trim() === ""){
    sh.clear();
    sh.appendRow(_DAILY_HEADERS);
  }

  return sh;
}

function _DAILY_headerMap(headers){
  const map = {};

  (headers || []).forEach((header, idx) => {
    map[_DAILY_norm(header)] = idx;
  });

  return map;
}

function _DAILY_get(row, col, key){
  const idx = col[_DAILY_norm(key)];

  if(idx === undefined){
    return "";
  }

  return row[idx];
}

function _DAILY_set(row, col, key, value){
  const idx = col[_DAILY_norm(key)];

  if(idx === undefined){
    return;
  }

  row[idx] = value;
}

function _DAILY_parseDate(value){
  if(value === null || value === undefined || value === ""){
    return null;
  }

  if(value instanceof Date){
    return isNaN(value.getTime()) ? null : value;
  }

  const raw = String(value).trim();

  if(!raw){
    return null;
  }

  // Treat yyyy-mm-dd as local date to avoid UTC shift (e.g. Bolivia GMT-4).
  const isoMatch = raw.match(/^(\d{4})-(\d{2})-(\d{2})$/);

  if(isoMatch){
    const y = Number(isoMatch[1]);
    const m = Number(isoMatch[2]);
    const d = Number(isoMatch[3]);
    const local = new Date(y, m - 1, d, 12, 0, 0, 0);

    return isNaN(local.getTime()) ? null : local;
  }

  const sanitized = raw.replace(/\s*\([^)]*\)\s*$/, "").trim();
  const dt = new Date(sanitized || raw);

  return isNaN(dt.getTime()) ? null : dt;
}

function _DAILY_dateKey(value){
  const dt = _DAILY_parseDate(value);

  if(!dt){
    return "";
  }

  return Utilities.formatDate(dt, Session.getScriptTimeZone(), "yyyy-MM-dd");
}

function _DAILY_nowText(){
  return Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
}

function _DAILY_rowToObj(headers, row, col, currentUser){
  const fechaRaw = _DAILY_get(row, col, "FECHA");
  const fecha = String(fechaRaw || "").trim();
  const fechaKey = _DAILY_dateKey(fechaRaw || fecha);
  const usuario = String(_DAILY_get(row, col, "USUARIO") || "").trim();
  const rol = String(_DAILY_get(row, col, "ROL") || "").trim();
  const necesitaAyuda = String(_DAILY_get(row, col, "NECESITO_AYUDA_DE") || "").trim();

  return {
    id: String(_DAILY_get(row, col, "ID") || "").trim(),
    fecha: fechaKey || fecha,
    fechaKey: fechaKey,
    usuario: usuario,
    rol: rol,
    avanceHoy: String(_DAILY_get(row, col, "AVANCE_HOY") || "").trim(),
    proximoPaso: String(_DAILY_get(row, col, "PROXIMO_PASO") || "").trim(),
    bloqueo: String(_DAILY_get(row, col, "BLOQUEO") || "").trim(),
    necesitoAyuda: /^(SI|NO)$/i.test(necesitaAyuda) ? necesitaAyuda.charAt(0).toUpperCase() + necesitaAyuda.slice(1).toLowerCase() : (necesitaAyuda ? "Si" : "No"),
    creadoAt: String(_DAILY_get(row, col, "CREADO_AT") || "").trim(),
    actualizadoAt: String(_DAILY_get(row, col, "ACTUALIZADO_AT") || "").trim(),
    canManage: _DAILY_norm(usuario) === _DAILY_norm(currentUser && currentUser.nombre || "")
  };
}

function _DAILY_findRowIndexById(data, col, id){
  const target = String(id || "").trim();

  for(let i = 1; i < data.length; i++){
    if(String(_DAILY_get(data[i], col, "ID") || "").trim() === target){
      return i;
    }
  }

  return -1;
}

function DAILY_hasToday(userFromClient){
  const user = _DAILY_requireUser(userFromClient);
  const sh = _DAILY_ensureSheet();
  const data = cacheGetOrSet("DAILY_ALL_V1", 20, () => sh.getDataRange().getValues());
  const headers = data[0] || [];
  const col = _DAILY_headerMap(headers);
  const today = _DAILY_dateKey(new Date());

  for(let i = 1; i < data.length; i++){
    const rowUser = String(_DAILY_get(data[i], col, "USUARIO") || "").trim();
    const rowDate = _DAILY_dateKey(_DAILY_get(data[i], col, "FECHA"));

    if(_DAILY_norm(rowUser) === _DAILY_norm(user.nombre) && rowDate === today){
      return {hasToday: true, id: String(_DAILY_get(data[i], col, "ID") || "").trim()};
    }
  }

  return {hasToday: false};
}

function DAILY_loadData(firstArg, secondArg){
  const user = _DAILY_requireUser(secondArg || firstArg);
  const sh = _DAILY_ensureSheet();
  const data = cacheGetOrSet("DAILY_ALL_V1", 20, () => sh.getDataRange().getValues());
  const headers = data[0] || [];
  const col = _DAILY_headerMap(headers);
  const rows = [];
  const helpSummary = [];
  const uniqueNames = {};
  let helpTotal = 0;

  for(let i = 1; i < data.length; i++){
    const row = data[i];
    const obj = _DAILY_rowToObj(headers, row, col, user);

    if(!obj.id){
      continue;
    }

    rows.push(obj);

    if(_DAILY_norm(obj.necesitoAyuda) === "SI"){
      helpTotal += 1;

      const key = _DAILY_norm(obj.usuario);

      if(!key || uniqueNames[key]){
        continue;
      }

      uniqueNames[key] = true;
      helpSummary.push({
        usuario: obj.usuario,
        rol: obj.rol,
        necesitoAyuda: obj.necesitoAyuda,
        fecha: obj.fecha
      });
    }
  }

  rows.sort((a, b) => String(b.actualizadoAt || b.creadoAt || b.fecha || "").localeCompare(String(a.actualizadoAt || a.creadoAt || a.fecha || "")));

  helpSummary.sort((a, b) => String(a.usuario).localeCompare(String(b.usuario)));

  return {
    data: rows,
    helpSummary: helpSummary,
    helpSummaryStats: {
      helpTotal: helpTotal,
      uniqueCount: helpSummary.length
    }
  };
}

function DAILY_save(payload, userFromClient){
  const user = _DAILY_requireUser(userFromClient);
  const sh = _DAILY_ensureSheet();
  const data = sh.getDataRange().getValues();
  const headers = data[0] || [];
  const col = _DAILY_headerMap(headers);

  const id = String(payload && payload.id || "").trim();
  const fechaInput = String(payload && payload.fecha || "").trim();
  const fecha = _DAILY_dateKey(fechaInput) || _DAILY_dateKey(new Date());
  const usuario = String(payload && payload.usuario || user.nombre || "").trim();
  const rol = String(payload && payload.rol || user.rol || "").trim();
  const avanceHoy = String(payload && payload.avanceHoy || "").trim();
  const proximoPaso = String(payload && payload.proximoPaso || "").trim();
  const bloqueo = String(payload && payload.bloqueo || "").trim();
  const necesitoAyuda = String(payload && payload.necesitoAyuda || "").trim();

  if(!avanceHoy || !proximoPaso || !bloqueo || !necesitoAyuda){
    throw new Error("VALIDATION_ERROR: debe responder las 4 preguntas");
  }

  if(_DAILY_norm(necesitoAyuda) !== "SI" && _DAILY_norm(necesitoAyuda) !== "NO"){
    throw new Error("VALIDATION_ERROR: Necesito ayuda debe ser Si o No");
  }

  let rowIndex = -1;

  if(id){
    rowIndex = _DAILY_findRowIndexById(data, col, id);

    if(rowIndex < 0){
      throw new Error("NOT_FOUND: daily scrum no encontrado");
    }

    const rowOwner = String(_DAILY_get(data[rowIndex], col, "USUARIO") || "").trim();

    if(_DAILY_norm(rowOwner) !== _DAILY_norm(user.nombre)){
      throw new Error("AUTH_FORBIDDEN: solo puedes editar tus propios daily scrum");
    }
  }else{
    for(let i = 1; i < data.length; i++){
      const rowUser = String(_DAILY_get(data[i], col, "USUARIO") || "").trim();
      const rowDate = _DAILY_dateKey(_DAILY_get(data[i], col, "FECHA"));

      if(_DAILY_norm(rowUser) === _DAILY_norm(usuario) && rowDate === _DAILY_dateKey(fecha)){
        rowIndex = i;
        break;
      }
    }
  }

  const now = _DAILY_nowText();

  if(rowIndex >= 0){
    const row = data[rowIndex];

    _DAILY_set(row, col, "FECHA", fecha);
    _DAILY_set(row, col, "USUARIO", usuario);
    _DAILY_set(row, col, "ROL", rol);
    _DAILY_set(row, col, "AVANCE_HOY", avanceHoy);
    _DAILY_set(row, col, "PROXIMO_PASO", proximoPaso);
    _DAILY_set(row, col, "BLOQUEO", bloqueo);
    _DAILY_set(row, col, "NECESITO_AYUDA_DE", _DAILY_norm(necesitoAyuda) === "SI" ? "Si" : "No");
    _DAILY_set(row, col, "ACTUALIZADO_AT", now);

    sh.getRange(rowIndex + 1, 1, 1, row.length).setValues([row]);

    cacheClear("DAILY_ALL_V1");

    return {ok: true, id: String(_DAILY_get(row, col, "ID") || "").trim()};
  }

  const newId = generateId("DS");
  const row = new Array(headers.length).fill("");

  _DAILY_set(row, col, "ID", newId);
  _DAILY_set(row, col, "FECHA", fecha);
  _DAILY_set(row, col, "USUARIO", usuario);
  _DAILY_set(row, col, "ROL", rol);
  _DAILY_set(row, col, "AVANCE_HOY", avanceHoy);
  _DAILY_set(row, col, "PROXIMO_PASO", proximoPaso);
  _DAILY_set(row, col, "BLOQUEO", bloqueo);
  _DAILY_set(row, col, "NECESITO_AYUDA_DE", _DAILY_norm(necesitoAyuda) === "SI" ? "Si" : "No");
  _DAILY_set(row, col, "CREADO_AT", now);
  _DAILY_set(row, col, "ACTUALIZADO_AT", now);

  sh.appendRow(row);

  cacheClear("DAILY_ALL_V1");

  return {ok: true, id: newId};
}
