function _NOTIF_normText(value){
  return String(value === null || value === undefined ? "" : value)
    .normalize("NFD")
    .replace(/[\u0300-\u036f]/g, "")
    .trim()
    .toUpperCase();
}

function _NOTIF_bool(value){
  const text = _NOTIF_normText(value);
  return text === "TRUE" || text === "1" || text === "SI" || text === "X";
}

function _NOTIF_getAny(record, keys){

  if(!record || typeof record !== "object"){
    return "";
  }

  const byNorm = {};

  Object.keys(record).forEach(key => {
    byNorm[_NOTIF_normText(key)] = record[key];
  });

  for(let i = 0; i < keys.length; i++){
    const key = _NOTIF_normText(keys[i]);

    if(Object.prototype.hasOwnProperty.call(byNorm, key)){
      return byNorm[key];
    }
  }

  return "";
}

function _NOTIF_toDate(value){

  if(!value){
    return null;
  }

  if(value instanceof Date){
    return isNaN(value.getTime()) ? null : value;
  }

  const text = String(value).trim();

  if(!text){
    return null;
  }

  // Remove locale hints in parentheses like "(hora de Bolivia)" for robust parsing.
  const sanitized = text.replace(/\s*\([^)]*\)\s*$/, "").trim();

  const isoCandidate = sanitized.indexOf(" ") > -1 ? sanitized.replace(" ", "T") : sanitized;
  let dt = new Date(isoCandidate);

  if(!isNaN(dt.getTime())){
    return dt;
  }

  const m = sanitized.match(/^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/);

  if(m){
    const day = Number(m[1]);
    const month = Number(m[2]) - 1;
    const year = Number(m[3]);
    const hh = Number(m[4] || 0);
    const mm = Number(m[5] || 0);
    const ss = Number(m[6] || 0);

    dt = new Date(year, month, day, hh, mm, ss);
    return isNaN(dt.getTime()) ? null : dt;
  }

  return null;
}

function _NOTIF_ensureSheet(name, headers){

  const ss = SpreadsheetApp.getActive();
  let sh = ss.getSheetByName(name);

  if(!sh){
    sh = ss.insertSheet(name);
    sh.appendRow(headers);
  }

  const currentHeaders = sh.getRange(1, 1, 1, sh.getLastColumn()).getValues()[0] || [];

  if(currentHeaders.length === 0 || String(currentHeaders[0] || "").trim() === ""){
    sh.clear();
    sh.appendRow(headers);
  }

  return sh;
}

function _NOTIF_ensureBroadcastAckSheet(){
  return _NOTIF_ensureSheet("BROADCAST_ACKS", ["USER", "ROL", "BROADCAST_ID", "PUBLICACION_AT", "ACK_AT"]);
}

function _NOTIF_readMapForUser(user){

  const sh = _NOTIF_ensureSheet("CHAT_READS", ["USER", "ROL", "TIPO", "ITEM_ID", "LAST_READ_AT"]);
  const data = cacheGetOrSet("CHAT_READS_ALL_V1", 20, () => sh.getDataRange().getValues());

  const userName = _NOTIF_normText(user && user.nombre || "");
  const role = _NOTIF_normText(user && user.rol || "");
  const map = {};

  for(let i = 1; i < data.length; i++){
    const rowUser = _NOTIF_normText(data[i][0]);
    const rowRole = _NOTIF_normText(data[i][1]);

    if(rowUser !== userName || rowRole !== role){
      continue;
    }

    const tipo = _NOTIF_normText(data[i][2]);
    const itemId = String(data[i][3] || "").trim();
    const readAt = _NOTIF_toDate(data[i][4]);

    if(!tipo || !itemId){
      continue;
    }

    map[`${tipo}:${itemId}`] = readAt;
  }

  return map;
}

function _NOTIF_parseComments(raw){

  try{
    const arr = JSON.parse(String(raw || "[]"));
    return Array.isArray(arr) ? arr : [];
  }catch(error){
    return [];
  }
}

function _NOTIF_lastComment(comments){

  if(!Array.isArray(comments) || comments.length === 0){
    return null;
  }

  let best = null;

  comments.forEach(comment => {
    const when = _NOTIF_toDate(comment && comment.fecha);

    if(!when){
      return;
    }

    if(!best || when.getTime() > best.when.getTime()){
      best = {
        raw: comment,
        when: when
      };
    }
  });

  if(best){
    return best;
  }

  const fallback = comments[comments.length - 1];

  return {
    raw: fallback,
    when: new Date(0)
  };
}

function _NOTIF_taskTitle(item){

  const text = _NOTIF_getAny(item, ["DESCRIPCION_CORTA", "DESCRIPCION", "TITULO"]);
  return String(text || "").trim();
}

function _NOTIF_collectUnread(items, tipo, readMap, user){

  const result = [];
  const userName = _NOTIF_normText(user && user.nombre || "");

  (items || []).forEach(item => {
    const itemId = String(_NOTIF_getAny(item, ["ID_OBSERVACION", "ID"]) || "").trim();

    if(!itemId){
      return;
    }

    const comments = _NOTIF_parseComments(_NOTIF_getAny(item, ["COMENTARIO_JSON", "DEV_COMENTARIO"]));

    if(comments.length === 0){
      return;
    }

    const last = _NOTIF_lastComment(comments);

    if(!last){
      return;
    }

    const lastBy = String(last.raw && last.raw.usuario || "").trim();

    if(_NOTIF_normText(lastBy) === userName){
      return;
    }

    const key = `${_NOTIF_normText(tipo)}:${itemId}`;
    const readAt = readMap[key];

    if(readAt && last.when.getTime() <= readAt.getTime()){
      return;
    }

    result.push({
      tipo: tipo,
      id: itemId,
      titulo: _NOTIF_taskTitle(item),
      estado: String(_NOTIF_getAny(item, ["ESTADO"]) || "").trim(),
      lastBy: lastBy,
      lastAt: Utilities.formatDate(last.when, Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm"),
      lastMessage: String(last.raw && last.raw.mensaje || "").trim()
    });
  });

  return result;
}

function _NOTIF_toList(value){

  if(Array.isArray(value)){
    return value.map(x => String(x || "").trim()).filter(Boolean);
  }

  const text = String(value || "").trim();

  if(!text){
    return [];
  }

  if((text.startsWith("[") && text.endsWith("]")) || (text.startsWith("{") && text.endsWith("}"))){
    try{
      const parsed = JSON.parse(text);

      if(Array.isArray(parsed)){
        return parsed.map(x => String(x || "").trim()).filter(Boolean);
      }

      if(parsed && typeof parsed === "object"){
        return Object.values(parsed)
          .reduce((acc, arr) => acc.concat(Array.isArray(arr) ? arr : []), [])
          .map(x => String(x || "").trim())
          .filter(Boolean);
      }
    }catch(error){
      // fallback below
    }
  }

  return text.split(/[;,]/).map(x => String(x || "").trim()).filter(Boolean);
}

function _NOTIF_getActiveBroadcasts(user){

  // Comparte el mismo cache de lectura que _BROAD_readAll (37_BroadcastService.gs)
  // en vez de leer la hoja BROADCAST por su cuenta.
  const data = _BROAD_getRawGrid();

  if(data.length <= 1){
    return [];
  }

  const headers = data[0] || [];
  const role = _NOTIF_normText(user && user.rol || "");
  const userName = _NOTIF_normText(user && user.nombre || "");
  const result = [];

  for(let i = 1; i < data.length; i++){
    const row = mapRow(headers, data[i]);

    if(_NOTIF_bool(_NOTIF_getAny(row, ["ELIMINADO"]))){
      continue;
    }

    if(!_NOTIF_bool(_NOTIF_getAny(row, ["ACTIVO"]))){
      continue;
    }

    const estado = _NOTIF_normText(_NOTIF_getAny(row, ["ESTADO"]));

    if(estado && estado !== "PUBLICADO"){
      continue;
    }

    const alcance = _NOTIF_normText(_NOTIF_getAny(row, ["ALCANCE"])) || "TODOS";

    let visible = false;
    let scopeLabel = "General";

    if(alcance === "USUARIOS"){
      const usuarios = _NOTIF_toList(_NOTIF_getAny(row, ["USUARIOS"])).map(_NOTIF_normText);
      visible = usuarios.includes(userName);
      scopeLabel = "Usuarios";
    }else if(alcance === "ROL" || alcance === "ROLES"){
      const roles = _NOTIF_toList(_NOTIF_getAny(row, ["ROLES"]))
        .map(_NOTIF_normText);

      visible = roles.includes("TODOS") || roles.includes(role);
      scopeLabel = "Rol";
    }else{
      visible = true;
      scopeLabel = "General";
    }

    if(!visible){
      continue;
    }

    result.push({
      id: String(_NOTIF_getAny(row, ["ID"]) || "").trim(),
      titulo: String(_NOTIF_getAny(row, ["TITULO"]) || "").trim(),
      mensaje: String(_NOTIF_getAny(row, ["MENSAJE"]) || "").trim(),
      nivel: String(_NOTIF_getAny(row, ["NIVEL"]) || "INFO").trim(),
      scopeLabel: scopeLabel,
      fechaPublicacion: String(_NOTIF_getAny(row, ["FECHA_PUBLICACION"]) || "").trim()
    });
  }

  return result;
}

function _NOTIF_getBroadcastAckSet(user){

  const sh = _NOTIF_ensureBroadcastAckSheet();
  const data = cacheGetOrSet("BROADCAST_ACKS_ALL_V1", 20, () => sh.getDataRange().getValues());
  const userName = _NOTIF_normText(user && user.nombre || "");
  const set = {};

  for(let i = 1; i < data.length; i++){
    const rowUser = _NOTIF_normText(data[i][0]);

    if(rowUser !== userName){
      continue;
    }

    const broadcastId = String(data[i][2] || "").trim();
    const publicationAt = String(data[i][3] || "").trim();
    const ackAt = String(data[i][4] || "").trim();

    if(!broadcastId){
      continue;
    }

    set[_NOTIF_normText(broadcastId)] = {
      publicationAt: publicationAt,
      ackAt: ackAt
    };
  }

  return set;
}

function _NOTIF_broadcastAnnouncements(user){

  const active = _NOTIF_getActiveBroadcasts(user);
  const ackSet = _NOTIF_getBroadcastAckSet(user);

  return active.filter(item => {
    const ack = ackSet[_NOTIF_normText(item.id)];

    if(!ack){
      return true;
    }

    // If there is an ACK timestamp, do not show again unless there is evidence of a newer publication.
    if(!String(ack.ackAt || "").trim()){
      return true;
    }

    const publicationText = String(item.fechaPublicacion || "").trim();
    const ackPublicationText = String(ack.publicationAt || "").trim();

    if(publicationText && ackPublicationText && publicationText === ackPublicationText){
      return false;
    }

    const pub = _NOTIF_toDate(item.fechaPublicacion);
    const ackPub = _NOTIF_toDate(ack.publicationAt);

    if(pub && ackPub){
      return pub.getTime() > ackPub.getTime();
    }

    // Safe fallback: if we cannot compare publication moments and ACK exists, suppress to avoid duplicate alerts.
    return false;
  });
}

function NOTIF_getNavbarSummary(userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  const readMap = _NOTIF_readMapForUser(user);

  const tasks = TASK_SERVICE_list(user, {view: "ACTIVAS"}) || [];
  const observations = OBS_SERVICE_list(user, {view: "ACTIVAS"}) || [];
  const broadcastAlerts = _NOTIF_broadcastAnnouncements(user);

  const unread = []
    .concat(_NOTIF_collectUnread(tasks, "TAREA", readMap, user))
    .concat(_NOTIF_collectUnread(observations, "OBSERVACION", readMap, user))
    .sort((a, b) => String(b.lastAt).localeCompare(String(a.lastAt)));

  return {
    unreadChatsCount: unread.length,
    unreadChats: unread.slice(0, 20),
    activeBroadcasts: _NOTIF_getActiveBroadcasts(user).slice(0, 20),
    broadcastAlerts: broadcastAlerts.slice(0, 20)
  };
}

function NOTIF_ackBroadcasts(payload, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  const ids = Array.isArray(payload && payload.ids)
    ? payload.ids
    : [payload && payload.id];

  const cleaned = ids.map(id => String(id || "").trim()).filter(Boolean);

  if(cleaned.length === 0){
    return {ok: true};
  }

  const sh = _NOTIF_ensureBroadcastAckSheet();
  const data = sh.getDataRange().getValues();
  const userName = String(user.nombre || "").trim();
  const role = String(user.rol || "").trim();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");
  const publications = {};

  const broadcastSource = _BROAD_readAll();
  const broadcastData = broadcastSource.data;
  const broadcastCol = broadcastSource.col;

  cleaned.forEach(id => {
    const rowIndex = _BROAD_findRowIndexById(broadcastData, broadcastCol, id);

    if(rowIndex >= 0){
      publications[id] = String(_BROAD_cell(broadcastData[rowIndex], broadcastCol, "FECHA_PUBLICACION") || "").trim();
    }
  });

  const existing = new Set();

  for(let i = 1; i < data.length; i++){
    const rowUser = _NOTIF_normText(data[i][0]);
    const rowRole = _NOTIF_normText(data[i][1]);

    if(rowUser !== _NOTIF_normText(userName) || rowRole !== _NOTIF_normText(role)){
      continue;
    }

    const broadcastId = String(data[i][2] || "").trim();
    const publicationAt = String(data[i][3] || "").trim();

    existing.add(`${_NOTIF_normText(broadcastId)}|${publicationAt}`);
  }

  cleaned.forEach(id => {
    const publicationAt = String(publications[id] || "").trim();
    const ackKey = `${_NOTIF_normText(id)}|${publicationAt}`;

    if(existing.has(ackKey)){
      return;
    }

    sh.appendRow([userName, role, id, publicationAt, now]);
    existing.add(ackKey);
  });

  cacheClear("BROADCAST_ACKS_ALL_V1");

  return {ok: true};
}

function NOTIF_markChatRead(payload, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  const tipo = _NOTIF_normText(payload && payload.tipo || "");
  const itemId = String(payload && payload.id || "").trim();

  if(!tipo || !itemId){
    throw new Error("VALIDATION_ERROR: tipo e id requeridos");
  }

  const sh = _NOTIF_ensureSheet("CHAT_READS", ["USER", "ROL", "TIPO", "ITEM_ID", "LAST_READ_AT"]);
  const data = sh.getDataRange().getValues();

  const userName = String(user.nombre || "").trim();
  const userRole = String(user.rol || "").trim();
  const now = Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "yyyy-MM-dd HH:mm:ss");

  for(let i = 1; i < data.length; i++){
    const rowUser = String(data[i][0] || "").trim();
    const rowRole = String(data[i][1] || "").trim();
    const rowTipo = _NOTIF_normText(data[i][2]);
    const rowId = String(data[i][3] || "").trim();

    if(_NOTIF_normText(rowUser) !== _NOTIF_normText(userName)){
      continue;
    }

    if(_NOTIF_normText(rowRole) !== _NOTIF_normText(userRole)){
      continue;
    }

    if(rowTipo === tipo && rowId === itemId){
      sh.getRange(i + 1, 5).setValue(now);
      cacheClear("CHAT_READS_ALL_V1");
      return {ok: true};
    }
  }

  sh.appendRow([userName, userRole, tipo, itemId, now]);

  cacheClear("CHAT_READS_ALL_V1");

  return {ok: true};
}
