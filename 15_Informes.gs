function _parseFechaInforme(value){

  if(!value){
    return null;
  }

  if(value instanceof Date){
    return value;
  }

  const raw = String(value).trim();

  if(!raw){
    return null;
  }

  const formatoLatam = raw.match(
    /^(\d{2})\/(\d{2})\/(\d{4})(?:\s+(\d{2}):(\d{2})(?::(\d{2}))?)?$/
  );

  if(formatoLatam){

    return new Date(
      Number(formatoLatam[3]),
      Number(formatoLatam[2]) - 1,
      Number(formatoLatam[1]),
      Number(formatoLatam[4] || 0),
      Number(formatoLatam[5] || 0),
      Number(formatoLatam[6] || 0)
    );
  }

  const iso = new Date(raw);

  if(!isNaN(iso.getTime())){
    return iso;
  }

  return null;
}

function _toBooleanInforme(value, defaultValue){

  if(value === undefined || value === null){
    return !!defaultValue;
  }

  if(typeof value === "boolean"){
    return value;
  }

  const parsed = String(value).trim().toUpperCase();

  return parsed === "TRUE" || parsed === "1" || parsed === "SI";
}

function _fileIdFromDriveUrl(url){

  if(!url){
    return "";
  }

  const match = String(url).match(/[-\w]{25,}/);
  return match ? match[0] : "";
}

function _formatFechaInforme(value){

  const dt = _parseFechaInforme(value);

  if(!dt){
    return "";
  }

  return Utilities.formatDate(
    dt,
    Session.getScriptTimeZone(),
    "dd/MM/yyyy HH:mm"
  );
}

function _normalizarPayloadInforme(payload){

  const base = payload || {};

  return {
    fromDate: base.fromDate || "",
    toDate: base.toDate || "",
    scope: base.scope || "SELECTED",
    selectedItemIds: Array.isArray(base.selectedItemIds) ? base.selectedItemIds : [],
    filteredItemIds: Array.isArray(base.filteredItemIds) ? base.filteredItemIds : [],
    includeTasks: _toBooleanInforme(base.includeTasks, true),
    includeObservations: _toBooleanInforme(base.includeObservations, true),
    includeDaily: _toBooleanInforme(base.includeDaily, false),
    includeChats: _toBooleanInforme(base.includeChats, true),
    includeImages: _toBooleanInforme(base.includeImages, true),
    includeStatusHistory: _toBooleanInforme(base.includeStatusHistory, true),
    groupBy: base.groupBy || "RESPONSABLE",
    detailLevel: base.detailLevel || "DETALLADO",
    requestedBy: base.requestedBy || ""
  };
}

function _buildCommentsInforme(item){

  let comments = [];

  try{
    comments = JSON.parse(item.COMENTARIO_JSON || "[]");
  }catch(error){
    Logger.log(error);
    comments = [];
  }

  return comments.map(comment => ({
    usuario: comment.usuario || "",
    rol: comment.rol || "",
    fecha: _formatFechaInforme(comment.fecha),
    mensaje: comment.mensaje || "",
    accion: comment.accion || "COMENTARIO"
  }));
}

function _buildStatusHistoryInforme(item, comments){

  const history = [];

  const add = (label, estado, fecha) => {

    if(!fecha){
      return;
    }

    history.push({
      label: label,
      estado: estado,
      fecha: _formatFechaInforme(fecha)
    });
  };

  add("Creado", "ABIERTO", item.FECHA);
  add("Asignado", "EN_PROCESO", item.FECHA_ASIGNACION);
  add("En revision", "EN_REVISION", item.FECHA_CORRECCION);
  add("Cerrado", "CERRADO", item.FECHA_CIERRE);

  comments.forEach(comment => {

    if(comment.accion === "EN_REVISION" || comment.accion === "APROBADO" || comment.accion === "RECHAZADO"){

      history.push({
        label: "Accion chat",
        estado: comment.accion,
        fecha: comment.fecha
      });
    }
  });

  return history;
}

function _buildImagenInforme(item, modo, includeImages){

  if(!includeImages){
    return {
      imageUrl: "",
      imageState: "SKIPPED",
      imageDebug: "includeImages=false"
    };
  }

  const fileId = _fileIdFromDriveUrl(item.CAPTURAS);

  if(!fileId){
    return {
      imageUrl: "",
      imageState: "MISSING",
      imageDebug: "No fileId parsed from CAPTURAS"
    };
  }

  if(modo === "preview"){
    return {
      imageUrl: `https://lh3.googleusercontent.com/d/${fileId}`,
      imageHtml: "",
      imageState: "OK",
      imageDebug: `preview fileId=${fileId}`
    };
  }

  try{
    // Enfoque directo: DriveApp -> base64 -> data URI.
    const file = DriveApp.getFileById(fileId);
    let blob = file.getBlob();
    let source = "driveapp";

    let mimeType = String(blob.getContentType() || "").toLowerCase();

    if(mimeType.indexOf("image/") !== 0){
      blob = blob.getAs(MimeType.JPEG);
      mimeType = "image/jpeg";
      source += "+coerced-jpeg";
    }

    if(mimeType === "image/webp"){
      try{
        blob = blob.getAs(MimeType.PNG);
        mimeType = "image/png";
        source += "+webp-to-png";
      }catch(conversionError){
        blob = blob.getAs(MimeType.JPEG);
        mimeType = "image/jpeg";
        source += "+webp-to-jpeg";
      }
    }

    const base64 = Utilities.base64Encode(blob.getBytes());
    const dataUri = `data:${mimeType};base64,${base64}`;
    const imageHtml = `<img src="${dataUri}" alt="Evidencia" />`;

    return {
      imageUrl: dataUri,
      imageHtml: imageHtml,
      imageState: "OK",
      imageDebug: `pdf fileId=${fileId} source=${source} mime=${mimeType} size=${blob.getBytes().length}`
    };

  }catch(error){

    Logger.log(`IMG_ERROR fileId=${fileId} modo=${modo} msg=${error && error.message ? error.message : error}`);

    return {
      imageUrl: "",
      imageHtml: "",
      imageState: "UNAVAILABLE",
      imageDebug: `ERROR fileId=${fileId} modo=${modo} msg=${error && error.message ? error.message : error}`
    };
  }
}

function _buildItemsInforme(payload, modo, user){

  let tasks = [];
  let observations = [];
  let dailies = [];

  if(payload.includeTasks){
    tasks = TASK_SERVICE_list(user, {}).map(item => {
      item.TIPO = "TAREA";
      return item;
    });
  }

  if(payload.includeObservations){
    observations = OBS_SERVICE_list(user, {}).map(item => {
      item.TIPO = item.TIPO || "OBSERVACION";
      return item;
    });
  }

  if(payload.includeDaily){
    dailies = (DAILY_loadData(user) || {}).data || [];
  }

  const merged = tasks.concat(observations).concat(dailies);

  const selectedIds = new Set(payload.selectedItemIds.map(String));
  const filteredIds = new Set(payload.filteredItemIds.map(String));

  const fromDate = payload.fromDate ? _parseFechaInforme(`${payload.fromDate} 00:00:00`) : null;
  const toDate = payload.toDate ? _parseFechaInforme(`${payload.toDate} 23:59:59`) : null;

  const byScope = merged.filter(item => {

    const itemId = String(item.ID_OBSERVACION || item.id || item.ID || "");

    if(payload.scope === "SELECTED"){
      return selectedIds.has(itemId);
    }

    if(payload.scope === "FILTERED"){
      return filteredIds.has(itemId);
    }

    if(payload.scope === "DATE_RANGE"){

      const createdAt = _parseFechaInforme(item.FECHA || item.fecha);

      if(!createdAt){
        return false;
      }

      if(fromDate && createdAt < fromDate){
        return false;
      }

      if(toDate && createdAt > toDate){
        return false;
      }

      return true;
    }

    return true;
  });

  return byScope
    .map(item => {

      const comments = payload.includeChats ? _buildCommentsInforme(item) : [];
      const imagen = _buildImagenInforme(item, modo, payload.includeImages);
      const history = payload.includeStatusHistory ? _buildStatusHistoryInforme(item, comments) : [];

      return {
        id: item.ID_OBSERVACION || item.id || item.ID || "",
        tipo: item.TIPO || "DAILY",
        modulo: item.MODULO || "DAILY_SCRUM",
        titulo: item.DESCRIPCION_CORTA || item.avanceHoy || "",
        descripcion: item.DESCRIPCION_CORTA || item.avanceHoy || "",
        pasos: item.PASOS_REPRODUCIR || item.proximoPaso || "",
        prioridad: item.PRIORIDAD || "",
        estado: item.ESTADO || (_toBooleanInforme(item.necesitoAyuda === "Si", false) ? "AYUDA_SOLICITADA" : "SIN_AYUDA"),
        qa: item.QA || "",
        dev: item.DEV_ASIGNADO || "",
        usuario: item.usuario || "",
        rol: item.rol || "",
        bloqueo: item.bloqueo || "",
        necesitoAyuda: item.necesitoAyuda || "",
        fecha: _formatFechaInforme(item.FECHA || item.fecha),
        actualizadoAt: _formatFechaInforme(item.ACTUALIZADO_AT || item.actualizadoAt || ""),
        fechaRaw: _parseFechaInforme(item.FECHA || item.fecha),
        imageUrl: item.TIPO === "DAILY" ? "" : imagen.imageUrl,
        imageHtml: item.TIPO === "DAILY" ? "" : imagen.imageHtml,
        imageState: item.TIPO === "DAILY" ? "SKIPPED" : imagen.imageState,
        imageDebug: item.TIPO === "DAILY" ? "daily without image" : imagen.imageDebug,
        comments: item.TIPO === "DAILY" ? [] : comments,
        history: item.TIPO === "DAILY" ? [] : history
      };
    })
    .sort((a, b) => {
      const aTime = a.fechaRaw ? a.fechaRaw.getTime() : 0;
      const bTime = b.fechaRaw ? b.fechaRaw.getTime() : 0;
      return bTime - aTime;
    });
}

function _groupItemsInforme(items, groupBy){

  const keyLabel = {
    RESPONSABLE: "Responsable",
    ESTADO: "Estado",
    PRIORIDAD: "Prioridad",
    TIPO: "Tipo"
  };

  const key = String(groupBy || "RESPONSABLE").toUpperCase();
  const map = {};

  items.forEach(item => {

    let groupValue = "Sin grupo";

    if(key === "ESTADO"){
      groupValue = item.estado || "Sin estado";
    }else if(key === "PRIORIDAD"){
      groupValue = item.prioridad || "Sin prioridad";
    }else if(key === "TIPO"){
      groupValue = item.tipo || "Sin tipo";
    }else{
      groupValue = item.dev || item.qa || item.usuario || "Sin responsable";
    }

    if(!map[groupValue]){
      map[groupValue] = [];
    }

    map[groupValue].push(item);
  });

  return {
    label: keyLabel[key] || keyLabel.RESPONSABLE,
    groups: Object.keys(map).sort().map(name => ({
      name: name,
      items: map[name]
    }))
  };
}

function _buildResumenInforme(items){

  return {
    totalTareas: items.filter(x => x.tipo === "TAREA").length,
    totalObservaciones: items.filter(x => x.tipo !== "TAREA").length,
    totalDaily: items.filter(x => x.tipo === "DAILY").length,
    totalCerradas: items.filter(x => x.estado === "CERRADO").length,
    totalEnRevision: items.filter(x => x.estado === "EN_REVISION").length,
    totalBloqueadas: items.filter(x => x.estado === "BLOQUEADO").length
  };
}

function _buildInformeDesarrollo(payload, modo, userFromClient){

  const user = getSession(userFromClient || (payload && payload.__user));

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  const normalized = _normalizarPayloadInforme(payload);

  if(!normalized.includeTasks && !normalized.includeObservations && !normalized.includeDaily && !normalized.includeChats){
    throw new Error("VALIDATION_ERROR: Debe seleccionar al menos un contenido");
  }

  if(normalized.scope === "DATE_RANGE" && (!normalized.fromDate || !normalized.toDate)){
    throw new Error("VALIDATION_ERROR: Debe completar desde y hasta");
  }

  const items = _buildItemsInforme(normalized, modo, user);
  const agrupado = _groupItemsInforme(items, normalized.groupBy);

  return {
    generatedAt: Utilities.formatDate(new Date(), Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm"),
    projectName: "Mini Jira",
    requestedBy: normalized.requestedBy || user.nombre || "",
    role: user.rol || "",
    filters: {
      scope: normalized.scope,
      fromDate: normalized.fromDate,
      toDate: normalized.toDate,
      includeTasks: normalized.includeTasks,
      includeObservations: normalized.includeObservations,
      includeDaily: normalized.includeDaily,
      includeChats: normalized.includeChats,
      includeImages: normalized.includeImages,
      includeStatusHistory: normalized.includeStatusHistory,
      detailLevel: normalized.detailLevel,
      groupBy: normalized.groupBy
    },
    resumen: _buildResumenInforme(items),
    imageDiagnostics: items.map(x => ({
      id: x.id,
      tipo: x.tipo,
      state: x.imageState,
      debug: x.imageDebug
    })),
    grouping: agrupado,
    items: items
  };
}

function generarPreviewInforme(
  idsOrPayload,
  modo = "preview"
){

  const payload = Array.isArray(idsOrPayload)
    ? {
      scope: "SELECTED",
      selectedItemIds: idsOrPayload,
      includeTasks: false,
      includeObservations: true,
      includeChats: false,
      includeImages: true,
      includeStatusHistory: false,
      groupBy: "TIPO",
      detailLevel: "DETALLADO"
    }
    : idsOrPayload;

  const reporte = _buildInformeDesarrollo(payload, modo, payload && payload.__user);

  const template = HtmlService.createTemplateFromFile("template_informe");

  template.reporte = reporte;
  template.fecha = reporte.generatedAt;
  template.observaciones = reporte.items;

  return template
    .evaluate()
    .getContent();
}