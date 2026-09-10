function uploadImage(base64, fileName){

  if(!base64){
    return "";
  }

  const folder =
    DriveApp.getFolderById(
      CONFIG.DRIVE_OBSERVACIONES
    );

  const bytes = Utilities.base64Decode(
    base64.split(",")[1]
  );

  const blob = Utilities.newBlob(
    bytes,
    "image/jpeg",
    fileName + ".jpg"
  );

  const file = folder.createFile(blob);

  file.setSharing(
    DriveApp.Access.ANYONE_WITH_LINK,
    DriveApp.Permission.VIEW
  );

  return file.getUrl();
}

function DOCS_listFolderItems(){

  const historyMap = DOCS_getHistoryMap();

  const folder = DriveApp.getFolderById(
    CONFIG.DRIVEW_DOCS
  );

  const files = folder.getFiles();
  const result = [];

  while(files.hasNext()){

    const file = files.next();
    const mime = file.getMimeType();

    result.push({
      id: file.getId(),
      nombre: file.getName(),
      url: file.getUrl(),
      tipo: mime,
      updatedAt: Utilities.formatDate(
        file.getLastUpdated(),
        Session.getScriptTimeZone(),
        "dd/MM/yyyy HH:mm"
      ),
      historial: historyMap[file.getId()] || []
    });
  }

  result.sort((a, b) => {
    return String(a.nombre).localeCompare(String(b.nombre));
  });

  return result;
}

function _getDocHistorySheet(){

  const ss = SpreadsheetApp.openById(CONFIG.ID_SHEET);
  let sh = ss.getSheetByName("DOC_HISTORIAL");

  if(!sh){
    sh = ss.insertSheet("DOC_HISTORIAL");
    sh.appendRow([
      "DOC_ID",
      "DOC_NOMBRE",
      "USUARIO",
      "ROL",
      "ABIERTO_AT"
    ]);
  }

  return sh;
}

function DOCS_registerOpen(docId, docNombre, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("No autenticado");
  }

  const sh = _getDocHistorySheet();

  sh.appendRow([
    docId,
    docNombre || "",
    user.nombre || "",
    user.rol || "",
    new Date()
  ]);

  cacheClear("DOC_HISTORIAL_ALL_V1");

  return {
    ok: true
  };
}

function DOCS_getHistoryMap(){

  return cacheGetOrSet("DOC_HISTORIAL_ALL_V1", 20, () => _DOCS_buildHistoryMap());
}

function _DOCS_buildHistoryMap(){

  const sh = _getDocHistorySheet();
  const data = sh.getDataRange().getValues();
  const result = {};

  for(let i = 1; i < data.length; i++){

    const docId = String(data[i][0] || "").trim();

    if(!docId){
      continue;
    }

    if(!result[docId]){
      result[docId] = [];
    }

    const when = data[i][4] instanceof Date
      ? Utilities.formatDate(data[i][4], Session.getScriptTimeZone(), "dd/MM/yyyy HH:mm")
      : String(data[i][4] || "");

    result[docId].push({
      usuario: data[i][2] || "",
      rol: data[i][3] || "",
      fecha: when
    });
  }

  Object.keys(result).forEach(docId => {
    // Orden mas reciente primero. Se devuelve la lista completa (sin truncar)
    // para que el contador y el modal de historial muestren el total real.
    result[docId] = result[docId].reverse();
  });

  return result;
}