function DB_USERS_getAll(){

  return cacheGetOrSet("USERS_ALL_V1", 20, () => {

    const sh = SpreadsheetApp
      .getActive()
      .getSheetByName("USUARIOS");

    const data = sh.getDataRange().getValues();
    const headers = data[0];
    const result = [];

    for(let i = 1; i < data.length; i++){
      result.push(mapRow(headers, data[i]));
    }

    return result;
  });
}

function _USERS_CACHE_clear(){
  cacheClear("USERS_ALL_V1");
  cacheClear("USUARIOS_RAW_GRID_V1");
}

function DB_USERS_findByName(nombre) {

  const target = nombre.toString().trim().toUpperCase();

  const found = DB_USERS_getAll().find(obj =>
    obj.Nombre.toString().trim().toUpperCase() === target
  );

  return found || null;
}
