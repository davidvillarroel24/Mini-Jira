function mapRow(headers, row) {

  const obj = {};

  headers.forEach((h, i) => {
    obj[h] = row[i];
  });

  return obj;
}

function generateId(prefix) {

  return (
    prefix +
    "_" +
    new Date().getTime()
  );
}

function include(filename) {

  return HtmlService
    .createHtmlOutputFromFile(filename)
    .getContent();
}

function cacheGetOrSet(key, ttlSeconds, buildFn){

  const cache = CacheService.getScriptCache();
  const cached = cache.get(key);

  if(cached){
    try{
      return JSON.parse(cached);
    }catch(error){
      Logger.log(error);
    }
  }

  const result = buildFn();

  try{
    cache.put(key, JSON.stringify(result), ttlSeconds);
  }catch(error){
    Logger.log(error);
  }

  return result;
}

function cacheClear(key){
  CacheService.getScriptCache().remove(key);
}

function findRowIndexById(sheet, idColIndex1Based, id){

  const lastRow = sheet.getLastRow();

  if(lastRow < 2){
    return -1;
  }

  const range = sheet.getRange(2, idColIndex1Based, lastRow - 1, 1);

  const found = range
    .createTextFinder(String(id || "").trim())
    .matchEntireCell(true)
    .findNext();

  return found ? found.getRow() : -1;
}