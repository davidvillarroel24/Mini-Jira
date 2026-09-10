function getDashboardData(userFromClient) {

  const user = getSession(userFromClient);  

  if (!user) {
    throw new Error("Sesión inválida");
  }

  return {
    user: user
  };
}

function DASHBOARD_loadData(filters = {}, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("Sesion invalida");
  }

  const categoria =
    filters.categoria ||
    "OBSERVACION";

  const data =
    categoria === "TAREA"
      ? TASK_SERVICE_list(user, filters)
      : OBS_SERVICE_list(user, filters);

  const sheetName =
    categoria === "TAREA"
      ? "TAREAS"
      : "OBSERVACIONES";

  const sh = SpreadsheetApp
    .getActive()
    .getSheetByName(sheetName);

  const columnOrder = sh
    ? sh
      .getRange(1, 1, 1, sh.getLastColumn())
      .getValues()[0]
      .map(h => String(h).trim())
      .filter(Boolean)
    : [];

  const contexto =
    `${categoria}_${filters.view || "ACTIVAS"}`;

  let preferencias = {
    visibleColumns: [],
    estado: "TODOS",
    prioridad: "TODOS",
    sprint: "TODOS"
  };

  try{
    preferencias = USER_PREF_getDashboard(contexto, userFromClient) || preferencias;
  }catch(error){
    Logger.log(error);
  }

  return {
    data: data,
    preferencias: preferencias,
    columnOrder: columnOrder
  };
}