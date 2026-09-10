function _KANBAN_columns(){
  return [
    {id: "ABIERTO", label: "Abierto"},
    {id: "EN_PROCESO", label: "En proceso"},
    {id: "EN_REVISION", label: "En revision"},
    {id: "BLOQUEADO", label: "Bloqueado"},
    {id: "CERRADO", label: "Cerrado"}
  ];
}

function _KANBAN_normalizeEstado(value){
  const allowed = _KANBAN_columns().map(c => c.id);
  const estado = String(value || "").trim().toUpperCase();
  return allowed.includes(estado) ? estado : "ABIERTO";
}

function _KANBAN_normalizeSprint(value){
  const sprint = String(value || "").trim();
  return sprint ? sprint : "BACKLOG";
}

function _KANBAN_resolveTaskSprint(task){
  if(!task || typeof task !== "object"){
    return "";
  }

  return task.SPRINT
    || task.SPRINT_ID
    || task.ID_SPRINT
    || task.VALIDACION_RELACIONADA
    || "";
}

function _KANBAN_userInList(list, userName){
  if(!Array.isArray(list)){
    return false;
  }

  const target = String(userName || "").trim().toUpperCase();

  if(!target){
    return false;
  }

  return list.some(name => String(name || "").trim().toUpperCase() === target);
}

function _KANBAN_isAssignedToUser(task, user){

  const userName = String(user && user.nombre || "").trim();

  if(!userName){
    return false;
  }

  if(String(task && task.DEV_ASIGNADO || "").trim().toUpperCase() === userName.toUpperCase()){
    return true;
  }

  try{
    const byRole = JSON.parse(String(task && task.USUARIOS_ASIGNADOS || "{}"));

    if(!byRole || typeof byRole !== "object"){
      return false;
    }

    return Object.keys(byRole).some(rol => _KANBAN_userInList(byRole[rol], userName));
  }catch(error){
    return false;
  }
}

function KANBAN_loadData(filters, userFromClient){

  const user = getSession(userFromClient);

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  const safeFilters = filters || {};
  const selectedSprint = String(safeFilters.sprintId || "TODOS").trim() || "TODOS";
  const selectedOnlyMine = !!safeFilters.onlyMine;

  const tasks = (TASK_SERVICE_list(user, {view: "ACTIVAS"}) || [])
    .map(task => Object.assign({__tipo: "TAREA"}, task));

  const observaciones = (OBS_SERVICE_list(user, {view: "ACTIVAS"}) || [])
    .map(obs => Object.assign({__tipo: "OBSERVACION"}, obs));

  const items = tasks.concat(observaciones);

  const filteredByOwner = selectedOnlyMine
    ? items.filter(item => _KANBAN_isAssignedToUser(item, user))
    : items;

  const mapped = filteredByOwner.map(item => ({
    id: String(item.ID_OBSERVACION || ""),
    tipo: item.__tipo,
    titulo: String(item.DESCRIPCION_CORTA || "").trim(),
    modulo: String(item.MODULO || "").trim(),
    prioridad: String(item.PRIORIDAD || "").trim(),
    estado: _KANBAN_normalizeEstado(item.ESTADO),
    qa: String(item.QA || "").trim(),
    dev: String(item.DEV_ASIGNADO || "").trim(),
    usuariosAsignados: String(item.USUARIOS_ASIGNADOS || "{}"),
    sprintId: _KANBAN_normalizeSprint(_KANBAN_resolveTaskSprint(item)),
    fecha: String(item.FECHA || "")
  }));

  const sprintSet = {};
  sprintSet.BACKLOG = true;

  mapped.forEach(item => {
    sprintSet[item.sprintId] = true;
  });

  const sprints = [{id: "TODOS", label: "Todos"}].concat(
    Object.keys(sprintSet)
      .sort()
      .map(id => ({id: id, label: id === "BACKLOG" ? "Backlog" : id}))
  );

  const filtered = selectedSprint === "TODOS"
    ? mapped
    : mapped.filter(item => item.sprintId === selectedSprint);

  const columns = _KANBAN_columns();
  const itemsByColumn = {};
  const counts = {};

  columns.forEach(col => {
    itemsByColumn[col.id] = [];
    counts[col.id] = 0;
  });

  filtered.forEach(item => {
    if(!itemsByColumn[item.estado]){
      itemsByColumn[item.estado] = [];
      counts[item.estado] = 0;
    }

    itemsByColumn[item.estado].push(item);
    counts[item.estado] += 1;
  });

  return {
    columns: columns,
    sprints: sprints,
    selectedSprint: selectedSprint,
    selectedOnlyMine: selectedOnlyMine,
    itemsByColumn: itemsByColumn,
    counts: counts
  };
}

function KANBAN_moveItem(payload, userFromClient){

  const user = getSession(userFromClient || (payload && payload.__user));

  if(!user){
    throw new Error("AUTH_REQUIRED");
  }

  const id = String(payload && payload.id || "").trim();
  const newState = _KANBAN_normalizeEstado(payload && payload.newState);

  if(!id){
    throw new Error("VALIDATION_ERROR: id requerido");
  }

  const visibleTasks = TASK_SERVICE_list(user, {view: "ACTIVAS"}) || [];
  const targetTask = visibleTasks.find(x => String(x.ID_OBSERVACION || "") === id);

  const tipo = targetTask ? "TAREA" : "OBSERVACION";

  if(!targetTask){
    const visibleObs = OBS_SERVICE_list(user, {view: "ACTIVAS"}) || [];
    const targetObs = visibleObs.find(x => String(x.ID_OBSERVACION || "") === id);

    if(!targetObs){
      throw new Error("AUTH_FORBIDDEN: item no disponible");
    }
  }

  if(!WORKFLOW_isTransitionAllowed(user, tipo, newState)){
    throw new Error("AUTH_FORBIDDEN: estado no permitido para el rol actual");
  }

  if(tipo === "TAREA"){
    DB_TASK_updateEstado(id, newState);
  }else{
    DB_OBS_updateEstado(id, newState);
  }

  return {
    ok: true,
    id: id,
    estado: newState
  };
}
