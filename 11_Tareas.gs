function obtenerTareas(
  user,
  filters = {}
){

  Logger.log(user);

  Logger.log(
    JSON.stringify(filters)
  );

  if(!user){

    throw new Error(
      "No autenticado"
    );
  }

  return TASK_SERVICE_list(
    user,
    filters
  );
}


function crearTarea(payload, userFromClient){

  const user =
    getSession(userFromClient || payload.user || payload.__user);

  if(!user){

    throw new Error(
      "No autenticado"
    );
  }

  return TASK_SERVICE_create(
    payload,
    user
  );
}
