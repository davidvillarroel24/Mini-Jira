function crearObservacion(payload, userFromClient) {

  const user = getSession(userFromClient || payload.user || payload.__user);

  if (!user) {
    throw new Error("No autenticado");
  }

  return OBS_SERVICE_create(payload, user);
}


function obtenerObservaciones(
  user,
  filters = {}
) {

  //const user =
  //  getSession();

  Logger.log(user);

  Logger.log(
    JSON.stringify(filters)
  );

  if(!user){

    throw new Error(
      "No autenticado"
    );
  }

  return OBS_SERVICE_list(
    user,
    filters
  );
}