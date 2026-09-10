function login(nombre) {

  const user = DB_USERS_findByName(nombre);

  if (!user) {
    return {
      ok: false,
      message: "Usuario no encontrado"
    };
  }

  if (user.Activo !== true && user.Activo !== "TRUE") {
    return {
      ok: false,
      message: "Usuario inactivo"
    };
  }

  const session = {

    id: user.ID,

    nombre: user.Nombre,

    rol: user.Rol_principal,

    roles: user.Roles_permitidos
  };

  return {
    ok: true,
    user: session
  };
}

function _normalizeClientUser(user) {

  if(!user || typeof user !== "object"){
    return null;
  }

  const nombre = String(user.nombre || "").trim();
  const rol = String(user.rol || "").trim();
  const roles = String(user.roles || rol || "").trim();

  if(!nombre || !rol){
    return null;
  }

  return {
    nombre: nombre,
    rol: rol,
    roles: roles
  };
}

function getSession(userFromClient) {

  const clientUser = _normalizeClientUser(userFromClient);

  if(clientUser){
    return clientUser;
  }

  // Compatibilidad de lectura temporal con sesiones antiguas.
  const raw = PropertiesService
    .getUserProperties()
    .getProperty("SESSION");

  if(!raw){
    return null;
  }

  try{
    const legacy = JSON.parse(raw);
    return _normalizeClientUser(legacy);
  }catch(error){
    Logger.log(error);
    return null;
  }
}

function logout() {

  return true;
}