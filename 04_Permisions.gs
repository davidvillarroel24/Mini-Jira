function PERM_isManager(user){

  const roles = String((user && user.roles) || "").toUpperCase();

  return (
    roles.includes("PM") ||
    roles.includes("LIDER_QA") ||
    roles.includes("LIDER_DEV")
  );
}

function PERM_isRole(user, rol){

  return String((user && user.rol) || "") === String(rol || "");
}
