function WORKFLOW_isTransitionAllowed(user, tipo, newState){

  if(PERM_isManager(user)){
    return true;
  }

  if(PERM_isRole(user, "DEV")){
    return newState === "EN_PROCESO" || newState === "EN_REVISION";
  }

  if(PERM_isRole(user, "QA")){
    return newState === "ABIERTO" || newState === "EN_REVISION" || newState === "CERRADO";
  }

  return false;
}
