# Especificacion de desarrollo

Fecha: 2026-07-11

## 1) Diccionario de datos

### CFG_APP
- CLAVE: Texto, obligatorio, unico.
- VALOR: Texto, obligatorio.
- DESCRIPCION: Texto, opcional.

### USERS
- USER_ID: Texto, PK.
- NOMBRE: Texto, obligatorio.
- EMAIL: Texto, obligatorio, formato correo.
- ROL_PRINCIPAL: Texto, obligatorio. Valores: PM, LIDER_DEV, LIDER_QA, DEV, QA.
- ROLES_PERMITIDOS: Texto, obligatorio, separado por ;.
- ACTIVO: Booleano.
- CREATED_AT: FechaHora auto.
- UPDATED_AT: FechaHora auto.

### WORK_ITEMS
- ITEM_ID: Texto, PK.
- TIPO: TAREA u OBSERVACION.
- TITULO_CORTO: Texto.
- DESCRIPCION: Texto largo.
- PASOS_REPRODUCIR: Texto largo opcional.
- MODULO: Texto.
- PRIORIDAD: BAJA, MEDIA, ALTA, CRITICA.
- ESTADO: ABIERTO, EN_PROCESO, EN_REVISION, BLOQUEADO, CERRADO.
- QA_OWNER_ID: FK USERS.
- DEV_OWNER_ID: FK USERS opcional.
- SPRINT_ID: FK SPRINTS opcional.
- STORY_POINTS: Numero entero positivo opcional.
- FECHA_CREACION: FechaHora auto.
- FECHA_ASIGNACION: FechaHora opcional.
- FECHA_CIERRE: FechaHora opcional.
- ELIMINADO: Booleano.
- CAPTURA_URL: URL Drive opcional.
- TAGS_JSON: JSON array valido opcional.
- CREATED_BY: USER_ID.
- UPDATED_BY: USER_ID.
- UPDATED_AT: FechaHora auto.

### WORK_ASSIGNMENTS
- ASSIGN_ID: Texto PK.
- ITEM_ID: FK WORK_ITEMS.
- USER_ID: FK USERS.
- ROL_EN_ITEM: PM, LIDER_DEV, LIDER_QA, DEV, QA.
- ACTIVO: Booleano.
- CREATED_AT: FechaHora auto.
- CLOSED_AT: FechaHora opcional.

### WORK_COMMENTS
- COMMENT_ID: Texto PK.
- ITEM_ID: FK WORK_ITEMS.
- USER_ID: FK USERS.
- ROL: rol activo.
- MENSAJE: Texto largo.
- ACCION: COMENTARIO, EN_REVISION, APROBADO, RECHAZADO.
- ESTADO_RESULTANTE: opcional.
- CREATED_AT: FechaHora auto.

### WORK_STATUS_HISTORY
- HIST_ID: Texto PK.
- ITEM_ID: FK WORK_ITEMS.
- FROM_STATUS: Texto.
- TO_STATUS: Texto.
- CHANGED_BY: USER_ID.
- CHANGED_AT: FechaHora auto.
- MOTIVO: opcional.

### USER_PREFERENCES
- PREF_ID: Texto PK.
- USER_ID: FK USERS.
- CONTEXTO: DASHBOARD_TASK, DASHBOARD_OBS, KANBAN.
- PREF_KEY: COLUMN_VISIBILITY, FILTERS, SORT, PAGE_SIZE.
- PREF_JSON: JSON valido.
- UPDATED_AT: FechaHora auto.

### SPRINTS
- SPRINT_ID: Texto PK.
- NOMBRE: Texto.
- FECHA_INICIO: Fecha obligatoria.
- FECHA_FIN: Fecha mayor a inicio.
- OBJETIVO: opcional.
- ESTADO: PLANIFICADO, ACTIVO, CERRADO.
- CREATED_AT: FechaHora auto.

### DOC_CATALOG
- DOC_ID: Texto PK.
- TITULO: Texto.
- CATEGORIA: Texto.
- URL: URL valida.
- ICONO: clase icono o emoji opcional.
- ORDEN: numero entero ascendente.
- ACTIVO: Booleano.
- DESCRIPCION: opcional.
- UPDATED_AT: FechaHora auto.

### AUDIT_LOG
- LOG_ID: Texto PK.
- USER_ID: USER_ID.
- ACCION: accion funcional.
- ENTIDAD: tabla afectada.
- ENTITY_ID: id afectado.
- PAYLOAD_JSON: snapshot opcional.
- CREATED_AT: FechaHora auto.

## 2) Matriz de permisos

Regla global:
- PM ve y opera todo.
- LIDER_DEV y LIDER_QA organizan DEV y QA.
- DEV y QA operan solo su flujo directo.

Resumen:
- PM, LIDER_DEV y LIDER_QA: acceso total operativo (listar, asignar, reasignar, mover estados, exportar, administrar documentos).
- DEV: trabaja solo items asignados; comenta y mueve estados permitidos para desarrollo.
- QA: trabaja sobre items creados/asignados; aprueba o rechaza en revision; comenta y exporta lo visible.

Validaciones backend:
- No confiar en rol del frontend.
- Resolver rol activo desde sesion de servidor.
- En conflicto de permisos, prevalece denegar.

## 3) Catalogo de endpoints

Respuesta estandar:
- ok: boolean
- data: objeto o arreglo
- error: { code, message, details? }

Codigos estandar:
- AUTH_REQUIRED
- AUTH_FORBIDDEN
- VALIDATION_ERROR
- NOT_FOUND
- CONFLICT
- INTERNAL_ERROR

### Auth
- login(nombre)
- logout()
- getSessionInfo()

### Work items
- listWorkItems
- getWorkItemById
- createWorkItem
- updateWorkItem
- deleteWorkItem
- restoreWorkItem

### Assignments
- setAssignments

### Comments/decisiones
- listComments
- addComment
- addDecision

### Kanban
- getKanbanBoard
- moveKanbanItem

### Preferencias
- getUserPreferences
- saveUserPreferences

### Documentacion
- listDocuments

### Reportes
- previewExportReport
- exportReportPdf

## 4) Modal de exportacion de informe

Objetivo:
- Un modal unico para exportar informe de desarrollo.
- Cobertura de chats, tareas asignadas y observaciones.
- Reuso de preview/PDF existentes.

Payload de exportacion:
- fromDate
- toDate
- scope
- selectedItemIds
- includeTasks
- includeObservations
- includeChats
- includeImages
- includeStatusHistory
- groupBy
- detailLevel

Criterios de aceptacion:
- Permite vista previa sin descargar.
- Descarga PDF con nombre Informe_Desarrollo_YYYYMMDD_HHMM.
- Respeta filtros y alcance.
- Incluye secciones segun checkboxes.
- Mantiene flujo actual de dashboard.
- Muestra errores claros.

## 5) Bitacora de cambios - 2026-08-10

Trabajo realizado en sesion de refactor/performance. Sirve como referencia para continuar en otra maquina sin perder el hilo.

### Fase 1 - Cimientos: permisos y workflow centralizados

- `04_Permisions.gs`: nuevo. `PERM_isManager(user)` (PM/LIDER_QA/LIDER_DEV) y `PERM_isRole(user, rol)`, reemplazan la logica de roles duplicada que existia en `31_TareasService.gs`, `32_ObservacionesService.gs` y `34_KanbanService.gs`.
- `30_Workflows.gs`: nuevo. `WORKFLOW_isTransitionAllowed(user, tipo, newState)` centraliza las reglas de que estado puede mover cada rol (antes hardcodeado dentro de `KANBAN_moveItem`).
- Modulos que quedan como esqueleto a proposito (sin logica que mover hoy): `06_Colors.gs`, `23_DBHistorial.gs`, `24_DB_Informes.gs`, `25_DB_InformesDetalle.gs`. Cada uno tiene un comentario explicando por que.

### Fase 2 - Fix bug intermitente "Mis Observaciones"

- `script_main.html`, funcion `getClientUserContext()`: antes releia `nombre`/`rol` desde el DOM (`#txtNavbarUser`, `#cmbRoles`) en cada llamada, lo cual era fragil ante timing de render. Ahora usa `window.currentUser`/`window.currentRole` (ya mantenidos sincronizados en el login y en el `onchange` de `#cmbRoles`), con el DOM solo como fallback.

### Fase 3 - Kanban unificado (Tareas + Observaciones)

- `34_KanbanService.gs`: `KANBAN_loadData` ahora fusiona `TASK_SERVICE_list` + `OBS_SERVICE_list`, cada item lleva `tipo: "TAREA"|"OBSERVACION"`. `KANBAN_moveItem` detecta el tipo del item y llama a `DB_TASK_updateEstado` o `DB_OBS_updateEstado` segun corresponda, validando con `WORKFLOW_isTransitionAllowed`.
- `22_DBObservaciones.gs`: nueva `DB_OBS_updateEstado(id, newEstado)`, mismo patron que `DB_TASK_updateEstado`.
- `script_kanban.html`: tarjetas con borde de color segun tipo (`.kanban-card-obs`, variable `--studio-obs` en `style_colors.html`). Nueva `abrirDetalleObservacionKanban(id)` (reusa el modal `#modalObservacion` que ya existia sin usar). El boton de chat ahora fija `#chatTipo`/`#DshCategoria` segun el tipo real del item en vez de siempre "TAREA".

### Fase 4 - Columnas redimensionables + fixes menores de UI

- `index.html`: se agrego `<script>` del plugin `colResizable` (CDN `jsdelivr`, verificado que responde JS real antes de usarlo).
- `script_dashboard.html`: se quitaron los anchos fijos por columna del DataTable (`columnDefs`/`width`) salvo la columna del checkbox (`width:"34px"`, excluida del resize via `disabledColumns:[0]`). `colResizable` se inicializa despues de crear la tabla.
- `05_Drive.gs`, `DOCS_getHistoryMap`: se quito el `.slice(-8)` que truncaba el historial de aperturas de documentos; ahora el badge "Abierto (N)" y el modal de historial muestran el total real y la lista completa (el truncado era 100% server-side, no hubo que tocar el frontend).

### Fase 5 - Rendimiento de lectura/escritura en Sheets

Motivacion: casi todo el backend leia la hoja entera (`getDataRange().getValues()`) y recorria con `for` en JS, tanto para listar como para encontrar 1 fila por ID. Se ataco de forma completa en 11 archivos (el catalogo inicial de 10 quedo corto por una limitacion de busqueda propia; se detecto y corrigio en la misma sesion, agregando `31_TareasService.gs` y completando `32_ObservacionesService.gs`).

Dos utilidades nuevas en `02_Utils.gs`:
- `cacheGetOrSet(key, ttlSeconds, buildFn)`: envuelve el patron cache-hit/cache-miss que ya usaban `DB_TASK_getAll`/`DB_OBS_getAll` (20s TTL).
- `findRowIndexById(sheet, idColIndex1Based, id)`: busca 1 fila por ID con `TextFinder` (`matchEntireCell(true)`) en vez de traer toda la hoja y recorrerla en JS. Devuelve el numero de fila absoluto (1-based) o `-1`.

Claves de cache en uso (todas 20s TTL, `CacheService.getScriptCache()`):

| Clave | Hoja | Se invalida en |
|---|---|---|
| `TASK_ALL_V1` | TAREAS | `_TASK_CACHE_clear()` (ya existia) |
| `OBS_ALL_V1` | OBSERVACIONES | `_OBS_CACHE_clear()` (ya existia) |
| `USERS_ALL_V1` + `USUARIOS_RAW_GRID_V1` | USUARIOS | `_USERS_CACHE_clear()` (20_DBUsers.gs) |
| `BACKLOG_ALL_V1` | BACKLOG | `cacheClear` en `BACKLOG_save`/`BACKLOG_delete` |
| `CALENDARIO_ALL_V1` | CALENDARIO | `cacheClear` en `CALENDAR_save`/`CALENDAR_delete` |
| `BROADCAST_ALL_V1` | BROADCAST | `cacheClear` en `BROADCAST_save`/`publish`/`delete`/`restore` |
| `BROADCAST_ACKS_ALL_V1` | BROADCAST_ACKS | `cacheClear` en `NOTIF_ackBroadcasts` |
| `CHAT_READS_ALL_V1` | CHAT_READS | `cacheClear` en `NOTIF_markChatRead` |
| `DAILY_ALL_V1` | DAILY_SCRUM | `cacheClear` en `DAILY_save` |
| `DOC_HISTORIAL_ALL_V1` | DOC_HISTORIAL | `cacheClear` en `DOCS_registerOpen` |

Archivos con logica de busqueda-por-ID migrada a `findRowIndexById`: `20_DBUsers.gs`, `10_Users.gs`, `21_DBTareas.gs`, `22_DBObservaciones.gs`, `31_TareasService.gs`, `32_ObservacionesService.gs`, `35_AdminService.gs`, `37_BroadcastService.gs`.

**Excepcion deliberada** (no se tocaron, se dejaron con lectura fresca de la hoja): `DAILY_save` (busca por usuario+fecha, no por 1 columna), `NOTIF_ackBroadcasts` y `NOTIF_markChatRead` (buscan por usuario+rol). `TextFinder` solo busca 1 columna, no soporta match compuesto, y usar la cache ahi arriesgaba sobrescribir cambios concurrentes de otro usuario al reescribir la fila completa. Estas 3 si invalidan su cache despues de escribir.

### Pendiente (no se hizo en esta sesion)

- **`LockService`**: sigue sin implementarse. Sigue siendo el principal riesgo de condicion de carrera con escrituras concurrentes en Sheets con multiples usuarios simultaneos.
- **Error "pagina no disponible" para algunos usuarios**: diagnosticado como problema de configuracion de despliegue (acceso/cuenta de Google), no de codigo. Queda pendiente revisar "Implementar > Administrar implementaciones" (acceso y "ejecutar como") y descartar el caso de multiples cuentas Google activas en el navegador del usuario afectado.
