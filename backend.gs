/**
 * BACKEND SISTEMA ESCOLAR - Google Apps Script (versión corregida)
 *
 * CORRECCIONES APLICADAS:
 * 1. Se eliminó Utilities.Charset.UTF_8 de Utilities.base64EncodeWebSafe()
 *    (esa sobrecarga NO existe en Apps Script y lanzaba:
 *    "Los parámetros (number[],Utilities.Charset) no coinciden con la firma").
 * 2. Se mantiene un único doPost() con todas las acciones.
 */

const ss = SpreadsheetApp.getActiveSpreadsheet();

// ==========================================
// CACHE (Apps Script) - evita releer las 12 hojas en cada carga
// ==========================================
const CACHE_KEY = "get_data_v1";
const CACHE_TTL = 180;      // segundos
const SHEET_TTL = 300;      // cache por hoja

function cacheGet(key) {
  try { return CacheService.getScriptCache().get(key); } catch (e) { return null; }
}

function cachePut(key, value, ttl) {
  try {
    if (!value || value.length > 90000) return;
    CacheService.getScriptCache().put(key, value, ttl);
  } catch (e) { /* sin cache no es un error fatal */ }
}

function cacheDrop(key) {
  try { CacheService.getScriptCache().remove(key); } catch (e) {}
}

function getSheetDataCached(sheetName) {
  const key = "sheet_" + sheetName;
  const c = cacheGet(key);
  if (c) {
    try { return JSON.parse(c); } catch (e) { /* cache corrupta: se vuelve a leer */ }
  }
  const data = getSheetData(sheetName);
  cachePut(key, JSON.stringify(data), SHEET_TTL);
  return data;
}

function cacheDropSheet(sheetName) {
  cacheDrop("sheet_" + sheetName);
}

function cacheDropAll() {
  cacheDrop(CACHE_KEY);
  ["Estudiantes", "Grados", "Secciones", "Actividades", "Examenes", "Asistencia",
   "Notas", "Bimestres", "Materias", "Usuarios", "UsuariosPP", "PadresHijos", "AsignacionesDocentes", "Ciclos"]
    .forEach(function (n) { cacheDropSheet(n); });
}

// ==========================================
// GET - Obtener todos los datos
// ==========================================
function doGet(e) {
  const action = e.parameter.action;

  try {
    if (action === "get_data") {
      const auth = readToken(e.parameter.token);
      if (!auth || auth.t !== "auth") return authFail();

      const force = e.parameter.refresh === "1";
      if (force) cacheDropAll();

      const grados = getSheetDataCached("Grados");
      const secciones = getSheetDataCached("Secciones");

      const payload = {
        students: normalizeStudentRefs(getSheetDataCached("Estudiantes"), grados, secciones),
        grados: grados,
        secciones: secciones,
        activities: getSheetDataCached("Actividades"),
        exams: getSheetDataCached("Examenes"),
        attendance: getSheetDataCached("Asistencia"),
        grades: getSheetDataCached("Notas"),
        bimestres: getSheetDataCached("Bimestres"),
        materias: getSheetDataCached("Materias"),
        users: getSheetDataCached("Usuarios"),
        padres: getSheetDataCached("UsuariosPP"),
        padres_hijos: getSheetDataCached("PadresHijos"),
        ciclos: getSheetDataCached("Ciclos"),
        teacher_assignments: getSheetDataCached("AsignacionesDocentes")
      };

      return jsonText(JSON.stringify(scopePayload(payload, auth)));
    }
    return jsonResponse({ error: "Acción GET no reconocida: " + action });
  } catch (err) {
    return jsonResponse({ error: err.toString() });
  }
}

// ==========================================
// POST - Todas las acciones de escritura
// ==========================================
function doPost(e) {
  let data;
  try {
    data = JSON.parse(e.postData.contents);
  } catch (err) {
    return jsonResponse({ success: false, error: "JSON inválido: " + err.toString() });
  }

  const action = data.action;

  try {
    if (action === "login") return doLogin(data);
    if (action === "verify_student_code") return verifyStudentCode(data);
    if (action === "register_parent") return registerParent(data);

    const auth = readToken(data.token);
    if (!auth || auth.t !== "auth") return authFail();
    const role = String(auth.r || "").toLowerCase();

    const ADMIN_ONLY = [
      "save_student", "save_grado", "save_seccion", "save_materia",
      "save_bimestre", "save_user", "save_parent_user", "save_assignments",
      "delete_student", "delete_grados", "delete_secciones", "delete_materias",
      "delete_activities", "migrate_student_labels", "save_cycle", "delete_cycle"
    ];
    const STAFF = ["save_grades", "save_exams", "save_attendance", "save_actividad"];

    if (ADMIN_ONLY.indexOf(action) > -1 && role !== "admin") {
      return jsonResponse({ success: false, code: "AUTH", error: "No autorizado para esta acción" });
    }
    if (STAFF.indexOf(action) > -1 && role !== "admin" && role !== "docente") {
      return jsonResponse({ success: false, code: "AUTH", error: "No autorizado para esta acción" });
    }

    // Asegura la columna cicloId en las hojas que se segregan por ciclo escolar.
    const CYCLE_SHEETS = {
      save_student: "Estudiantes",
      save_actividad: "Actividades",
      save_grades: "Notas",
      save_exams: "Examenes",
      save_attendance: "Asistencia",
      save_assignments: "AsignacionesDocentes"
    };
    if (CYCLE_SHEETS[action]) ensureCycleColumns(CYCLE_SHEETS[action]);

    let out;
    switch (action) {
      case "save_student":
        out = saveRow("Estudiantes", data.payload, "id");
        break;
      case "save_grado":
        out = saveRow("Grados", data.payload, "id");
        break;
      case "save_seccion":
        out = saveRow("Secciones", data.payload, "id");
        break;
      case "save_materia":
        out = saveRow("Materias", data.payload, "id");
        break;
      case "save_actividad":
        out = saveRow("Actividades", data.payload, "id");
        break;
      case "save_bimestre":
        out = saveRow("Bimestres", data.payload, "id");
        break;
      case "save_user":
        out = saveRow("Usuarios", data.payload, "id");
        break;
      case "save_parent_user":
        ensureSheetColumns("UsuariosPP", ["childrenIds"]);
        out = saveRow("UsuariosPP", data.payload, "id");
        break;
      case "save_assignments":
        out = saveAssignments(data.payload);
        break;

      case "save_grades":
        out = batchUpdate("Notas", data.payload, ["studentId", "activityId"]);
        break;
      case "save_exams":
        out = batchUpdate("Examenes", data.payload, ["studentId", "materiaId", "bimestreId"]);
        break;
      case "save_attendance":
        out = saveAttendance(data.payload);
        break;
      case "save_cycle":
        out = saveCycle(data.payload);
        break;
      case "delete_cycle":
        out = deleteCycle(data.payload);
        break;

      case "delete_student":
        out = deleteRow("Estudiantes", data.id, "id");
        break;
      case "delete_grados":
        out = deleteRow("Grados", data.payload.id, "id");
        break;
      case "delete_secciones":
        out = deleteRow("Secciones", data.payload.id, "id");
        break;
      case "delete_materias":
        out = deleteRow("Materias", data.payload.id, "id");
        break;
      case "delete_activities":
        out = deleteRow("Actividades", data.payload.id, "id");
        break;

      case "migrate_student_labels":
        out = migrateStudentLabels();
        break;

      default:
        return jsonResponse({ success: false, message: "Acción no reconocida: " + action });
    }

    cacheDrop(CACHE_KEY);
    return out;
  } catch (err) {
    cacheDrop(CACHE_KEY);
    return jsonResponse({ success: false, error: err.toString() });
  }
}

// ==========================================
// LOGIN
// ==========================================
function doLogin(data) {
  const username = String((data && data.username) || "").trim().toLowerCase();
  const password = String((data && data.password) || "");
  if (!username || !password) {
    return jsonResponse({ success: false, message: "Usuario y contraseña son obligatorios" });
  }

  const users = getSheetDataCached("Usuarios") || [];
  const padres = getSheetDataCached("UsuariosPP") || [];

  function buscar(lista, rolForzado) {
    for (var i = 0; i < lista.length; i++) {
      const u = getFieldCI(lista[i], ["usuario", "username", "user", "email", "correo", "login"]);
      if (u && String(u).trim().toLowerCase() === username) {
        return { rec: lista[i], rol: rolForzado };
      }
    }
    return null;
  }

  let hit = buscar(users, null);
  if (!hit) hit = buscar(padres, "padre");

  if (!hit && !users.length && !padres.length) {
    const boot = { id: "ADM-" + new Date().getTime(), usuario: username, nombre: username, rol: "admin" };
    return jsonResponse({ success: true, bootstrap: true, user: boot, token: makeToken(boot) });
  }

  if (!hit) return jsonResponse({ success: false, message: "Usuario no encontrado" });

  const stored = String(getFieldCI(hit.rec, ["contraseña", "contrasena", "password", "clave", "pass", "pwd"]) || "");
  if (!stored) {
    return jsonResponse({ success: false, message: "La cuenta no tiene contraseña configurada. Contacta al administrador." });
  }
  if (stored !== password) {
    return jsonResponse({ success: false, message: "Contraseña incorrecta" });
  }

  const rol = normalizeRoleBackend(getFieldCI(hit.rec, ["rol", "role", "tipo"]) || hit.rol || "admin");
  const user = {
    id: hit.rec.id,
    usuario: getFieldCI(hit.rec, ["usuario", "username", "user"]) || username,
    nombre: getFieldCI(hit.rec, ["nombre", "name", "nombres"]) || username,
    rol: rol,
    childrenIds: getFieldCI(hit.rec, ["childrenIds", "hijos", "children", "estudiantesIds"]) || ""
  };

  return jsonResponse({ success: true, user: user, token: makeToken(user) });
}

// ==========================================
// SEGURIDAD: TOKENS DE SESIÓN, TICKETS Y LÍMITES
// ==========================================
const TOKEN_TTL_MS = 12 * 60 * 60 * 1000;  // sesión: 12 horas
const TICKET_TTL_MS = 15 * 60 * 1000;      // verificación de código: 15 minutos

function getSecret() {
  const props = PropertiesService.getScriptProperties();
  let s = props.getProperty("TOKEN_SECRET");
  if (!s) {
    s = Utilities.getUuid() + Utilities.getUuid();
    props.setProperty("TOKEN_SECRET", s);
  }
  return s;
}

// ✅ CORREGIDO: sin Utilities.Charset.UTF_8 (esa sobrecarga no existe)
function signPayload(payload) {
  const body = Utilities.base64EncodeWebSafe(JSON.stringify(payload));
  const sig = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(body, getSecret())
  );
  return body + "." + sig;
}

// ✅ CORREGIDO: sin Utilities.Charset.UTF_8
function readToken(token) {
  if (!token) return null;
  const parts = String(token).split(".");
  if (parts.length !== 2) return null;
  const body = parts[0];
  const expected = Utilities.base64EncodeWebSafe(
    Utilities.computeHmacSha256Signature(body, getSecret())
  );
  if (expected !== parts[1]) return null;
  let p = null;
  try {
    p = JSON.parse(Utilities.newBlob(Utilities.base64DecodeWebSafe(body)).getDataAsString());
  } catch (e) {
    return null;
  }
  if (!p || !p.e || Number(p.e) < Date.now()) return null;
  return p;
}

function makeToken(user) {
  return signPayload({
    t: "auth",
    u: String(user.id),
    r: String(user.rol || "padre"),
    n: String(user.nombre || ""),
    usuario: String(user.usuario || ""),
    e: Date.now() + TOKEN_TTL_MS
  });
}

function makeTicket(studentId) {
  return signPayload({ t: "reg", s: String(studentId), e: Date.now() + TICKET_TTL_MS });
}

function authFail(msg) {
  return jsonResponse({ success: false, code: "AUTH", error: msg || "Sesión no válida o expirada" });
}

function rateLimit(key, max, windowSec) {
  try {
    const cache = CacheService.getScriptCache();
    const k = "rl_" + key;
    const cur = Number(cache.get(k) || 0);
    if (cur >= max) return false;
    cache.put(k, String(cur + 1), windowSec);
    return true;
  } catch (e) {
    return true;
  }
}

function getFieldCI(obj, cands) {
  if (!obj) return "";
  const keys = Object.keys(obj);
  for (var i = 0; i < cands.length; i++) {
    const objetivo = String(cands[i]).toLowerCase();
    const k = keys.find(function (x) { return String(x).toLowerCase().trim() === objetivo; });
    if (k !== undefined && obj[k] !== undefined && obj[k] !== null) return obj[k];
  }
  return "";
}

function normalizeRoleBackend(v) {
  const s = String(v === undefined || v === null ? "" : v).toLowerCase();
  if (!s) return "padre";
  if (s.indexOf("admin") > -1 || s.indexOf("director") > -1) return "admin";
  if (s.indexOf("docen") > -1 || s.indexOf("profes") > -1 || s.indexOf("maestr") > -1) return "docente";
  if (s.indexOf("padre") > -1 || s.indexOf("madre") > -1 || s.indexOf("tutor") > -1 ||
      s.indexOf("apoderad") > -1 || s.indexOf("familia") > -1) return "padre";
  return s;
}

function parseChildIds(v) {
  if (v === undefined || v === null || v === "") return [];
  if (Array.isArray(v)) {
    return v.map(function (x) { return String(x).replace(/[\s\.\,]/g, "").trim(); }).filter(function (x) { return x !== ""; });
  }
  let s = String(v).trim();
  if (s.charAt(0) === "[") {
    try {
      const arr = JSON.parse(s);
      if (Array.isArray(arr)) {
        return arr.map(function (x) { return String(x).replace(/[\s\.\,]/g, "").trim(); }).filter(function (x) { return x !== ""; });
      }
    } catch (e) { /* sigue como texto */ }
  }
  // Si el valor viene con formato de miles de Google Sheets (ej. 17,777,854,507,211,700,000,000,000 o 1.77778E+24)
  // o separado por comas reales, se limpia cada fragmento
  return s.split(/[,;\n\t]+/).map(function (x) {
    // Si contiene comas o espacios internos de miles (ej. "17,777,854..."), remover comas que parezcan de miles si el fragmento no es una lista
    let clean = x.trim();
    // Si el fragmento es un número formateado con comas de miles (ej. 17,777,854) y no varios IDs cortos
    if (/^\d{1,3}(,\d{3})+$/.test(clean)) {
      clean = clean.replace(/,/g, "");
    }
    return clean;
  }).filter(function (x) { return x !== ""; });
}

function stripPasswords(rows) {
  return (rows || []).map(function (r) {
    const o = {};
    let has = 0;
    Object.keys(r).forEach(function (k) {
      if (/^(password|contraseña|contrasena|clave|pass|pwd)$/i.test(k)) {
        if (r[k] !== "" && r[k] !== null && r[k] !== undefined) has = 1;
        o[k] = "";
      } else {
        o[k] = r[k];
      }
    });
    o.__hasPass = has;
    return o;
  });
}

function scopePayload(payload, auth) {
  const role = String((auth && auth.r) || "").toLowerCase();
  payload.users = stripPasswords(payload.users);
  payload.padres = stripPasswords(payload.padres);

  if (role === "admin") return payload;

  if (role === "docente") {
    payload.users = [];
    payload.padres = [];
    payload.padres_hijos = [];
    return payload;
  }

  const me = (payload.padres || []).find(function (p) { return String(p.id) === String(auth.u); });
  const ids = {};
  parseChildIds(me && me.childrenIds).forEach(function (id) { ids[id] = true; });

  // Claves del padre: id de UsuariosPP, usuario del token y usuario de su fila
  const misClaves = [];
  [auth.u, auth.usuario, auth.n, me && getFieldCI(me, ["usuario", "username", "user", "email", "correo", "login"]), me && getFieldCI(me, ["nombre", "name", "nombres"])]
    .forEach(function (v) {
      const s = String(v === undefined || v === null ? "" : v).trim().toLowerCase();
      if (s && misClaves.indexOf(s) === -1) misClaves.push(s);
    });

  // Devuelve el primer valor no vacío entre los nombres de columna candidatos
  const pickLink = function (l, cands) {
    for (var i = 0; i < cands.length; i++) {
      const v = getFieldCI(l, [cands[i]]);
      if (v !== undefined && v !== null && String(v).trim() !== "") return String(v).trim();
    }
    return "";
  };

  // Vinculación por la hoja relacional PadresHijos
  const misLinks = [];
  (payload.padres_hijos || []).forEach(function (l) {
    const activo = String(pickLink(l, ["activo", "estado", "active", "vigente"]) || "SI").trim().toUpperCase();
    if (activo === "NO" || activo === "FALSE" || activo === "0" || activo === "INACTIVO") return;

    const sid = pickLink(l, ["studentId", "estudianteId", "alumnoId", "idEstudiante", "student_id", "estudiante", "alumno", "student"]);
    if (!sid) return;

    const padreId = pickLink(l, ["padreId", "parentId", "idPadre", "apoderadoId", "padre_id", "padre", "apoderado", "tutor"]).toLowerCase();
    const linkUser = pickLink(l, ["usuario", "username", "user", "email", "correo", "login"]).toLowerCase();

    const coincide = (padreId && misClaves.indexOf(padreId) > -1) || (linkUser && misClaves.indexOf(linkUser) > -1);
    if (!coincide) return;

    // El vínculo puede traer el id del estudiante o su nombre completo
    const sidKey = sid.toLowerCase();
    const byId = (payload.students || []).find(function (s) {
      return String(s.id === undefined || s.id === null ? "" : s.id).trim().toLowerCase() === sidKey;
    });
    if (byId) {
      ids[String(byId.id)] = true;
    } else {
      const byCode = (payload.students || []).find(function (s) {
        return String(getFieldCI(s, ["codigo", "code", "codigoEstudiante", "dni"]) || "").trim().toLowerCase() === sidKey;
      });
      if (byCode) {
        ids[String(byCode.id)] = true;
      } else {
        const byName = (payload.students || []).find(function (s) {
          return String(getFieldCI(s, ["nombre", "name", "nombres"]) || "").trim().toLowerCase() === sidKey;
        });
        if (byName) ids[String(byName.id)] = true;
      }
    }
    misLinks.push(l);
  });

  // Compatibilidad: columna padreId dentro de la hoja Estudiantes
  (payload.students || []).forEach(function (s) {
    if (String(s.padreId || "") === String(auth.u)) ids[String(s.id)] = true;
  });

  const keep = function (r) { return ids[String(r.studentId || "")] === true; };
  payload.students = (payload.students || []).filter(function (s) { return ids[String(s.id)] === true; });
  payload.grades = (payload.grades || []).filter(keep);
  payload.exams = (payload.exams || []).filter(keep);
  payload.attendance = (payload.attendance || []).filter(keep);
  payload.padres = me ? [me] : [];
  payload.padres_hijos = misLinks;
  payload.users = [];
  payload.teacher_assignments = [];
  return payload;
}

// ==========================================
// REGISTRO DE PADRES: verificación por código
// ==========================================
function verifyStudentCode(data) {
  const code = String((data && data.code) || "").trim().toUpperCase();
  if (!code) return jsonResponse({ success: false, message: "Ingresa el código del estudiante" });

  if (!rateLimit("vf_all", 120, 600) || !rateLimit("vf_" + code, 8, 600)) {
    return jsonResponse({
      success: false,
      code: "RATE",
      message: "Demasiados intentos. Espera unos minutos e intenta de nuevo."
    });
  }

  const grados = getSheetDataCached("Grados");
  const secciones = getSheetDataCached("Secciones");
  const students = normalizeStudentRefs(getSheetDataCached("Estudiantes"), grados, secciones) || [];

  const found = students.find(function (s) {
    const sc = s.code || s.codigo || s.codigoEstudiante;
    return String(sc === undefined || sc === null ? "" : sc).trim().toUpperCase() === code;
  });

  if (!found) {
    return jsonResponse({ success: false, message: "No se encontró ningún estudiante con ese código" });
  }

  const gradosLbl = found.grade || found.grado || "";
  const seccionLbl = found.section || found.seccion || "";

  return jsonResponse({
    success: true,
    student: { nombre: found.name || found.nombre || "Estudiante", grado: gradosLbl, seccion: seccionLbl },
    ticket: makeTicket(found.id)
  });
}

// ==========================================
// REGISTRO DE PADRES: creación de la cuenta
// ==========================================
function registerParent(data) {
  const p = (data && data.payload) || {};
  const name = String(p.name || p.nombre || "").trim();
  const username = String(p.username || p.usuario || "").trim();
  const password = String(p.password || p.contraseña || p.contrasena || "");

  if (!name) return jsonResponse({ success: false, message: "Escribe tu nombre completo" });
  if (username.length < 3) return jsonResponse({ success: false, message: "El usuario debe tener al menos 3 caracteres" });
  if (password.length < 6) return jsonResponse({ success: false, message: "La contraseña debe tener al menos 6 caracteres" });

  const ticket = readToken(p.ticket);
  if (!ticket || ticket.t !== "reg") {
    return jsonResponse({ success: false, code: "AUTH", message: "La verificación del estudiante expiró. Vuelve a verificar el código." });
  }

  if (!rateLimit("reg_all", 30, 3600)) {
    return jsonResponse({ success: false, code: "RATE", message: "Demasiados registros. Intenta más tarde." });
  }

  const users = getSheetDataCached("Usuarios") || [];
  const padres = getSheetDataCached("UsuariosPP") || [];
  const existe = users.concat(padres).some(function (u) {
    const val = getFieldCI(u, ["usuario", "username", "user", "email", "correo", "login"]);
    return String(val || "").trim().toLowerCase() === username.toLowerCase();
  });
  if (existe) return jsonResponse({ success: false, message: "Ese usuario ya está en uso. Elige otro." });

  const students = getSheetDataCached("Estudiantes") || [];
  const student = students.find(function (s) { return String(s.id) === String(ticket.s); });
  if (!student) return jsonResponse({ success: false, message: "El estudiante ya no existe. Vuelve a verificar el código." });

  const yaVinculado = padres.some(function (u) {
    return parseChildIds(getFieldCI(u, ["childrenIds", "hijos", "children", "estudiantesIds"])).indexOf(String(student.id)) > -1;
  });

  ensureSheetColumns("UsuariosPP", ["childrenIds"]);
  const sheet = ss.getSheetByName("UsuariosPP");
  if (!sheet) return jsonResponse({ success: false, message: "No existe la hoja UsuariosPP" });

  const lastCol = sheet.getLastColumn();
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  if (headers.indexOf("id") === -1) {
    return jsonResponse({ success: false, message: "La hoja UsuariosPP no tiene columna 'id'" });
  }

  const rec = {
    id: String(new Date().getTime()) + String(Math.floor(Math.random() * 100)).padStart(2, "0"),
    name: name,
    nombre: name,
    username: username,
    usuario: username,
    password: password,
    "contraseña": password,
    role: "padre",
    rol: "padre",
    childrenIds: String(student.id)
  };

  const row = headers.map(function (h) { return rec[h] !== undefined ? rec[h] : ""; });
  sheet.appendRow(row);
  cacheDropSheet("UsuariosPP");
  cacheDrop(CACHE_KEY);

  const user = { id: rec.id, usuario: username, nombre: name, rol: "padre", childrenIds: rec.childrenIds };
  return jsonResponse({
    success: true,
    user: user,
    token: makeToken(user),
    alreadyLinked: yaVinculado,
    message: "Registro completado"
  });
}

// ==========================================
// HELPERS
// ==========================================
function getSheetData(sheetName) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return [];

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return [];

  const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0].map(function (h) { return String(h).trim(); });

  const tz = Session.getScriptTimeZone();
  const out = [];
  for (var r = 1; r < all.length; r++) {
    const row = all[r];
    var empty = true;
    for (var c = 0; c < row.length; c++) {
      if (row[c] !== "" && row[c] !== null && row[c] !== undefined) { empty = false; break; }
    }
    if (empty) continue;

    const obj = {};
    for (var i = 0; i < headers.length; i++) {
      var val = row[i];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, tz, "yyyy-MM-dd");
      } else if (typeof val === "object" && val !== null) {
        if (typeof val.join === "function") {
          val = val.join(",");
        } else {
          val = String(val);
        }
      }
      obj[headers[i]] = val;
    }
    out.push(obj);
  }
  return out;
}

function getRefData(sheetName) {
  return getSheetDataCached(sheetName);
}

function backendLabel(o) {
  if (!o) return "";
  const ks = Object.keys(o);
  const pref = ["nombre", "name", "titulo", "descripcion", "description", "codigo"];
  for (var p = 0; p < pref.length; p++) {
    const k = ks.find(function (x) { return x.toLowerCase() === pref[p]; });
    if (k !== undefined && o[k] !== undefined && o[k] !== null && String(o[k]) !== "") return String(o[k]);
  }
  return String(o.id === undefined ? "" : o.id);
}

function refLabel(list, value) {
  if (value === undefined || value === null || value === "") return value;
  const s = String(value).trim();
  if (s === "") return value;
  const hit = (list || []).find(function (o) { return String(o.id).trim() === s; });
  return hit ? backendLabel(hit) : value;
}

function normalizeStudentRefs(rows, grados, secciones) {
  const out = [];
  (rows || []).forEach(function (r) {
    const o = Object.assign({}, r);
    Object.keys(o).forEach(function (k) {
      if (/^(grado|grade)$/i.test(k)) o[k] = refLabel(grados, o[k]);
      else if (/^(seccion|section|aula)$/i.test(k)) o[k] = refLabel(secciones, o[k]);
    });
    out.push(o);
  });
  return out;
}

// ==========================================
// MIGRACIÓN
// ==========================================
function migrateStudentLabels() {
  const sheet = ss.getSheetByName("Estudiantes");
  if (!sheet) return jsonResponse({ success: false, message: "Hoja 'Estudiantes' no encontrada" });

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 2 || lastCol < 1) return jsonResponse({ success: true, updated: 0, message: "La hoja no tiene datos" });

  const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0].map(function (h) { return String(h).trim(); });

  const gCol = headers.findIndex(function (h) { return /^(grado|grade)$/i.test(h); });
  const sCol = headers.findIndex(function (h) { return /^(seccion|section|aula)$/i.test(h); });

  const grados = getRefData("Grados");
  const secciones = getRefData("Secciones");

  const changedRows = [];
  for (var r = 1; r < all.length; r++) {
    let touched = false;
    if (gCol > -1) {
      const nv = refLabel(grados, all[r][gCol]);
      if (String(nv) !== String(all[r][gCol])) { all[r][gCol] = nv; touched = true; }
    }
    if (sCol > -1) {
      const nv = refLabel(secciones, all[r][sCol]);
      if (String(nv) !== String(all[r][sCol])) { all[r][sCol] = nv; touched = true; }
    }
    if (touched) changedRows.push(r + 1);
  }

  changedRows.forEach(function (rowNum) {
    sheet.getRange(rowNum, 1, 1, lastCol).setValues([all[rowNum - 1]]);
  });
  cacheDropSheet("Estudiantes");

  return jsonResponse({
    success: true,
    updated: changedRows.length,
    rows: changedRows,
    message: changedRows.length ? ("Filas corregidas: " + changedRows.join(", ")) : "No habia IDs que convertir"
  });
}

function ensureSheetColumns(sheetName, cols) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return;
  const lastCol = sheet.getLastColumn();
  if (lastCol < 1) return;
  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0]
    .map(function (h) { return String(h).trim(); });
  const missing = cols.filter(function (c) { return headers.indexOf(c) === -1; });
  if (!missing.length) return;
  sheet.getRange(1, lastCol + 1, 1, missing.length).setValues([missing]);
  cacheDropSheet(sheetName);
}

/**
 * Agrega la columna "cicloId" a una hoja si no existe.
 */
function ensureCycleColumns(sheetName) {
  ensureSheetColumns(sheetName, ["cicloId"]);
}

/**
 * Devuelve el id del ciclo marcado como activo (o el primero de la hoja "Ciclos").
 */
function activeCycleSheetId() {
  const rows = getSheetDataCached("Ciclos") || [];
  if (!rows.length) return "";
  const isOn = function (v) {
    if (v === true) return true;
    const s = String(v === undefined || v === null ? "" : v).trim().toLowerCase();
    return s === "true" || s === "1" || s === "si" || s === "sí" || s === "activo" || s === "x";
  };
  const act = rows.find(function (r) { return isOn(r.activo !== undefined ? r.activo : r.active); });
  const pick = act || rows[0];
  return String(pick.id === undefined || pick.id === null ? "" : pick.id);
}

/**
 * Ciclo base: el primero de la hoja "Ciclos" (donde quedan los datos históricos).
 */
function defaultCycleSheetId() {
  const rows = getSheetDataCached("Ciclos") || [];
  if (!rows.length) return "";
  return String(rows[0].id === undefined || rows[0].id === null ? "" : rows[0].id);
}

function saveRow(sheetName, payload, idKey) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ success: false, message: "Hoja no encontrada: " + sheetName });

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) {
    return jsonResponse({ success: false, message: "La hoja no tiene encabezados: " + sheetName });
  }

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const idIndex = headers.indexOf(idKey);

  if (idIndex === -1) return jsonResponse({ success: false, message: "No se encontró la columna '" + idKey + "' en la hoja " + sheetName + ". Encabezados actuales: " + headers.join(", ") });

  if (!payload[idKey]) {
    payload[idKey] = "ID-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000);
  }

  // Vincula automáticamente al ciclo activo si no se indicó uno.
  if ((sheetName === "Estudiantes" || sheetName === "Actividades") && headers.indexOf("cicloId") > -1) {
    const cyc = payload.cicloId;
    if (cyc === undefined || cyc === null || String(cyc).trim() === "") {
      const ac = activeCycleSheetId();
      if (ac) payload.cicloId = ac;
    }
  }

  let rowIndex = -1;
  if (lastRow > 1) {
    const ids = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(payload[idKey])) {
        rowIndex = i + 2;
        break;
      }
    }
  }

  const rowValues = headers.map(function (h) {
    var val = payload[h];
    if (val instanceof Date) {
      val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
    } else if (typeof val === "object" && val !== null) {
      if (typeof val.join === "function") {
        val = val.join(",");
      } else {
        val = String(val);
      }
    }
    return val !== undefined ? val : "";
  });

  if (sheetName === "UsuariosPP" && payload.childrenIds !== undefined) {
    // Forzar limpieza final de hijos para evitar formatos numéricos en celdas
    const ciIdx = headers.indexOf("childrenIds");
    if (ciIdx > -1) rowValues[ciIdx] = parseChildIds(rowValues[ciIdx]).join(",");
  }

  if (sheetName === "Estudiantes") {
    const gCol = headers.findIndex(function (h) { return /^(grado|grade)$/i.test(h); });
    const sCol = headers.findIndex(function (h) { return /^(seccion|section|aula)$/i.test(h); });
    if (gCol > -1 || sCol > -1) {
      if (gCol > -1) rowValues[gCol] = refLabel(getRefData("Grados"), rowValues[gCol]);
      if (sCol > -1) rowValues[sCol] = refLabel(getRefData("Secciones"), rowValues[sCol]);
    }
  }

  if ((sheetName === "Usuarios" || sheetName === "UsuariosPP") && rowIndex > 0) {
    const pCols = [];
    headers.forEach(function (h, i) {
      if (/^(password|contraseña|contrasena|clave|pass|pwd)$/i.test(h)) pCols.push(i);
    });
    if (pCols.length) {
      const actual = sheet.getRange(rowIndex, 1, 1, headers.length).getValues()[0];
      pCols.forEach(function (i) {
        const v = rowValues[i] === undefined || rowValues[i] === null ? "" : String(rowValues[i]);
        if (v === "") rowValues[i] = actual[i];
      });
    }
  }

  if (rowIndex > 0) {
    sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
  } else {
    sheet.appendRow(rowValues);
  }
  cacheDropSheet(sheetName);

  const record = {};
  headers.forEach(function (h, i) { record[h] = rowValues[i]; });

  if (sheetName === "Usuarios" || sheetName === "UsuariosPP") {
    let has = 0;
    Object.keys(record).forEach(function (k) {
      if (/^(password|contraseña|contrasena|clave|pass|pwd)$/i.test(k)) {
        if (record[k] !== "" && record[k] !== null && record[k] !== undefined) has = 1;
        record[k] = "";
      }
    });
    record.__hasPass = has;
  }

  return jsonResponse({ success: true, id: payload[idKey], record: record });
}

function deleteRow(sheetName, id, idKey) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ success: false, message: "Hoja no encontrada: " + sheetName });

  const lastRow = sheet.getLastRow();
  const lastCol = sheet.getLastColumn();
  if (lastRow < 1 || lastCol < 1) return jsonResponse({ success: false, message: "La hoja no tiene encabezados: " + sheetName });

  const headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  const idIndex = headers.indexOf(idKey);

  if (idIndex === -1) return jsonResponse({ success: false, message: "No se encontró la columna '" + idKey + "'" });

  if (lastRow > 1) {
    const ids = sheet.getRange(2, idIndex + 1, lastRow - 1, 1).getValues();
    for (var i = 0; i < ids.length; i++) {
      if (String(ids[i][0]) === String(id)) {
        sheet.deleteRow(i + 2);
        cacheDropSheet(sheetName);
        return jsonResponse({ success: true });
      }
    }
  }
  return jsonResponse({ success: false, message: "Registro no encontrado (id: " + id + ")" });
}

// ==========================================
// ACTUALIZACIÓN POR LOTE
// ==========================================
function batchUpdate(sheetName, payloadArray, requiredIdCols) {
  const sheet = ss.getSheetByName(sheetName);
  if (!sheet) return jsonResponse({ success: false, message: "Hoja no encontrada: " + sheetName });

  const lastCol = sheet.getLastColumn();
  if (lastCol === 0) return jsonResponse({ success: false, message: "La hoja '" + sheetName + "' no tiene encabezados" });

  const lastRow = Math.max(sheet.getLastRow(), 1);
  const all = sheet.getRange(1, 1, lastRow, lastCol).getValues();
  const headers = all[0].map(function (h) { return String(h).trim(); });
  const data = all;

  const missingCols = [];
  requiredIdCols.forEach(function (col) {
    if (headers.indexOf(col) === -1) missingCols.push(col);
  });
  if (missingCols.length > 0) {
    return jsonResponse({
      success: false,
      message: "La hoja '" + sheetName + "' no tiene las columnas requeridas: " + missingCols.join(", ") + ". Encabezados actuales: " + headers.join(", ")
    });
  }

  const idIdx = headers.indexOf("id");
  const sIdIdx = headers.indexOf("studentId");
  const aIdIdx = headers.indexOf("activityId");
  const mIdIdx = headers.indexOf("materiaId");
  const bIdIdx = headers.indexOf("bimestreId");
  const dIdx = headers.indexOf("date");
  const cIdIdx = headers.indexOf("cicloId");

  // Vincula automáticamente al ciclo activo los registros sin ciclo.
  if (cIdIdx > -1) {
    const ac = activeCycleSheetId();
    payloadArray.forEach(function (item) {
      if (item && (item.cicloId === undefined || item.cicloId === null || String(item.cicloId).trim() === "")) {
        if (ac) item.cicloId = ac;
      }
    });
  }

  const sameCycle = function (row, item) {
    if (cIdIdx === -1) return true;
    const rc = String(row[cIdIdx] === undefined || row[cIdIdx] === null ? "" : row[cIdIdx]).trim();
    const ic = String(item.cicloId === undefined || item.cicloId === null ? "" : item.cicloId).trim();
    return !rc || !ic || rc === ic;
  };

  const notaLike = ["nota", "score", "puntaje", "valor", "calificacion", "notaFinal", "points"];
  let notaCol = null;
  notaLike.forEach(function (c) {
    if (notaCol === null && headers.indexOf(c) !== -1) notaCol = c;
  });

  let updated = 0;
  let inserted = 0;

  payloadArray.forEach(function (item) {
    let rowIndex = -1;

    if (!item.id) {
      item.id = "ID-" + new Date().getTime() + "-" + Math.floor(Math.random() * 10000);
    }

    if (notaCol !== null && (item[notaCol] === undefined || item[notaCol] === "")) {
      for (var n = 0; n < notaLike.length; n++) {
        if (item[notaLike[n]] !== undefined && item[notaLike[n]] !== "") {
          item[notaCol] = item[notaLike[n]];
          break;
        }
      }
    }

    for (let i = 1; i < data.length; i++) {
      if (sheetName === "Notas") {
        if (String(data[i][sIdIdx]) === String(item.studentId) &&
            String(data[i][aIdIdx]) === String(item.activityId) &&
            sameCycle(data[i], item)) {
          rowIndex = i + 1;
          break;
        }
      } else if (sheetName === "Examenes") {
        if (String(data[i][sIdIdx]) === String(item.studentId) &&
            String(data[i][mIdIdx]) === String(item.materiaId) &&
            String(data[i][bIdIdx]) === String(item.bimestreId) &&
            sameCycle(data[i], item)) {
          rowIndex = i + 1;
          break;
        }
      } else if (sheetName === "Asistencia") {
        let rowDate = data[i][dIdx];
        let dateStr = "";
        if (rowDate instanceof Date) {
          dateStr = Utilities.formatDate(rowDate, Session.getScriptTimeZone(), "yyyy-MM-dd");
        } else {
          dateStr = String(rowDate).split("T")[0].trim();
        }
        let itemDate = String(item.date).split("T")[0].trim();
        if (String(data[i][sIdIdx]) === String(item.studentId) && dateStr === itemDate && sameCycle(data[i], item)) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    const rowValues = headers.map(function (h) {
      var val = item[h];
      if (val instanceof Date) {
        val = Utilities.formatDate(val, Session.getScriptTimeZone(), "yyyy-MM-dd");
      } else if (typeof val === "object" && val !== null) {
        if (typeof val.join === "function") {
          val = val.join(",");
        } else {
          val = String(val);
        }
      }
      return val !== undefined ? val : "";
    });

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, headers.length).setValues([rowValues]);
      updated++;
    } else {
      sheet.appendRow(rowValues);
      inserted++;
    }
  });

  cacheDropSheet(sheetName);
  return jsonResponse({ success: true, updated: updated, inserted: inserted });
}

// ==========================================
// ASISTENCIA
// ==========================================
function normalizeAttendanceStatus(v) {
  if (v === undefined || v === null || v === "") return "";
  var k = String(v).trim().toLowerCase();
  var aliases = {
    present: "present", presente: "present", p: "present", as: "present",
    absent: "absent", ausente: "absent", a: "absent", falto: "absent", no: "absent",
    late: "late", tarde: "late", tardanza: "late", t: "late",
    justified: "justified", justificado: "justified", justificacion: "justified",
    justificada: "justified", j: "justified", perm: "justified", excused: "justified"
  };
  return aliases[k] || k;
}

function saveAttendance(payload) {
  var sheet = ss.getSheetByName("Asistencia");
  if (!sheet) return jsonResponse({ success: false, error: "Hoja 'Asistencia' no encontrada" });

  var lastCol = sheet.getLastColumn();
  if (lastCol === 0) {
    sheet.getRange(1, 1, 1, 4).setValues([["studentId", "date", "status", "comment"]]);
    lastCol = 4;
  }
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });

  function findCol(names) {
    for (var n = 0; n < names.length; n++) {
      for (var i = 0; i < headers.length; i++) {
        if (headers[i].toLowerCase() === names[n].toLowerCase()) return i;
      }
    }
    return -1;
  }

  var statusCandidates = ["status", "estado", "asistencia"];
  var statusCol = findCol(statusCandidates);
  if (statusCol === -1) {
    sheet.getRange(1, lastCol + 1).setValue("status");
    headers.push("status");
    statusCol = lastCol;
    lastCol++;
  }

  var commentCandidates = ["comment", "comentario", "observacion"];
  var commentCol = findCol(commentCandidates);
  if (commentCol === -1) {
    sheet.getRange(1, lastCol + 1).setValue("comment");
    headers.push("comment");
    commentCol = lastCol;
    lastCol++;
  }

  var sidCol = findCol(["studentId", "studentid", "estudianteId"]);
  var dCol = findCol(["date", "fecha"]);
  if (sidCol === -1 || dCol === -1) {
    return jsonResponse({
      success: false,
      error: "La hoja 'Asistencia' necesita las columnas studentId y date. Encabezados actuales: " + headers.join(", ")
    });
  }
  var idCol = findCol(["id"]);
  var cycCol = findCol(["cicloId"]);

  var data = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), lastCol).getValues();
  var tz = Session.getScriptTimeZone();
  var updated = 0;
  var inserted = 0;

  (payload || []).forEach(function (item) {
    if (!item || item.studentId === undefined || item.studentId === "") return;

    var itemDate = String(item.date || "").split("T")[0].trim();

    var rowIndex = -1;
    for (var i = 1; i < data.length; i++) {
      var rowDate = data[i][dCol];
      var dateStr = "";
      if (rowDate instanceof Date) {
        dateStr = Utilities.formatDate(rowDate, tz, "yyyy-MM-dd");
      } else {
        dateStr = String(rowDate).split("T")[0].trim();
      }
      if (String(data[i][sidCol]) === String(item.studentId) && dateStr === itemDate) {
        var rowCyc = cycCol > -1 ? String(data[i][cycCol] === undefined || data[i][cycCol] === null ? "" : data[i][cycCol]).trim() : "";
        var itemCyc = String(item.cicloId === undefined || item.cicloId === null ? "" : item.cicloId).trim();
        if (!rowCyc || !itemCyc || rowCyc === itemCyc) {
          rowIndex = i + 1;
          break;
        }
      }
    }

    var rawStatus = "";
    var lookKeys = statusCandidates.concat([headers[statusCol]]);
    for (var k = 0; k < lookKeys.length; k++) {
      if (item[lookKeys[k]] !== undefined && item[lookKeys[k]] !== "") { rawStatus = item[lookKeys[k]]; break; }
    }
    var commentVal = "";
    var cKeys = commentCandidates.concat([headers[commentCol]]);
    for (var c = 0; c < cKeys.length; c++) {
      if (item[cKeys[c]] !== undefined && item[cKeys[c]] !== null && item[cKeys[c]] !== "") { commentVal = item[cKeys[c]]; break; }
    }

    var values = [];
    for (var h = 0; h < lastCol; h++) {
      if (h === statusCol) {
        values.push(normalizeAttendanceStatus(rawStatus));
      } else if (h === commentCol) {
        values.push(String(commentVal));
      } else if (h === sidCol) {
        values.push(item.studentId);
      } else if (h === dCol) {
        values.push(itemDate);
      } else if (h === idCol) {
        if (item.id) {
          values.push(item.id);
        } else if (rowIndex > 0) {
          var prevId = data[rowIndex - 1][idCol];
          values.push(prevId === undefined || prevId === null || prevId === "" ? "ID-" + new Date().getTime() : prevId);
        } else {
          values.push("ID-" + new Date().getTime() + "-" + Math.floor(Math.random() * 10000));
        }
      } else if (cycCol > -1 && h === cycCol) {
        values.push(item.cicloId && String(item.cicloId).trim() !== "" ? item.cicloId : activeCycleSheetId());
      } else {
        var v = item[headers[h]];
        if ((v === undefined || v === null || v === "") && rowIndex > 0) {
          var prev = data[rowIndex - 1][h];
          v = (prev === undefined || prev === null) ? "" : prev;
        }
        if (v === undefined || v === null) v = "";
        if (v instanceof Date) v = Utilities.formatDate(v, tz, "yyyy-MM-dd");
        values.push(v);
      }
    }

    if (rowIndex > 0) {
      sheet.getRange(rowIndex, 1, 1, lastCol).setValues([values]);
      updated++;
    } else {
      sheet.appendRow(values);
      inserted++;
    }
  });

  cacheDropSheet("Asistencia");
  return jsonResponse({ success: true, updated: updated, inserted: inserted });
}

/**
 * Guarda un ciclo escolar en la hoja "Ciclos" (crear o editar).
 * Payload esperado: { id?, nombre, fechaInicio, fechaFin, activo, activate? }
 * - Sin "id": genera el siguiente número disponible (2024, 2025, ...).
 * - Con "id": actualiza ese ciclo.
 * - Si "activo"/"activate" es verdadero, deja ese ciclo como único activo.
 */
function saveCycle(payload) {
  payload = payload || {};
  var sheet = ss.getSheetByName("Ciclos");
  if (!sheet) {
    sheet = ss.getActiveSpreadsheet().insertSheet("Ciclos");
    sheet.getRange(1, 1, 1, 5).setValues([["id", "nombre", "fechaInicio", "fechaFin", "activo"]]);
  }

  var lastCol = sheet.getLastColumn();
  var headers = sheet.getRange(1, 1, 1, lastCol).getValues()[0].map(function (h) { return String(h).trim(); });
  var col = function (name) { return headers.indexOf(name); };
  var idCol = col("id"), nomCol = col("nombre"), iniCol = col("fechaInicio"), finCol = col("fechaFin"), actCol = col("activo");
  if (idCol === -1) return jsonResponse({ success: false, error: "La hoja 'Ciclos' no tiene la columna 'id'" });

  var lastRow = sheet.getLastRow();
  var rows = lastRow > 1 ? sheet.getRange(2, 1, lastRow - 1, lastCol).getValues() : [];

  var wantedId = payload.id === undefined || payload.id === null ? "" : String(payload.id).trim();
  var rowIndex = -1;
  var existingIds = [];

  for (var i = 0; i < rows.length; i++) {
    var rowId = String(rows[i][idCol] === undefined ? "" : rows[i][idCol]).trim();
    if (!rowId) continue;
    if (wantedId && rowId === wantedId) rowIndex = i;
    var num = Number(rowId.replace(/[^0-9]/g, ""));
    if (!isNaN(num)) existingIds.push(num);
  }

  var nextId = existingIds.length > 0 ? Math.max.apply(Math, existingIds) + 1 : new Date().getFullYear();
  var targetId = wantedId || String(nextId);
  var nombre = payload.nombre !== undefined && String(payload.nombre).trim() !== ""
    ? String(payload.nombre).trim()
    : (rowIndex > -1 ? String(rows[rowIndex][nomCol] || targetId) : "Ciclo " + targetId);
  var fechaInicio = payload.fechaInicio !== undefined ? payload.fechaInicio
    : (rowIndex > -1 && iniCol > -1 ? rows[rowIndex][iniCol] : "01/02/" + targetId);
  var fechaFin = payload.fechaFin !== undefined ? payload.fechaFin
    : (rowIndex > -1 && finCol > -1 ? rows[rowIndex][finCol] : "30/11/" + targetId);
  var activo = payload.activo !== undefined ? payload.activo
    : (rowIndex > -1 && actCol > -1 ? rows[rowIndex][actCol] : true);
  var makeActive = payload.activate === true || activo === true || String(activo).toLowerCase() === "true";

  // Un solo ciclo activo a la vez
  if (makeActive && actCol > -1) {
    var offs = [];
    for (var r = 0; r < rows.length; r++) {
      var rid = String(rows[r][idCol] === undefined ? "" : rows[r][idCol]).trim();
      if (rid && rid !== targetId) offs.push(r);
    }
    offs.forEach(function (r) {
      sheet.getRange(r + 2, actCol + 1).setValue("false");
    });
  }

  var values = [];
  for (var c = 0; c < headers.length; c++) {
    var h = headers[c];
    if (h === "id") values.push(targetId);
    else if (h === "nombre") values.push(nombre);
    else if (h === "fechaInicio") values.push(fechaInicio);
    else if (h === "fechaFin") values.push(fechaFin);
    else if (h === "activo") values.push(makeActive ? "true" : "false");
    else values.push("");
  }

  if (rowIndex > -1) sheet.getRange(rowIndex + 2, 1, 1, headers.length).setValues([values]);
  else sheet.appendRow(values);

  cacheDropSheet("Ciclos");
  cacheDropAll();

  var record = {};
  headers.forEach(function (h, idx) { record[h] = values[idx]; });
  return jsonResponse({ success: true, id: targetId, nombre: nombre, record: record });
}

/**
 * Elimina un ciclo de la hoja "Ciclos".
 */
function deleteCycle(payload) {
  return deleteRow("Ciclos", payload && payload.id, "id");
}

// ==========================================
// ASIGNACIONES DOCENTES
// ==========================================
function saveAssignments(payload) {
  let sheet = ss.getSheetByName("AsignacionesDocentes");
  if (!sheet) return jsonResponse({ success: false, message: "Hoja 'AsignacionesDocentes' no encontrada" });

  const data = sheet.getRange(1, 1, Math.max(sheet.getLastRow(), 1), Math.max(sheet.getLastColumn(), 1)).getValues();

  for (let i = data.length - 1; i > 0; i--) {
    if (String(data[i][1]) === String(payload.userId)) {
      sheet.deleteRow(i + 1);
    }
  }

  if (payload.assignments && payload.assignments.length > 0) {
    payload.assignments.forEach(function (asig) {
      sheet.appendRow([
        "ASG-" + new Date().getTime() + "-" + Math.floor(Math.random() * 1000),
        payload.userId,
        asig.gradoId,
        asig.seccionId
      ]);
    });
  }

  cacheDropSheet("AsignacionesDocentes");
  return jsonResponse({ success: true });
}

function jsonResponse(obj) {
  return jsonText(JSON.stringify(obj));
}

function jsonText(txt) {
  return ContentService.createTextOutput(txt)
    .setMimeType(ContentService.MimeType.JSON);
}