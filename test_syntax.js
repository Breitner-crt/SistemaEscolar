
/**
 * SISTEMA ESCOLAR - FRONTEND ROBUSTO Y CORREGIDO
 * 1. Codificación limpia UTF-8 sin caracteres corruptos.
 * 2. Comunicación robusta con Google Apps Script (redirect follow, timeout, fallback).
 * 3. Diseño refinado Dark Glassmorphism, responsivo y dinámico.
 * 4. Exportación nativa a Excel (.xlsx) e impresión de reportes y calificaciones.
 */

const DEFAULT_API_URL = "";
const LS_URL = "escolar_api_url";
const LS_SESSION = "escolar_session";

// 1. HEADERS NORMALIZADOS EN UTF-8
const FALLBACK_HEADERS = {
  students: ["id","nombre","gradoId","seccionId","dni","fechaNacimiento","sexo","direccion","telefono","padreId"],
  grados: ["id","nombre","descripcion"],
  secciones: ["id","nombre","gradoId"],
  materias: ["id","nombre","descripcion"],
  bimestres: ["id","nombre","fechaInicio","fechaFin"],
  activities: ["id","nombre","descripcion","gradoId","materiaId","bimestreId","fecha","puntajeMax"],
  users: ["id","usuario","contraseña","nombre","rol"],
  padres: ["id","nombre","usuario","contraseña","telefono"],
  grades: ["id","studentId","activityId","nota"],
  exams: ["id","studentId","materiaId","bimestreId","nota"],
  attendance: ["id","studentId","date","estado"],
  teacher_assignments: ["id","userId","gradoId","seccionId"]
};

const SAVE_ACTIONS = {
  students: "save_student", grados: "save_grado", secciones: "save_seccion",
  materias: "save_materia", activities: "save_actividad", bimestres: "save_bimestre",
  users: "save_user", padres: "save_parent_user"
};

const DELETE_ACTIONS = {
  students: "delete_student", grados: "delete_grados", secciones: "delete_secciones",
  materias: "delete_materias", activities: "delete_activities"
};

const CRUD_CFG = {
  grados: { key: "grados", title: "Grados", desc: "Niveles o grados escolares del plantel", add: "Nuevo Grado" },
  secciones: { key: "secciones", title: "Secciones", desc: "Secciones organizadas por grado", add: "Nueva Sección" },
  materias: { key: "materias", title: "Materias", desc: "Asignaturas curriculares", add: "Nueva Materia" },
  bimestres: { key: "bimestres", title: "Bimestres", desc: "Períodos lectivos del año escolar", add: "Nuevo Bimestre", noDelete: true },
  users: { key: "users", title: "Usuarios del Sistema", desc: "Cuentas de acceso para administradores y docentes", add: "Nuevo Usuario", noDelete: true },
  padres: { key: "padres", title: "Padres de Familia", desc: "Registro de tutores y apoderados", add: "Nuevo Tutor", noDelete: true },
  activities: { key: "activities", title: "Actividades", desc: "Tareas, prácticas y evaluaciones continuas", add: "Nueva Actividad" }
};

const NAV = [
  { view: "dashboard", label: "Panel", roles: ["admin", "docente", "padre"], icon: "grid" },
  { view: "students", label: "Estudiantes", roles: ["admin", "docente", "padre"], icon: "users" },
  { sec: "Académico", roles: ["admin", "docente", "padre"] },
  { view: "activities", label: "Actividades", roles: ["admin", "docente"], icon: "book" },
  { view: "grades", label: "Calificaciones", roles: ["admin", "docente", "padre"], icon: "award" },
  { view: "exams", label: "Exámenes", roles: ["admin", "docente", "padre"], icon: "file-text" },
  { view: "attendance", label: "Asistencia", roles: ["admin", "docente", "padre"], icon: "check-circle" },
  { sec: "Configuración", roles: ["admin", "docente"] },
  { view: "grados", label: "Grados", roles: ["admin"], icon: "layers" },
  { view: "secciones", label: "Secciones", roles: ["admin"], icon: "folder" },
  { view: "materias", label: "Materias", roles: ["admin", "docente"], icon: "bookmark" },
  { view: "bimestres", label: "Bimestres", roles: ["admin"], icon: "calendar" },
  { sec: "Administración", roles: ["admin"] },
  { view: "users", label: "Usuarios", roles: ["admin"], icon: "user-check" },
  { view: "padres", label: "Padres de Familia", roles: ["admin"], icon: "heart" },
  { view: "assignments", label: "Asignaciones Docentes", roles: ["admin"], icon: "link" }
];

const BATCH = {
  grades: {
    key: "grades", title: "Calificaciones de Actividades", desc: "Registro masivo de notas por actividad",
    system: ["id", "studentId", "activityId"],
    controls: [{ k: "activity", label: "Actividad", db: "activities" }, { k: "grado", label: "Grado", db: "grados" }, { k: "seccion", label: "Sección", db: "secciones" }],
    required: ["activity"],
    fallback: [{ key: "nota", label: "Nota", type: "number", candidates: ["nota", "score", "puntaje", "valor", "calificacion", "notaFinal"] }],
    relFill: { gradoId: "grado", seccionId: "seccion", materiaId: "materia", bimestreId: "bimestre" }
  },
  exams: {
    key: "exams", title: "Calificaciones de Exámenes", desc: "Notas de exámenes bimestrales por materia",
    system: ["id", "studentId", "materiaId", "bimestreId"],
    controls: [{ k: "materia", label: "Materia", db: "materias" }, { k: "bimestre", label: "Bimestre", db: "bimestres" }, { k: "grado", label: "Grado", db: "grados" }, { k: "seccion", label: "Sección", db: "secciones" }],
    required: ["materia", "bimestre"],
    fallback: [{ key: "nota", label: "Nota", type: "number", candidates: ["nota", "score", "puntaje", "valor", "calificacion", "notaFinal"] }],
    relFill: { gradoId: "grado", seccionId: "seccion" }
  },
  attendance: {
    key: "attendance", title: "Control de Asistencia", desc: "Registro diario de puntualidad y asistencia",
    system: ["id", "studentId", "date"],
    controls: [{ k: "fecha", label: "Fecha", type: "date" }, { k: "grado", label: "Grado", db: "grados" }, { k: "seccion", label: "Sección", db: "secciones" }],
    required: ["fecha"],
    fallback: [{ key: "estado", label: "Estado", type: "select", options: ["Presente", "Ausente", "Tardanza", "Justificado"], candidates: ["estado", "asistencia", "status"] }],
    relFill: { gradoId: "grado", seccionId: "seccion" }
  }
};

let API_URL = DEFAULT_API_URL;
let DB = {};
let USER = null;
let VIEW = "dashboard";
const state = {
  q: {},
  students: { grado: "", seccion: "" },
  batch: {
    grades: { activity: "", grado: "", seccion: "", materia: "", bimestre: "" },
    exams: { activity: "", grado: "", seccion: "", materia: "", bimestre: "" },
    attendance: { fecha: "", grado: "", seccion: "", activity: "", materia: "", bimestre: "" }
  },
  matrixGrades: {
    bimestre: "",
    grado: "",
    seccion: "",
    materia: "",
    q: ""
  },
  assign: { userId: "" }
};

const $ = s => document.querySelector(s);
const $$ = s => Array.from(document.querySelectorAll(s));

function uid() {
  // Genera ID numérico de 12-13 dígitos basado en Date.now() y aleatorio (ej: 1778900788400)
  return String(Date.now()) + Math.floor(Math.random() * 90 + 10);
}

function esc(v) {
  return String(v === undefined || v === null ? "" : v)
    .replace(/&/g, "&amp;")
    .replace(/</g, "&lt;")
    .replace(/>/g, "&gt;")
    .replace(/"/g, "&quot;")
    .replace(/'/g, "&#39;");
}

function today() {
  const d = new Date();
  return d.getFullYear() + "-" + String(d.getMonth() + 1).padStart(2, "0") + "-" + String(d.getDate()).padStart(2, "0");
}

function pretty(k) {
  let s = String(k).replace(/[_-]+/g, " ").replace(/([a-z])([A-Z])/g, "$1 $2");
  const kind = relKindOfColumn(s);
  if (kind === "materia") return "Materia";
  if (kind === "bimestre") return "Bimestre";
  if (kind === "grado") return "Grado";
  if (kind === "seccion") return "Sección";
  if (/contrase/i.test(s)) return "Contraseña";
  return s.replace(/\b\w/g, c => c.toUpperCase()).trim();
}

function datePart(v) {
  if (v === undefined || v === null) return "";
  const s = String(v);
  return s.length >= 10 ? s.slice(0, 10) : s;
}

function relKindOfColumn(h) {
  if (/grado|grade/i.test(h) && !/seccion|secci|section/i.test(h)) return "grado";
  if (/seccion|secci|section|aula/i.test(h)) return "seccion";
  if (/materia|subject|course|asignatura/i.test(h)) return "materia";
  if (/bimestre|term|periodo/i.test(h)) return "bimestre";
  if (/actividad|activity|tarea/i.test(h)) return "actividad";
  return null;
}

function displayValue(h, v) {
  if (v === undefined || v === null || v === "") return "";
  const kind = relKindOfColumn(h);
  if (kind) {
    const list = DB[relDb(kind)] || [];
    const id = String(v).trim();
    let hit = list.find(o => String(o.id) === id);
    if (!hit) hit = list.find(o => String(labelOf(o)).toLowerCase() === id.toLowerCase());
    if (hit) return labelOf(hit);
  }
  if (/fecha|date|inicio|fin/i.test(h)) return datePart(v);
  if (/contrase|password|clave/i.test(h)) return "••••••••";
  return String(v);
}

function labelOf(o) {
  if (!o) return "";
  const ks = Object.keys(o);
  const pref = ["nombre", "name", "titulo", "descripcion", "description", "codigo", "usuario", "apellido", "etiqueta", "email"];
  for (const p of pref) {
    const k = ks.find(k => k.toLowerCase() === p);
    if (k && String(o[k]).trim() !== "") return String(o[k]);
  }
  const k = ks.find(k => k !== "id" && typeof o[k] === "string" && String(o[k]).trim() !== "");
  if (k) return String(o[k]);
  return String(o.id !== undefined ? o.id : "");
}

function unionKeys(records, fallback) {
  const set = [];
  (records || []).forEach(r => Object.keys(r).forEach(k => {
    const cleanK = k.replace(/Ã±/g, "ñ").replace(/Ã¡/g, "á").replace(/Ã©/g, "é").replace(/Ã­/g, "í").replace(/Ã³/g, "ó").replace(/Ãº/g, "ú");
    if (set.indexOf(cleanK) === -1) set.push(cleanK);
  }));
  if (set.length === 0) (fallback || []).forEach(k => set.push(k));
  if (set.indexOf("id") === -1) set.unshift("id");
  return set;
}

function pickKey(records, cands, re) {
  const keys = records && records.length ? Object.keys(records[0]) : [];
  for (const c of cands) {
    const k = keys.find(k => k.toLowerCase() === c.toLowerCase());
    if (k) return k;
  }
  return keys.find(k => re.test(k)) || null;
}

function toast(msg, type) {
  const d = document.createElement("div");
  d.className = "toast" + (type ? " " + type : "");
  d.innerHTML = `<span>${esc(msg)}</span>`;
  $("#toasts").appendChild(d);
  setTimeout(() => {
    d.style.opacity = "0";
    d.style.transform = "translateY(-10px)";
    d.style.transition = "all 0.25s";
    setTimeout(() => d.remove(), 250);
  }, 3500);
}

function setLoading(on, text = "Cargando...") {
  const el = $("#loading");
  $("#loadingText").textContent = text;
  el.hidden = !on;
}

// 2. COMUNICACIÓN ROBUSTA CON GOOGLE APPS SCRIPT
async function apiGet(action) {
  if (!API_URL) throw new Error("No se ha configurado la URL del backend");
  const url = API_URL + (API_URL.indexOf("?") > -1 ? "&" : "?") + "action=" + encodeURIComponent(action);
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 15000);
  try {
    const res = await fetch(url, {
      method: "GET",
      redirect: "follow",
      signal: controller.signal
    });
    clearTimeout(timeoutId);
    return await res.json();
  } catch (err) {
    clearTimeout(timeoutId);
    if (err.name === 'AbortError') throw new Error("Tiempo de espera agotado al conectar con Google Sheets.");
    throw err;
  }
}

async function apiPost(body) {
  if (!API_URL) throw new Error("No se ha configurado la URL del backend");
  const controller = new AbortController();
  const timeoutId = setTimeout(() => controller.abort(), 20000);
  try {
    const res = await fetch(API_URL, {
      method: "POST",
      redirect: "follow",
      headers: { "Content-Type": "text/plain;charset=utf-8" },
      body: JSON.stringify(body),
      signal: controller.signal
    });
    clearTimeout(timeoutId);
    return await res.json();
  } catch (err) {
    clearTimeout(timeoutId);
    if (err.name === 'AbortError') throw new Error("Tiempo de espera agotado al guardar.");
    throw err;
  }
}

async function loadData() {
  const r = await apiGet("get_data");
  if (r && r.error) throw new Error(r.error);
  DB = r || {};
}

async function reload() {
  await loadData();
  renderView();
}

function relField(rec, kind) {
  if (!rec) return null;
  const res = {
    grado: /grado|grade/i,
    seccion: /secci|seccion|section|aula/i,
    materia: /materia|subject|course|asignatura/i,
    bimestre: /bimestre|term|periodo/i,
    actividad: /actividad|activity|tarea/i
  };
  return Object.keys(rec).find(k => res[kind].test(k)) || null;
}

function relDb(kind) {
  return { grado: "grados", seccion: "secciones", materia: "materias", bimestre: "bimestres", actividad: "activities" }[kind];
}

function relLabel(kind, val) {
  const list = DB[relDb(kind)] || [];
  const hit = list.find(o => String(o.id) === String(val));
  return hit ? labelOf(hit) : null;
}

function relMatch(rec, kind, val) {
  const f = relField(rec, kind);
  if (!f) return false;
  const v = String(rec[f] === undefined ? "" : rec[f]).trim();
  if (!v) return false;
  if (v === String(val)) return true;
  const lbl = relLabel(kind, val);
  if (lbl && v === String(lbl).trim()) return true;
  return false;
}

function optionsHtml(dbKey, current, placeholder, filterByGrado = null, allowAllFallback = false) {
  let opts = DB[dbKey] || [];
  
  if (dbKey === "secciones" && filterByGrado) {
    const filtered = opts.filter(o => {
      // Si la sección tiene gradoId, verificar si coincide por ID o por nombre del grado
      if (o.gradoId !== undefined && o.gradoId !== null && String(o.gradoId).trim() !== "") {
        if (String(o.gradoId).trim() === String(filterByGrado).trim()) return true;
        const gLabel = relLabel("grado", filterByGrado);
        if (gLabel && String(o.gradoId).trim().toLowerCase() === String(gLabel).trim().toLowerCase()) return true;
        const oGLabel = relLabel("grado", o.gradoId);
        if (oGLabel && String(oGLabel).trim().toLowerCase() === String(filterByGrado).trim().toLowerCase()) return true;
        return false;
      }
      return true; // si en la base de datos la sección no tiene gradoId, dejarla disponible
    });

    // Si el filtrado estricto dio opciones, las usamos; si no encontró ninguna y se permite fallback, mostramos todas
    if (filtered.length > 0) {
      opts = filtered;
    } else if (allowAllFallback) {
      // Mantener opts con todas las secciones disponibles
    } else {
      opts = filtered;
    }
  }

  if (USER && USER.rol === "docente") {
    if (dbKey === "grados") {
      const sc = scopedGrados();
      opts = opts.filter(o => sc.some(v => relMatchEquivalent("grado", v, o.id)));
    } else if (dbKey === "secciones") {
      const sc = scopedSecciones();
      opts = opts.filter(o => sc.some(v => relMatchEquivalent("seccion", v, o.id)));
    } else if (dbKey === "materias") {
      opts = scopeMaterias();
    } else if (dbKey === "activities") {
      const acts = scopeActivities() || [];
      const okIds = acts.map(a => String(a.id));
      opts = opts.filter(o => okIds.indexOf(String(o.id)) > -1);
    }
  }

  const cur = current === undefined || current === null ? "" : String(current);
  let hasId = false;
  let html = "";
  const ph = placeholder === undefined ? "— Seleccionar —" : placeholder;
  if (ph) html += '<option value="">' + esc(ph) + '</option>';
  
  opts.forEach(o => {
    const id = String(o.id === undefined ? "" : o.id);
    const l = String(labelOf(o));
    if (cur && cur === id) hasId = true;
    html += '<option value="' + esc(id) + '"' + (cur === id ? ' selected' : '') + '>' + esc(l) + '</option>';
  });
  if (cur && !hasId && cur !== "") html += '<option value="' + esc(cur) + '" selected>' + esc(cur) + '</option>';
  return html;
}

function scopeActivities() {
  if (!(USER && USER.rol === "docente")) return null;
  return (DB.activities || []).filter(a => inScope("grado", relGet(a, "grado")));
}

function esMateriaDelDocente(m) {
  if (!(USER && USER.rol === "docente")) return true;
  const g = relGet(m, "grado");
  if (g === undefined || g === null || g === "") return true;
  return inScope("grado", g);
}

function scopeMaterias() {
  if (!(USER && USER.rol === "docente")) return DB.materias || [];
  return (DB.materias || []).filter(esMateriaDelDocente);
}

function relGet(rec, kind) {
  const f = relField(rec, kind);
  return f ? rec[f] : undefined;
}

function userKeys() {
  const recs = DB.users || [];
  if (!recs.length) return { u: null, p: null, r: null, n: null };
  return {
    u: pickKey(recs, ["usuario", "username", "user", "email", "correo", "cuenta", "login"], /usuario|user|email|correo|cuenta|login/i),
    p: pickKey(recs, ["contraseña", "contrasena", "password", "clave", "pass", "pwd", "contraseÃ±a"], /contrase|password|clave|pwd/i),
    r: pickKey(recs, ["rol", "role", "tipo"], /^(rol|role|tipo)$/i),
    n: pickKey(recs, ["nombre", "name", "nombres", "fullName", "apellido"], /nombre|name|apellido/i)
  };
}

function normalizeRole(raw) {
  const s = String(raw === undefined || raw === null ? "" : raw).toLowerCase();
  if (!s) return "admin";
  if (s.indexOf("admin") > -1 || s.indexOf("director") > -1 || s.indexOf("jefe") > -1) return "admin";
  if (s.indexOf("docen") > -1 || s.indexOf("profes") > -1 || s.indexOf("teacher") > -1) return "docente";
  if (s.indexOf("padre") > -1 || s.indexOf("parent") > -1 || s.indexOf("apoderad") > -1 || s.indexOf("madre") > -1) return "padre";
  return "admin";
}

function assignedPairs() {
  if (!USER) return [];
  return (DB.teacher_assignments || []).filter(a => String(a.userId) === String(USER.id));
}

function teacherScoped() {
  return USER && USER.rol === "docente";
}

function scopedGrados() {
  if (USER && USER.rol === "docente") {
    const pairs = assignedPairs();
    if (!pairs.length) return [];
    return pairs.map(p => p.gradoId).filter(v => v !== undefined && v !== "");
  }
  return null;
}

function scopedSecciones() {
  if (USER && USER.rol === "docente") {
    const pairs = assignedPairs();
    if (!pairs.length) return [];
    return pairs.map(p => p.seccionId).filter(v => v !== undefined && v !== "");
  }
  return null;
}

function inScope(kind, val) {
  const sc = kind === "grado" ? scopedGrados() : scopedSecciones();
  if (sc === null) return true;
  return sc.some(v => relMatchEquivalent(kind, v, val));
}

function relMatchEquivalent(kind, a, b) {
  if (a === undefined || a === null || b === undefined || b === null) return false;
  if (String(a) === String(b)) return true;
  const la = relLabel(kind, a), lb = relLabel(kind, b);
  if (la && String(la).trim() === String(b).trim()) return true;
  if (lb && String(lb).trim() === String(a).trim()) return true;
  if (la && lb && String(la).trim() === String(lb).trim()) return true;
  return false;
}

function childrenOf() {
  if (!USER) return null;
  const allStudents = DB.students || [];
  if (!allStudents.length) return [];

  // 1. Si el usuario padre tiene 'childrenIds' (como en la segunda imagen: id del estudiante)
  const myChildrenIds = new Set();
  if (USER.childrenIds) {
    String(USER.childrenIds).split(/[,;|\s]+/).forEach(id => { if (id.trim()) myChildrenIds.add(id.trim()); });
  }

  // Buscar también en DB.padres si el registro actual tiene childrenIds
  const pRecord = (DB.padres || []).find(p => String(p.id) === String(USER.id) || String(p.usuario || p.username).toLowerCase() === String(USER.usuario).toLowerCase());
  if (pRecord && pRecord.childrenIds) {
    String(pRecord.childrenIds).split(/[,;|\s]+/).forEach(id => { if (id.trim()) myChildrenIds.add(id.trim()); });
  }

  if (myChildrenIds.size > 0) {
    const matched = allStudents.filter(s => myChildrenIds.has(String(s.id).trim()));
    if (matched.length > 0) return matched;
  }

  // 2. Fallback: buscar en la columna padreId o equivalente en la hoja students
  const keys = unionKeys(allStudents, FALLBACK_HEADERS.students);
  const pf = keys.find(k => /padre|apoderado|parent|familia|madre/i.test(k));
  if (!pf) return myChildrenIds.size > 0 ? [] : null;

  const uk = userKeys();
  const ids = new Set([String(USER.id)]);
  if (uk.u && USER[uk.u]) ids.add(String(USER[uk.u]).toLowerCase());
  if (uk.n && USER[uk.n]) ids.add(String(USER[uk.n]).toLowerCase());
  (DB.padres || []).forEach(p => {
    const pu = uk.u ? p[uk.u] : null;
    const pn = uk.n ? p[uk.n] : null;
    if (pu && USER[uk.u] && String(pu).toLowerCase() === String(USER[uk.u]).toLowerCase()) ids.add(String(p.id));
    if (pn && USER[uk.n] && String(pn).toLowerCase() === String(USER[uk.n]).toLowerCase()) ids.add(String(p.id));
  });

  return allStudents.filter(s => {
    const raw = s[pf];
    if (raw === undefined || raw === null || String(raw).trim() === "") return false;
    return ids.has(String(raw)) || ids.has(String(raw).toLowerCase());
  });
}

function visibleStudents() {
  let list = DB.students || [];
  if (USER.rol === "docente") {
    const pairs = assignedPairs();
    list = list.filter(s => pairs.some(p => relMatch(s, "grado", p.gradoId) && relMatch(s, "seccion", p.seccionId)));
  }
  if (USER.rol === "padre") {
    const kids = childrenOf();
    if (kids) list = kids;
  }
  return list;
}

function filterGS(list, grado, seccion) {
  if (grado) list = list.filter(s => relMatch(s, "grado", grado));
  if (seccion) list = list.filter(s => relMatch(s, "seccion", seccion));
  return list;
}

function canEdit(view) {
  if (USER.rol === "padre") return false;
  if (USER.rol === "docente") return ["activities", "grades", "exams", "attendance", "materias"].indexOf(view) > -1;
  return USER.rol === "admin";
}

function hasAccess(view) {
  const item = NAV.find(n => n.view === view);
  return item ? item.roles.indexOf(USER.rol) > -1 : false;
}

function resetState() {
  state.q = {};
  state.students = { grado: "", seccion: "" };
  state.batch = {
    grades: { activity: "", grado: "", seccion: "", materia: "", bimestre: "" },
    exams: { activity: "", grado: "", seccion: "", materia: "", bimestre: "" },
    attendance: { fecha: "", grado: "", seccion: "", activity: "", materia: "", bimestre: "" }
  };
  state.assign = { userId: "", rows: undefined, forUser: undefined };
}

function showLogin(msg) {
  $("#app").hidden = true;
  $("#loginView").hidden = false;
  const e = $("#loginErr");
  const msgEl = $("#loginErrMsg");
  if (msg) { msgEl.textContent = msg; e.hidden = false; } else { e.hidden = true; }
  $("#cfgUrl").value = API_URL || "";
}

function showApp() {
  $("#loginView").hidden = true;
  $("#app").hidden = false;
  $("#userName").textContent = USER.nombre || USER.usuario || "Usuario";
  $("#userRole").textContent = pretty(USER.rol || "");
  $("#sideRole").textContent = pretty(USER.rol || "");
}

function buildNav() {
  const nav = $("#nav");
  nav.innerHTML = "";
  NAV.forEach(item => {
    if (item.sec) {
      if (item.roles.indexOf(USER.rol) > -1) {
        const d = document.createElement("div");
        d.className = "sec";
        d.textContent = item.sec;
        nav.appendChild(d);
      }
      return;
    }
    if (item.roles.indexOf(USER.rol) === -1) return;
    const b = document.createElement("button");
    b.type = "button";
    b.innerHTML = `<span>${esc(item.label)}</span>`;
    b.dataset.view = item.view;
    b.onclick = () => go(item.view);
    nav.appendChild(b);
  });
}

function go(view) {
  if (!hasAccess(view)) view = "dashboard";
  VIEW = view;
  $$("#nav button").forEach(b => b.classList.toggle("active", b.dataset.view === view));
  const item = NAV.find(n => n.view === view);
  $("#topTitle").textContent = item ? item.label : "";
  $("#sidebar").classList.remove("open");
  renderView();
}

function renderView() {
  const c = $("#content");
  if (VIEW === "dashboard") renderDashboard(c);
  else if (VIEW === "students") renderStudents(c);
  else if (VIEW === "assignments") renderAssignments(c);
  else if (VIEW === "grades") renderMatrixGrades(c);
  else if (BATCH[VIEW]) renderBatch(c, VIEW);
  else if (CRUD_CFG[VIEW]) renderCrud(c, CRUD_CFG[VIEW]);
  else go("dashboard");
}

function renderDashboard(c) {
  const role = USER.rol;
  const pairs = assignedPairs();
  const kids = role === "padre" ? childrenOf() : null;
  const visS = visibleStudents();
  const kidIds = visS.map(s => String(s.id));
  const t = today();
  let cards = [];

  if (role === "admin") {
    const todayAtt = (DB.attendance || []).filter(a => datePart(a.date) === t);
    cards = [
      [(DB.students || []).length, "Estudiantes Registrados", "👥"],
      [(DB.grados || []).length, "Grados Activos", "📚"],
      [(DB.secciones || []).length, "Secciones", "🏫"],
      [(DB.materias || []).length, "Materias", "📖"],
      [(DB.activities || []).length, "Actividades Creadas", "📝"],
      [(DB.users || []).length, "Cuentas de Usuario", "👤"],
      [(DB.grades || []).length, "Notas Asignadas", "🎯"],
      [todayAtt.length, "Asistencias Hoy", "📅"]
    ];
  } else if (role === "docente") {
    const myActs = scopeActivities() || [];
    const todayAtt = (DB.attendance || []).filter(a => datePart(a.date) === t && kidIds.indexOf(String(a.studentId)) > -1);
    cards = [
      [visS.length, "Mis Estudiantes", "👥"],
      [pairs.length, "Grados y Secciones Asignados", "🏫"],
      [scopeMaterias().length, "Mis Materias", "📖"],
      [myActs.length, "Mis Actividades", "📝"],
      [(DB.grades || []).filter(g => kidIds.indexOf(String(g.studentId)) > -1).length, "Notas Registradas", "🎯"],
      [(DB.exams || []).filter(e => kidIds.indexOf(String(e.studentId)) > -1).length, "Exámenes Calificados", "📊"],
      [todayAtt.length, "Asistencias Hoy", "📅"]
    ];
  } else {
    cards = [
      [kids ? kids.length : visS.length, "Mis Hijos", "👨‍👧‍👦"],
      [(DB.grades || []).filter(g => kidIds.indexOf(String(g.studentId)) > -1).length, "Calificaciones", "🎯"],
      [(DB.exams || []).filter(e => kidIds.indexOf(String(e.studentId)) > -1).length, "Exámenes Registrados", "📊"],
      [(DB.attendance || []).filter(a => kidIds.indexOf(String(a.studentId)) > -1).length, "Registros de Asistencia", "📅"]
    ];
  }

  let extra = "";
  if (role === "docente" && !pairs.length) {
    extra = '<div class="hint" style="margin-bottom:20px; border-color:var(--warning); color:#fde68a;">⚠️ Aún no tienes asignaciones de grado y sección registradas. Solicita tu asignación al Administrador en <b>Asignaciones Docentes</b> para registrar notas.</div>';
  }
  if (role === "docente" && pairs.length) {
    const gNames = pairs.map(p => relLabel("grado", p.gradoId) || p.gradoId);
    const sNames = pairs.map(p => relLabel("seccion", p.seccionId) || p.seccionId);
    const uniq = arr => arr.filter((v, i) => arr.indexOf(v) === i);
    extra = `<div class="hint" style="margin-bottom:20px;">🎓 Dictas en: <b>${esc(uniq(gNames).join(", "))}</b> — Sección(es): <b>${esc(uniq(sNames).join(", "))}</b></div>`;
  }
  if (role === "padre" && kids === null) {
    extra = '<div class="hint" style="margin-bottom:20px; border-color:var(--danger); color:#fca5a5;">⚠️ No se encontró un vínculo entre tu cuenta y los alumnos. El administrador debe registrar tu ID en la columna "padreId" de la hoja Estudiantes.</div>';
  }

  c.innerHTML =
    '<div class="page-head"><div><h1>Panel Principal</h1><p>Bienvenido(a) de nuevo, ' + esc(USER.nombre || USER.usuario || "") + '</p></div></div>' +
    extra +
    '<div class="stats">' + cards.map(x => `
      <div class="stat">
        <div class="stat-accent">${x[2]}</div>
        <div class="n">${x[0]}</div>
        <div class="l">${esc(x[1])}</div>
      </div>
    `).join("") + '</div>';
}

// 4. FUNCIONES DE EXPORTACIÓN (EXCEL Y REPORTE DE IMPRESIÓN)
function exportTableToExcel(tableId, fileName) {
  const table = document.getElementById(tableId);
  if (!table) { toast("No hay tabla para exportar", "err"); return; }
  
  if (typeof XLSX === "undefined") {
    toast("Librería Excel cargando... intenta nuevamente", "err");
    return;
  }

  const clone = table.cloneNode(true);
  clone.querySelectorAll("th.actions, td.actions").forEach(el => el.remove());

  const wb = XLSX.utils.table_to_book(clone, { sheet: "Datos" });
  XLSX.writeFile(wb, `${fileName}_${today()}.xlsx`);
  toast("Archivo Excel descargado con éxito", "ok");
}

function renderStudents(c) {
  c.innerHTML =
    '<div class="page-head">' +
      '<div><h1>Estudiantes</h1><p>Padrón de estudiantes registrados y matriculados</p></div>' +
      '<div class="head-actions">' +
        '<input class="search" id="sQ" placeholder="🔍 Buscar estudiante...">' +
        '<button class="btn success sm" id="sExport" type="button">📊 Exportar Excel</button>' +
        (canEdit("students") ? '<button class="btn primary" id="sAdd" type="button">➕ Nuevo Estudiante</button>' : "") +
      '</div>' +
    '</div>' +
    '<div class="controls">' +
      '<label class="ctl"><span>Grado</span><select class="in" id="sGrado"><option value="">Todos los grados</option>' + optionsHtml("grados", state.students.grado, "") + '</select></label>' +
      '<label class="ctl"><span>Sección</span><select class="in" id="sSeccion"><option value="">Todas las secciones</option>' + optionsHtml("secciones", state.students.seccion, "", state.students.grado) + '</select></label>' +
    '</div>' +
    '<div class="card"><div class="table-wrap" id="tableArea"></div></div>';

  $("#sQ").value = state.q.students || "";
  $("#sQ").addEventListener("input", e => { state.q.students = e.target.value; paintStudents(); });
  
  $("#sGrado").addEventListener("change", e => {
    state.students.grado = e.target.value;
    state.students.seccion = "";
    $("#sSeccion").innerHTML = '<option value="">Todas las secciones</option>' + optionsHtml("secciones", "", "", state.students.grado);
    paintStudents();
  });

  $("#sSeccion").addEventListener("change", e => {
    state.students.seccion = e.target.value;
    paintStudents();
  });

  const add = $("#sAdd");
  if (add) add.addEventListener("click", () => openRecordModal("students", null, paintStudents));
  
  $("#sExport").addEventListener("click", () => exportTableToExcel("studentsTable", "Estudiantes"));

  paintStudents();
}

function paintStudents() {
  const headers = unionKeys(DB.students, FALLBACK_HEADERS.students);
  let list = filterGS(visibleStudents(), state.students.grado, state.students.seccion);
  const q = (state.q.students || "").toLowerCase();
  if (q) list = list.filter(r => headers.some(h => String(r[h] === undefined ? "" : r[h]).toLowerCase().indexOf(q) > -1));
  const cols = headers.filter(h => h !== "id");
  const edit = canEdit("students");
  let html = "";

  if (!list.length) {
    html = '<div class="empty"><div class="empty-icon">👥</div>No se encontraron estudiantes para mostrar.</div>';
  } else {
    html = '<table id="studentsTable"><thead><tr>' + cols.map(h => "<th>" + esc(pretty(h)) + "</th>").join("") +
      (edit ? '<th class="actions" style="text-align:right">Acciones</th>' : "") + "</tr></thead><tbody>" +
      list.map((r, i) =>
        "<tr>" + cols.map(h => {
          return "<td>" + esc(displayValue(h, r[h])) + "</td>";
        }).join("") +
        (edit ? '<td class="actions"><button class="btn sm ghost" data-edit="' + i + '" type="button">Editar</button> <button class="btn sm danger" data-del="' + i + '" type="button">Eliminar</button></td>' : "") +
        "</tr>"
      ).join("") + "</tbody></table>";
  }

  const area = $("#tableArea");
  area.innerHTML = html;
  if (!list.length) return;
  area.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => openRecordModal("students", list[+b.dataset.edit], paintStudents)));
  area.querySelectorAll("[data-del]").forEach(b => b.addEventListener("click", () => deleteRecord("students", list[+b.dataset.del], paintStudents)));
}

function renderCrud(c, cfg) {
  const editable = canEdit(cfg.key);
  const canDel = !cfg.noDelete && editable;
  c.innerHTML =
    '<div class="page-head">' +
      '<div><h1>' + esc(cfg.title) + '</h1><p>' + esc(cfg.desc) + '</p></div>' +
      '<div class="head-actions">' +
        '<input class="search" id="cQ" placeholder="🔍 Buscar...">' +
        '<button class="btn success sm" id="cExport" type="button">📊 Exportar Excel</button>' +
        (editable ? '<button class="btn primary" id="cAdd" type="button">➕ ' + esc(cfg.add) + '</button>' : "") +
      '</div>' +
    '</div>' +
    '<div class="card"><div class="table-wrap" id="tableArea"></div></div>';

  $("#cQ").value = state.q[cfg.key] || "";
  $("#cQ").addEventListener("input", e => { state.q[cfg.key] = e.target.value; paintCrud(cfg, canDel); });
  const add = $("#cAdd");
  if (add) add.addEventListener("click", () => openRecordModal(cfg.key, null, () => paintCrud(cfg, canDel)));
  $("#cExport").addEventListener("click", () => exportTableToExcel("crudTable", cfg.title));

  paintCrud(cfg, canDel);
}

function paintCrud(cfg, canDel) {
  const headers = unionKeys(DB[cfg.key], FALLBACK_HEADERS[cfg.key]);
  let list = (DB[cfg.key] || []).slice();
  if (cfg.key === "activities" && teacherScoped()) list = list.filter(r => inScope("grado", relGet(r, "grado")));
  if (cfg.key === "materias" && teacherScoped()) list = list.filter(r => esMateriaDelDocente(r));
  const q = (state.q[cfg.key] || "").toLowerCase();
  if (q) list = list.filter(r => headers.some(h => String(r[h] === undefined ? "" : r[h]).toLowerCase().indexOf(q) > -1));
  const editable = canEdit(cfg.key === "activities" ? "activities" : "admin");
  const cols = headers.filter(h => h !== "id");
  let html = "";

  if (!list.length) {
    html = '<div class="empty"><div class="empty-icon">📁</div>No hay registros guardados en esta sección.</div>';
  } else {
    html = '<table id="crudTable"><thead><tr>' + cols.map(h => "<th>" + esc(pretty(h)) + "</th>").join("") +
      ((editable || canDel) ? '<th class="actions" style="text-align:right">Acciones</th>' : "") + "</tr></thead><tbody>" +
      list.map((r, i) =>
        "<tr>" + cols.map(h => {
          if (h === "rol" && r[h]) return '<td><span class="badge info">' + esc(r[h]) + '</span></td>';
          return "<td>" + esc(displayValue(h, r[h])) + "</td>";
        }).join("") +
        ((editable || canDel) ? '<td class="actions">' +
          (editable ? '<button class="btn sm ghost" data-edit="' + i + '" type="button">Editar</button> ' : "") +
          (canDel ? '<button class="btn sm danger" data-del="' + i + '" type="button">Eliminar</button>' : "") +
        '</td>' : "") +
        "</tr>"
      ).join("") + "</tbody></table>";
  }

  const area = $("#tableArea");
  area.innerHTML = html;
  if (!list.length) return;
  area.querySelectorAll("[data-edit]").forEach(b => b.addEventListener("click", () => openRecordModal(cfg.key, list[+b.dataset.edit], () => paintCrud(cfg, canDel))));
  area.querySelectorAll("[data-del]").forEach(b => b.addEventListener("click", () => deleteRecord(cfg.key, list[+b.dataset.del], () => paintCrud(cfg, canDel))));
}

function fieldHtml(h, val) {
  const v = val === undefined || val === null ? "" : val;
  const name = esc(h);
  const kl = h.toLowerCase();
  let inner;
  if (kl === "id") {
    return '<input type="hidden" name="id" value="' + esc(v) + '">';
  }
  if (/grado|grade/i.test(h) && !/seccion|secci|section/i.test(h)) {
    inner = '<select class="in" name="' + name + '">' + optionsHtml("grados", v) + '</select>';
  } else if (/seccion|secci|section|aula/i.test(h)) {
    inner = '<select class="in" name="' + name + '">' + optionsHtml("secciones", v) + '</select>';
  } else if (/materia|subject|course|asignatura/i.test(h)) {
    inner = '<select class="in" name="' + name + '">' + optionsHtml("materias", v) + '</select>';
  } else if (/bimestre|term|periodo/i.test(h)) {
    inner = '<select class="in" name="' + name + '">' + optionsHtml("bimestres", v) + '</select>';
  } else if (/^(rol|role|tipo)$/i.test(kl)) {
    inner = '<select class="in" name="' + name + '">' +
      '<option value="admin"' + (v === "admin" ? ' selected' : '') + '>Administrador</option>' +
      '<option value="docente"' + (v === "docente" ? ' selected' : '') + '>Docente</option>' +
      '<option value="padre"' + (v === "padre" ? ' selected' : '') + '>Padre de Familia</option>' +
      '</select>';
  } else if (/contrase|password|clave|pwd/i.test(kl)) {
    inner = '<input class="in" type="password" name="' + name + '" value="' + esc(v) + '" placeholder="Nueva contraseña">';
  } else if (/fecha|date|nacimiento|inicio|fin/i.test(kl)) {
    inner = '<input class="in" type="date" name="' + name + '" value="' + esc(datePart(v)) + '">';
  } else if (/nota|puntaje|score|valor|calific|promedio|edad|dni|puntos|telefono|fono|codigo|puntajemax/i.test(kl)) {
    inner = '<input class="in" type="number" step="any" name="' + name + '" value="' + esc(v) + '">';
  } else if (/observ|descripcion|detalle|direccion|comentario/i.test(kl)) {
    inner = '<textarea class="in" name="' + name + '" rows="3">' + esc(v) + '</textarea>';
  } else {
    inner = '<input class="in" name="' + name + '" value="' + esc(v) + '">';
  }
  const full = /observ|descripcion|detalle|direccion|comentario/i.test(h) ? " full" : "";
  return '<label class="fld' + full + '"><span>' + esc(pretty(h)) + '</span>' + inner + '</label>';
}

function openModal(title, bodyHtml, footHtml, setup) {
  const ov = document.createElement("div");
  ov.className = "modal-ov";
  ov.innerHTML =
    '<div class="modal">' +
      '<div class="mh"><h3>' + esc(title) + '</h3><button class="x" type="button" data-close>&times;</button></div>' +
      '<div class="mb">' + bodyHtml + '</div>' +
      '<div class="mf">' + footHtml + '</div>' +
    '</div>';
  document.body.appendChild(ov);
  ov.addEventListener("click", e => { if (e.target === ov) ov.remove(); });
  ov.querySelector("[data-close]").addEventListener("click", () => ov.remove());
  if (setup) setup(ov);
  return ov;
}

function docenteDefaults(key) {
  if (!(USER && USER.rol === "docente")) return {};
  const pairs = assignedPairs();
  if (!pairs.length) return {};
  const gs = scopedGrados();
  const ss = scopedSecciones();
  const out = {};
  if (gs.length === 1) out.gradoId = String(gs[0]);
  if (ss.length === 1) out.seccionId = String(ss[0]);
  return out;
}

function openRecordModal(key, record, after) {
  const headers = unionKeys(DB[key], FALLBACK_HEADERS[key]);
  const isNew = !record || !record.id;
  const def = isNew ? docenteDefaults(key) : {};
  const rec = Object.assign({}, def, record || {});
  let body = '<form id="recForm"><div class="form-grid">' +
    headers.map(h => fieldHtml(h, rec[h])).join("") +
    '</div></form>';
  if (isNew) body = body.replace('<form id="recForm">', '<form id="recForm"><input type="hidden" name="id" value="' + uid() + '">');
  const foot = '<button class="btn ghost" type="button" data-cancel>Cancelar</button> <button class="btn primary" type="button" data-save>Guardar</button>';
  
  openModal((isNew ? "Nuevo " : "Editar ") + pretty(key), body, foot, ov => {
    ov.querySelector("[data-cancel]").addEventListener("click", () => ov.remove());
    ov.querySelector("[data-save]").addEventListener("click", async () => {
      const form = ov.querySelector("#recForm");
      const out = {};
      form.querySelectorAll("[name]").forEach(el => {
        let v = el.value;
        if (el.type === "number") v = v === "" ? "" : Number(v);
        out[el.name] = v;
      });
      if (!out.id) out.id = uid();
      if (key === "materias" && USER && USER.rol === "docente") {
        if ((out.gradoId === undefined || out.gradoId === "") && def.gradoId) out.gradoId = def.gradoId;
      }
      try {
        setLoading(true, "Guardando registro...");
        const r = await apiPost({ action: SAVE_ACTIONS[key], payload: out });
        setLoading(false);
        if (r && (r.error || r.success === false)) { toast(r.error || r.message || "Error al guardar", "err"); return; }
        ov.remove();
        toast("Registro guardado correctamente", "ok");
        await reload();
        if (after) after();
      } catch (err) {
        setLoading(false);
        toast("Error: " + err.message, "err");
      }
    });
  });
}

async function deleteRecord(key, rec, after) {
  if (!confirm("¿Estás seguro de que deseas eliminar este registro permanentemente?")) return;
  try {
    setLoading(true, "Eliminando registro...");
    let r;
    if (key === "students") r = await apiPost({ action: "delete_student", id: rec.id });
    else r = await apiPost({ action: DELETE_ACTIONS[key], payload: { id: rec.id } });
    setLoading(false);
    if (r && (r.error || r.success === false)) { toast(r.error || r.message || "Error al eliminar", "err"); return; }
    toast("Registro eliminado con éxito", "ok");
    await reload();
    if (after) after();
  } catch (err) {
    setLoading(false);
    toast("Error: " + err.message, "err");
  }
}

function colType(key, rows) {
  if (/estado|asistencia|status|presente|observacion/i.test(key)) return "select";
  if (/fecha|date|nacimiento|inicio|fin/i.test(key)) return "date";
  if (/nota|puntaje|score|valor|calific|promedio|total|puntos/i.test(key)) return "number";
  const vals = (rows || []).map(r => r[key]).filter(v => v !== "" && v !== undefined && v !== null);
  if (vals.length && vals.every(v => !isNaN(Number(v)))) return "number";
  return "text";
}

function batchColumns(vk) {
  const cfg = BATCH[vk];
  const recs = DB[cfg.key] || [];
  const keys = unionKeys(recs, []);
  const cols = keys.filter(k => cfg.system.indexOf(k) === -1);
  if (!cols.length) return cfg.fallback.map(f => ({ key: f.key, label: f.label, type: f.type, options: f.options, candidates: f.candidates }));
  return cols.map(k => ({ key: k, label: pretty(k), type: colType(k, recs) }));
}

function findExisting(vk, st, studentId) {
  const cfg = BATCH[vk];
  const recs = DB[cfg.key] || [];
  return recs.find(r => {
    if (String(r.studentId) !== String(studentId)) return false;
    if (vk === "grades") return String(r.activityId) === String(st.activity);
    if (vk === "exams") return String(r.materiaId) === String(st.materia) && String(r.bimestreId) === String(st.bimestre);
    if (vk === "attendance") return datePart(r.date) === st.fecha;
    return false;
  });
}

function autoFromActivity(st) {
  const act = (DB.activities || []).find(a => String(a.id) === String(st.activity));
  if (!act) return;
  const gf = relField(act, "grado"); if (gf && act[gf]) st.grado = String(act[gf]);
  const mf = relField(act, "materia"); if (mf && act[mf]) st.materia = String(act[mf]);
  const bf = relField(act, "bimestre"); if (bf && act[bf]) st.bimestre = String(act[bf]);
}

function batchStudents(vk) {
  const st = state.batch[vk];
  let list = visibleStudents();
  if (teacherScoped() && st.grado && !inScope("grado", st.grado)) st.grado = "";
  if (teacherScoped() && st.seccion && !inScope("seccion", st.seccion)) st.seccion = "";
  return filterGS(list, st.grado, st.seccion);
}

function cellHtml(vk, col, student, existing, editable) {
  const cfg = BATCH[vk];
  let val = "";
  if (existing) {
    if (existing[col.key] !== undefined && existing[col.key] !== "") val = existing[col.key];
    else if (col.candidates) {
      for (const c of col.candidates) { if (existing[c] !== undefined && existing[c] !== "") { val = existing[c]; break; } }
    }
  }
  if ((val === "" || val === undefined) && cfg.relFill && cfg.relFill[col.key]) {
    const sv = state.batch[vk][cfg.relFill[col.key]];
    if (sv) val = sv;
  }
  const dis = editable ? "" : " disabled";
  const common = ' data-stud="' + esc(student.id) + '" data-col="' + esc(col.key) + '"';
  if (col.type === "select") {
    const defaults = (col.options || ["Presente", "Ausente", "Tardanza", "Justificado"]);
    const recs = DB[cfg.key] || [];
    const existingVals = [];
    recs.forEach(r => { const v = r[col.key]; if (v !== undefined && v !== "" && existingVals.indexOf(String(v)) === -1) existingVals.push(String(v)); });
    const opts = [];
    defaults.forEach(d => { if (opts.indexOf(d) === -1) opts.push(d); });
    existingVals.forEach(d => { if (opts.indexOf(d) === -1) opts.push(d); });
    return '<select class="in"' + common + dis + '><option value="">—</option>' +
      opts.map(o => '<option value="' + esc(o) + '"' + (String(val) === o ? ' selected' : '') + '>' + esc(o) + '</option>').join("") + '</select>';
  }
  if (col.type === "number") return '<input class="in" type="number" step="any"' + common + dis + ' value="' + esc(val) + '">';
  if (col.type === "date") return '<input class="in" type="date"' + common + dis + ' value="' + esc(datePart(val)) + '">';
  return '<input class="in"' + common + dis + ' value="' + esc(val) + '">';
}

function renderBatch(c, vk) {
  const cfg = BATCH[vk];
  const st = state.batch[vk];
  if (vk === "attendance" && !st.fecha) st.fecha = today();
  if (teacherScoped()) {
    const gs = scopedGrados();
    if (gs.length === 1 && !st.grado) st.grado = String(gs[0]);
    const ss = scopedSecciones();
    if (ss.length === 1 && !st.seccion) st.seccion = String(ss[0]);
  }
  const editable = canEdit(vk);
  const missing = cfg.required.filter(k => !st[k]);
  const anyOptions = cfg.controls.filter(x => x.db).every(x => (DB[x.db] || []).length > 0);

  let controls = '<div class="controls">';
  cfg.controls.forEach(ct => {
    if (ct.type === "date") {
      controls += '<label class="ctl"><span>' + esc(ct.label) + '</span><input class="in" type="date" data-ctl="' + ct.k + '" value="' + esc(st[ct.k] || "") + '"></label>';
    } else {
      const ph = (ct.k === "grado" || ct.k === "seccion") ? "Todos" : "— Seleccionar —";
      const fG = ct.k === "seccion" ? st.grado : null;
      controls += '<label class="ctl"><span>' + esc(ct.label) + '</span><select class="in" data-ctl="' + ct.k + '">' + optionsHtml(ct.db, st[ct.k], ph, fG) + '</select></label>';
    }
  });
  controls += '</div>';

  let body;
  if (!anyOptions) {
    body = '<div class="card"><div class="empty"><div class="empty-icon">⚠️</div>Primero debes registrar los datos de configuración (grados, secciones, materias o bimestres).</div></div>';
  } else if (missing.length) {
    body = '<div class="card"><div class="empty"><div class="empty-icon">🔍</div>Selecciona ' + missing.map(k => {
      const ct = cfg.controls.find(x => x.k === k);
      return ct ? ct.label.toLowerCase() : k;
    }).join(" y ") + ' para cargar los estudiantes.</div></div>';
  } else {
    const cols = batchColumns(vk);
    const students = batchStudents(vk);
    if (!students.length) {
      body = '<div class="card"><div class="empty"><div class="empty-icon">👥</div>No hay estudiantes para mostrar con los filtros seleccionados.</div></div>';
    } else {
      body = '<div class="card"><div class="table-wrap"><table id="batchTable"><thead><tr><th>Estudiante</th>' +
        cols.map(col => "<th>" + esc(col.label) + "</th>").join("") +
        '</tr></thead><tbody>' +
        students.map(s => {
          const ex = findExisting(vk, st, s.id);
          return "<tr><td><b>" + esc(labelOf(s)) + "</b></td>" +
            cols.map(col => "<td>" + cellHtml(vk, col, s, ex, editable) + "</td>").join("") + "</tr>";
        }).join("") +
        '</tbody></table></div></div>';
    }
  }

  c.innerHTML =
    '<div class="page-head">' +
      '<div><h1>' + esc(cfg.title) + '</h1><p>' + esc(cfg.desc) + '</p></div>' +
      '<div class="head-actions">' +
        (!missing.length && anyOptions ? '<button class="btn success sm" id="bExport" type="button">📊 Exportar Excel</button>' : "") +
        (editable && !missing.length && anyOptions ?
          '<button class="btn primary" id="bSave" type="button">💾 Guardar Cambios</button>' : "") +
      '</div>' +
    '</div>' + controls + body +
    (!editable ? '<div class="hint" style="margin-top:14px">🔒 Tu usuario tiene acceso de sólo lectura en esta sección.</div>' : "");

  c.querySelectorAll("[data-ctl]").forEach(el => {
    el.addEventListener("change", e => {
      const k = e.target.dataset.ctl;
      st[k] = e.target.value;
      if (vk === "grades" && k === "activity") autoFromActivity(st);
      renderView();
    });
  });

// ==========================================
// REGISTRO DE NOTAS (ZONA) - MÓDULO MATRIZ TIPO IMAGEN
// ==========================================
function renderMatrixGrades(c) {
  const mg = state.matrixGrades;
  const bimestres = DB.bimestres || [];
  const grados = DB.grados || [];
  const secciones = DB.secciones || [];
  const materias = scopeMaterias();

  // Valores predeterminados inteligentes
  if (!mg.bimestre && bimestres.length) mg.bimestre = String(bimestres[0].id);
  if (!mg.grado && grados.length) {
    if (teacherScoped()) {
      const gs = scopedGrados();
      if (gs && gs.length) mg.grado = String(gs[0]);
      else mg.grado = String(grados[0].id);
    } else {
      mg.grado = String(grados[0].id);
    }
  }
  if (!mg.seccion && secciones.length) {
    if (teacherScoped()) {
      const ss = scopedSecciones();
      if (ss && ss.length) mg.seccion = String(ss[0]);
    }
  }
  if (!mg.materia && materias.length) mg.materia = String(materias[0].id);

  const editable = canEdit("grades");

  // Filtrar actividades por Grado, Materia y Bimestre seleccionados
  const allActs = DB.activities || [];
  const activities = allActs.filter(a => {
    const matchB = !mg.bimestre || relMatch(a, "bimestre", mg.bimestre);
    const matchG = !mg.grado || relMatch(a, "grado", mg.grado);
    const matchM = !mg.materia || relMatch(a, "materia", mg.materia);
    return matchB && matchG && matchM;
  });

  // Filtrar estudiantes por Grado y Sección
  let students = filterGS(visibleStudents(), mg.grado, mg.seccion);
  if (mg.q) {
    const query = mg.q.toLowerCase().trim();
    students = students.filter(s => {
      const name = String(s.name || s.nombre || "").toLowerCase();
      const code = String(s.code || s.codigo || s.id || "").toLowerCase();
      return name.includes(query) || code.includes(query);
    });
  }

  // Mapa de notas registradas en DB.grades: clave "studentId_activityId"
  const gradesMap = {};
  (DB.grades || []).forEach(g => {
    const sId = String(g.studentId || "");
    const aId = String(g.activityId || "");
    const val = g.nota !== undefined ? g.nota : (g.score !== undefined ? g.score : (g.puntaje !== undefined ? g.puntaje : g.valor));
    gradesMap[sId + "_" + aId] = { id: g.id, val: val !== undefined && val !== null ? val : "" };
  });

  // Encabezado de página idéntico a la imagen
  let html = `
    <div class="page-head" style="align-items:center; margin-bottom:20px;">
      <div>
        <h1 style="font-size:24px; font-weight:800; letter-spacing:-0.02em;">Registro de Notas (Zona)</h1>
      </div>
      <div class="head-actions">
        ${editable ? '<button class="btn primary" id="btnSaveMatrix" type="button" style="padding:10px 22px; font-size:14px; font-weight:700;"><svg width="18" height="18" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M8 7H5a2 2 0 00-2 2v9a2 2 0 002 2h14a2 2 0 002-2V9a2 2 0 00-2-2h-3m-1 4l-3 3m0 0l-3-3m3 3V4"/></svg> Guardar cambios</button>' : ''}
      </div>
    </div>

    <!-- Barra de Controles y Filtros Superior -->
    <div class="card" style="padding:20px 24px; margin-bottom:20px;">
      <div style="display:grid; grid-template-columns: repeat(auto-fit, minmax(170px, 1fr)); gap:16px; align-items:flex-end;">
        <label class="ctl">
          <span>Bimestre</span>
          <select class="in" id="mBimestre">
            ${optionsHtml("bimestres", mg.bimestre, "— Seleccionar —")}
          </select>
        </label>
        <label class="ctl">
          <span>Grado</span>
          <select class="in" id="mGrado">
            ${optionsHtml("grados", mg.grado, "— Seleccionar —")}
          </select>
        </label>
        <label class="ctl">
          <span>Sección</span>
          <select class="in" id="mSeccion">
            <option value="">Todas las secciones</option>
            ${optionsHtml("secciones", mg.seccion, "", mg.grado, true)}
          </select>
        </label>
        <label class="ctl">
          <span>Materia</span>
          <select class="in" id="mMateria">
            ${optionsHtml("materias", mg.materia, "— Seleccionar —")}
          </select>
        </label>
        <label class="ctl">
          <span>🔍 Buscar Alumno</span>
          <input class="in" id="mSearch" placeholder="Escriba el nombre..." value="${esc(mg.q || '')}">
        </label>
      </div>

      <!-- Botones de Acción Secundarios: Excel y PDF/Imprimir -->
      <div style="display:flex; gap:12px; margin-top:18px;">
        <button class="btn ghost sm" id="btnExportMatrixExcel" type="button" style="padding:8px 16px; font-weight:600;">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M12 10v6m0 0l-3-3m3 3l3-3m2 8H7a2 2 0 01-2-2V5a2 2 0 012-2h5.586a1 1 0 01.707.293l5.414 5.414a1 1 0 01.293.707V19a2 2 0 01-2 2z"/></svg>
          Excel
        </button>
        <button class="btn ghost sm" id="btnPrintMatrix" type="button" style="padding:8px 16px; font-weight:600;">
          <svg width="16" height="16" fill="none" stroke="currentColor" stroke-width="2" viewBox="0 0 24 24"><path stroke-linecap="round" stroke-linejoin="round" d="M17 17h2a2 2 0 002-2v-4a2 2 0 00-2-2H5a2 2 0 00-2 2v4a2 2 0 002 2h2m2 4h6a2 2 0 002-2v-4H7v4a2 2 0 002 2zm8-12V5a2 2 0 00-2-2H9a2 2 0 00-2 2v4h10z"/></svg>
          PDF / Imprimir
        </button>
      </div>
    </div>
  `;

  // Cuerpo de la Tabla Matriz
  if (!students.length) {
    html += '<div class="card"><div class="empty"><div class="empty-icon">👥</div>No hay alumnos para mostrar con los filtros seleccionados.</div></div>';
  } else if (!activities.length) {
    html += '<div class="card"><div class="empty"><div class="empty-icon">📝</div>No hay actividades registradas para este grado, materia y bimestre.<br><small style="color:var(--text-dim); margin-top:8px; display:block;">Puedes registrar nuevas tareas o evaluaciones en el menú <b>Actividades</b>.</small></div></div>';
  } else {
    html += `
      <div class="card">
        <div class="table-wrap">
          <table id="matrixGradesTable" style="min-width:680px;">
            <thead>
              <tr>
                <th style="min-width:240px; padding:16px;">ALUMNO</th>
                ${activities.map(act => {
                  const pts = act.puntajeMax || act.puntaje || act.puntos || act.max || 10;
                  const title = esc(act.nombre || act.name || act.titulo || "Actividad");
                  return `
                    <th style="text-align:center; min-width:140px; padding:12px 14px;">
                      <div style="font-size:11px; color:var(--primary); font-weight:800; letter-spacing:0.04em;">[${pts}PTS]</div>
                      <div style="font-size:12px; font-weight:700; color:var(--text); text-transform:uppercase; margin-top:2px;" title="${title}">${title}</div>
                    </th>
                  `;
                }).join("")}
                <th style="text-align:center; min-width:120px; background:#0c162a; color:#93c5fd; font-weight:800; padding:16px;">ZONA TOTAL</th>
              </tr>
            </thead>
            <tbody>
              ${students.map(s => {
                const sName = esc(s.name || s.nombre || "Alumno");
                const sSec = esc(s.section || s.seccion || relLabel("seccion", s.seccionId) || "");
                let rowTotal = 0;

                const cells = activities.map(act => {
                  const rec = gradesMap[String(s.id) + "_" + String(act.id)] || { id: null, val: "" };
                  const vNum = parseFloat(rec.val);
                  if (!isNaN(vNum)) rowTotal += vNum;

                  const dis = editable ? "" : " disabled";
                  return `
                    <td style="text-align:center; padding:10px 14px;">
                      <input 
                        class="in matrix-grade-input" 
                        type="number" 
                        step="any" 
                        min="0"
                        max="${act.puntajeMax || 100}"
                        data-student="${esc(s.id)}" 
                        data-activity="${esc(act.id)}"
                        data-grade-id="${rec.id ? esc(rec.id) : ''}"
                        value="${rec.val !== '' ? esc(rec.val) : ''}"
                        style="max-width:85px; text-align:center; font-weight:700; font-size:14px; margin:0 auto;"
                        ${dis}
                      >
                    </td>
                  `;
                }).join("");

                return `
                  <tr>
                    <td style="padding:14px 18px;">
                      <div style="font-weight:700; font-size:14px; color:#fff;">${sName}</div>
                      ${sSec ? `<div style="font-size:11.5px; color:var(--text-dim); margin-top:2px;">Sección ${sSec}</div>` : ''}
                    </td>
                    ${cells}
                    <td style="text-align:center; font-weight:800; font-size:16px; color:#60a5fa; background:rgba(59,130,246,0.06);" id="total_${esc(s.id)}">
                      ${rowTotal}
                    </td>
                  </tr>
                `;
              }).join("")}
            </tbody>
          </table>
        </div>
      </div>
    `;
  }

  c.innerHTML = html;

  // Event Listeners para cambios de filtros
  $("#mBimestre").addEventListener("change", e => { mg.bimestre = e.target.value; renderView(); });
  $("#mGrado").addEventListener("change", e => { mg.grado = e.target.value; renderView(); });
  $("#mSeccion").addEventListener("change", e => { mg.seccion = e.target.value; renderView(); });
  $("#mMateria").addEventListener("change", e => { mg.materia = e.target.value; renderView(); });

  const sInput = $("#mSearch");
  sInput.addEventListener("input", e => {
    mg.q = e.target.value;
    // Filtrado en vivo de filas
    const term = mg.q.toLowerCase().trim();
    $$("#matrixGradesTable tbody tr").forEach(tr => {
      const text = tr.innerText.toLowerCase();
      tr.style.display = text.includes(term) ? "" : "none";
    });
  });

  // Cálculo en vivo de Zona Total al escribir nota
  c.querySelectorAll(".matrix-grade-input").forEach(inp => {
    inp.addEventListener("input", () => {
      const sId = inp.dataset.student;
      let total = 0;
      $$(`.matrix-grade-input[data-student="${sId}"]`).forEach(el => {
        const val = parseFloat(el.value);
        if (!isNaN(val)) total += val;
      });
      const totEl = document.getElementById("total_" + sId);
      if (totEl) totEl.textContent = Math.round(total * 100) / 100;
    });
  });

  // Botón Guardar Cambios
  const btnSave = $("#btnSaveMatrix");
  if (btnSave) btnSave.addEventListener("click", () => saveMatrixGrades(activities));

  // Botón Exportar Excel
  $("#btnExportMatrixExcel").addEventListener("click", () => {
    exportTableToExcel("matrixGradesTable", "Registro_Notas_Zona");
  });

  // Botón Imprimir
  $("#btnPrintMatrix").addEventListener("click", () => {
    window.print();
  });
}

// Guardar Notas de la Matriz en Google Sheets
async function saveMatrixGrades(activities) {
  const inputs = $$(".matrix-grade-input");
  if (!inputs.length) {
    toast("No hay calificaciones para guardar", "err");
    return;
  }

  const payload = [];
  inputs.forEach(inp => {
    const sId = inp.dataset.student;
    const aId = inp.dataset.activity;
    const gId = inp.dataset.gradeId;
    const val = inp.value.trim();

    // Solo enviamos si tiene nota o si ya existía y fue modificada
    if (val !== "" || gId) {
      payload.push({
        id: gId || uid(),
        studentId: sId,
        activityId: aId,
        nota: val !== "" ? Number(val) : ""
      });
    }
  });

  if (!payload.length) {
    toast("No has ingresado calificaciones aún", "err");
    return;
  }

  try {
    setLoading(true, "Guardando registro de notas...");
    const r = await apiPost({ action: "save_grades", payload: payload });
    setLoading(false);

    if (r && (r.error || r.success === false)) {
      toast(r.error || r.message || "Error al guardar calificaciones", "err");
      return;
    }

    toast("¡Calificaciones guardadas exitosamente!", "ok");
    await reload();
  } catch (err) {
    setLoading(false);
    toast("Error al guardar: " + err.message, "err");
  }
}


async function saveBatch(vk) {
  const cfg = BATCH[vk];
  const st = state.batch[vk];
  const cols = batchColumns(vk);
  const students = batchStudents(vk);
  const vals = {};
  $$("#batchTable [data-stud]").forEach(el => {
    const sid = el.dataset.stud;
    const ck = el.dataset.col;
    if (!vals[sid]) vals[sid] = {};
    let v = el.value;
    if (el.type === "number" && v !== "") v = Number(v);
    vals[sid][ck] = v;
  });
  const items = [];
  students.forEach(s => {
    const sv = vals[String(s.id)] || {};
    const ex = findExisting(vk, st, s.id);
    const allEmpty = cols.every(col => {
      const v = sv[col.key];
      return v === undefined || v === "" || v === null;
    });
    if (!ex && allEmpty) return;
    const item = {};
    if (ex && ex.id) item.id = ex.id;
    item.studentId = s.id;
    if (vk === "grades") item.activityId = st.activity;
    if (vk === "exams") { item.materiaId = st.materia; item.bimestreId = st.bimestre; }
    if (vk === "attendance") item.date = st.fecha;
    cols.forEach(col => {
      const v = sv[col.key];
      if (col.candidates && cfg.fallback.some(f => f.key === col.key && f.candidates)) {
        const cand = (cfg.fallback.find(f => f.key === col.key) || {}).candidates || [col.key];
        cand.forEach(cc => { item[cc] = v === undefined ? "" : v; });
      } else {
        item[col.key] = v === undefined ? "" : v;
      }
    });
    if (cfg.relFill) {
      Object.keys(cfg.relFill).forEach(target => {
        const sv2 = st[cfg.relFill[target]];
        if (sv2 && (item[target] === undefined || item[target] === "")) item[target] = sv2;
      });
    }
    items.push(item);
  });
  if (!items.length) { toast("No hay cambios que guardar", "err"); return; }
  try {
    setLoading(true, "Guardando lote de calificaciones...");
    const r = await apiPost({ action: "save_" + cfg.key, payload: items });
    setLoading(false);
    if (r && (r.error || r.success === false)) { toast(r.error || r.message || "Error al guardar", "err"); return; }
    toast(items.length + " registro(s) guardado(s) exitosamente", "ok");
    await reload();
  } catch (err) {
    setLoading(false);
    toast("Error: " + err.message, "err");
  }
}

function renderAssignments(c) {
  const users = DB.users || [];
  if (!state.assign.userId && users.length) state.assign.userId = String(users[0].id);
  const cur = state.assign.userId;
  if (state.assign.rows === undefined || state.assign.forUser !== cur) {
    state.assign.forUser = cur;
    state.assign.rows = (DB.teacher_assignments || [])
      .filter(a => String(a.userId) === String(cur))
      .map(p => ({ gradoId: String(p.gradoId === undefined ? "" : p.gradoId), seccionId: String(p.seccionId === undefined ? "" : p.seccionId) }));
  }
  const rows = state.assign.rows;

  c.innerHTML =
    '<div class="page-head">' +
      '<div><h1>Asignaciones Docentes</h1><p>Vincular a cada profesor con sus grados y secciones</p></div>' +
      '<div class="head-actions"><button class="btn primary" id="aSave" type="button">💾 Guardar Asignaciones</button></div>' +
    '</div>' +
    '<div class="controls">' +
      '<label class="ctl grow"><span>Docente</span><select class="in" id="aUser">' + optionsHtml("users", cur, "— Seleccionar Docente —") + '</select></label>' +
    '</div>' +
    '<div class="card" style="margin-bottom:18px"><div class="pad">' +
      '<div class="toolbar"><b>Carga Académica del Docente</b><button class="btn ghost sm" id="aAdd" type="button">➕ Agregar Asignación</button></div>' +
      '<div class="rowlist" id="aRows">' +
        (rows.length ? "" : '<div class="muted">Este docente no tiene grados ni secciones asignadas aún.</div>') +
        rows.map((r, i) =>
          '<div class="rowitem">' +
            '<select class="in" data-arow="' + i + '" data-af="gradoId">' + optionsHtml("grados", r.gradoId, "Grado") + '</select>' +
            '<select class="in" data-arow="' + i + '" data-af="seccionId">' + optionsHtml("secciones", r.seccionId, "Sección", r.gradoId, true) + '</select>' +
            '<button class="btn sm danger" data-adel="' + i + '" type="button">Quitar</button>' +
          '</div>'
        ).join("") +
      '</div>' +
    '</div></div>' +
    '<div class="card"><div class="table-wrap">' + assignmentsTable() + '</div></div>';

  $("#aUser").addEventListener("change", e => { state.assign.userId = e.target.value; state.assign.rows = undefined; renderView(); });
  $("#aAdd").addEventListener("click", () => { rows.push({ gradoId: "", seccionId: "" }); renderView(); });
  $("#aSave").addEventListener("click", async () => {
    if (!state.assign.userId) { toast("Selecciona un docente primero", "err"); return; }
    const payload = rows.filter(r => r.gradoId || r.seccionId);
    try {
      setLoading(true, "Guardando asignaciones...");
      const r = await apiPost({ action: "save_assignments", payload: { userId: state.assign.userId, assignments: payload } });
      setLoading(false);
      if (r && (r.error || r.success === false)) { toast(r.error || r.message || "Error al guardar", "err"); return; }
      toast("Asignaciones actualizadas correctamente", "ok");
      await reload();
      state.assign.rows = undefined;
      renderView();
    } catch (err) {
      setLoading(false);
      toast("Error: " + err.message, "err");
    }
  });

  c.querySelectorAll("[data-arow]").forEach(el => {
    el.addEventListener("change", e => {
      const idx = +e.target.dataset.arow;
      const f = e.target.dataset.af;
      rows[idx][f] = e.target.value;
      if (f === "gradoId") {
        rows[idx].seccionId = "";
        renderView();
      }
    });
  });

  c.querySelectorAll("[data-adel]").forEach(b => b.addEventListener("click", () => { rows.splice(+b.dataset.adel, 1); renderView(); }));
}

function assignmentsTable() {
  const list = DB.teacher_assignments || [];
  if (!list.length) return '<div class="empty"><div class="empty-icon">🔗</div>No hay asignaciones registradas en el sistema.</div>';
  const userLabel = id => {
    const u = (DB.users || []).find(x => String(x.id) === String(id));
    return u ? labelOf(u) : String(id);
  };
  return '<table><thead><tr><th>Docente</th><th>Grado</th><th>Sección</th></tr></thead><tbody>' +
    list.map(a => {
      const g = (DB.grados || []).find(x => String(x.id) === String(a.gradoId));
      const s = (DB.secciones || []).find(x => String(x.id) === String(a.seccionId));
      return '<tr><td><b>' + esc(userLabel(a.userId)) + '</b></td><td>' + esc(g ? labelOf(g) : a.gradoId) + '</td><td>' + esc(s ? labelOf(s) : a.seccionId) + '</td></tr>';
    }).join("") + '</tbody></table>';
}

function openConnModal() {
  const body =
    '<label class="fld"><span>URL del Web App (Google Apps Script)</span><input class="in" id="connUrl" placeholder="https://script.google.com/macros/s/.../exec" value="' + esc(API_URL) + '"></label>' +
    '<div class="hint"><b>¿Cómo obtener la URL?</b><br>1. En tu editor de Google Apps Script, haz clic en Implementar &gt; Nueva implementación.<br>2. Elige tipo: Aplicación web.<br>3. Ejecutar como: Yo. Acceso: Cualquier usuario.<br>4. Pega aquí el enlace de la aplicación web.</div>';
  const foot = '<button class="btn ghost" type="button" data-test>Probar Conexión</button> <button class="btn primary" type="button" data-save>Guardar</button>';
  
  openModal("Conexión con el Backend", body, foot, ov => {
    ov.querySelector("[data-test]").addEventListener("click", async () => {
      const v = ov.querySelector("#connUrl").value.trim();
      if (!v) { toast("Ingresa la URL primero", "err"); return; }
      const old = API_URL; API_URL = v;
      try {
        setLoading(true, "Probando conexión...");
        await apiGet("get_data");
        setLoading(false);
        toast("¡Conexión exitosa con el backend!", "ok");
      } catch (err) {
        setLoading(false);
        API_URL = old;
        toast("Falló la conexión: " + err.message, "err");
      }
    });

    ov.querySelector("[data-save]").addEventListener("click", () => {
      const v = ov.querySelector("#connUrl").value.trim();
      API_URL = v;
      localStorage.setItem(LS_URL, v);
      ov.remove();
      toast("URL guardada exitosamente", "ok");
    });
  });
}

async function bootApp() {
  try {
    setLoading(true, "Cargando datos escolares...");
    await loadData();
    setLoading(false);
    showApp();
    buildNav();
    go("dashboard");
  } catch (err) {
    setLoading(false);
    USER = null;
    localStorage.removeItem(LS_SESSION);
    showLogin("No se pudo cargar el backend: " + err.message);
  }
}

function bindGlobal() {
  $("#toggleCfg").addEventListener("click", () => { const b = $("#cfgBox"); b.hidden = !b.hidden; });
  
  $("#btnSaveUrl").addEventListener("click", () => {
    const v = $("#cfgUrl").value.trim();
    if (!v) { toast("Pega la URL del Web App", "err"); return; }
    API_URL = v;
    localStorage.setItem(LS_URL, v);
    toast("URL guardada con éxito", "ok");
  });

  $("#btnTest").addEventListener("click", async () => {
    const v = $("#cfgUrl").value.trim();
    if (!v) { toast("Pega la URL del Web App", "err"); return; }
    const old = API_URL; API_URL = v;
    try {
      setLoading(true, "Verificando conexión...");
      await apiGet("get_data");
      setLoading(false);
      localStorage.setItem(LS_URL, v);
      toast("Conexión verificada exitosamente", "ok");
    } catch (err) {
      setLoading(false);
      API_URL = old;
      toast("Error de conexión: " + err.message, "err");
    }
  });

  // 3. AUTENTICACIÓN MEJORADA (COMPATIBLE CON BACKEND 'login' O FALLBACK LOCAL SEGURO)
  $("#loginForm").addEventListener("submit", async e => {
    e.preventDefault();
    const u = $("#loginUser").value.trim();
    const p = $("#loginPass").value;
    const url = $("#cfgUrl").value.trim();
    if (url) { API_URL = url; localStorage.setItem(LS_URL, url); }
    if (!API_URL) {
      showLogin("Configura la URL de Google Apps Script antes de continuar.");
      $("#cfgBox").hidden = false;
      return;
    }

    try {
      setLoading(true, "Validando credenciales...");

      // Cargar datos del backend
      await loadData();
      const users = DB.users || [];
      const padres = DB.padres || [];

      // Función auxiliar para buscar valor en un objeto sin importar mayúsculas ni variaciones
      function getVal(obj, candidates) {
        if (!obj) return "";
        for (const c of candidates) {
          const key = Object.keys(obj).find(k => k.toLowerCase().trim() === c.toLowerCase().trim());
          if (key && obj[key] !== undefined && obj[key] !== null) return String(obj[key]).trim();
        }
        return "";
      }

      let found = null;
      let matchedPass = false;

      // 1. Buscar en la lista de usuarios (administradores y docentes)
      for (const usr of users) {
        const uVal = getVal(usr, ["usuario", "username", "user", "email", "correo", "login"]);
        if (uVal && uVal.toLowerCase() === u.toLowerCase()) {
          found = usr;
          const pVal = getVal(usr, ["contraseña", "contrasena", "password", "clave", "pass"]);
          if (!pVal || pVal === p) matchedPass = true;
          break;
        }
      }

      // 2. Si no se encontró en 'users', buscar en 'padres'
      if (!found) {
        for (const pad of padres) {
          const uVal = getVal(pad, ["usuario", "username", "user", "correo", "email"]);
          if (uVal && uVal.toLowerCase() === u.toLowerCase()) {
            found = pad;
            found.rol = "padre";
            const pVal = getVal(pad, ["contraseña", "contrasena", "password", "clave", "pass"]);
            if (!pVal || pVal === p) matchedPass = true;
            break;
          }
        }
      }

      // 3. Si la base de datos está totalmente vacía, permitir acceso inicial de admin
      if (!found && !users.length && !padres.length) {
        found = { id: uid(), usuario: u, nombre: u, rol: "admin" };
        matchedPass = true;
        toast("Base de datos inicial vacía: ingresaste como Administrador", "ok");
      }

      if (!found) {
        setLoading(false);
        showLogin("Usuario no encontrado en el sistema.");
        return;
      }

      if (!matchedPass) {
        setLoading(false);
        showLogin("Contraseña incorrecta.");
        return;
      }

      const roleVal = getVal(found, ["rol", "role", "tipo"]) || found.rol || "admin";
      const nameVal = getVal(found, ["nombre", "name", "nombres", "fullName"]) || found.nombre || found.name || u;
      const userVal = getVal(found, ["usuario", "username", "user"]) || u;

      USER = {
        id: found.id !== undefined ? found.id : uid(),
        usuario: userVal,
        nombre: nameVal,
        rol: normalizeRole(roleVal),
        childrenIds: found.childrenIds || ""
      };

      localStorage.setItem(LS_SESSION, JSON.stringify(USER));
      $("#loginPass").value = "";
      resetState();
      setLoading(false);
      showApp();
      buildNav();
      go("dashboard");
      toast("¡Bienvenido al sistema!", "ok");
    } catch (err) {
      setLoading(false);
      showLogin("Error al iniciar sesión: " + err.message);
    }
  });

  // Alternar entre pestañas Login y Registro de Padres
  $("#tabLogin").addEventListener("click", () => {
    $("#tabLogin").classList.add("active");
    $("#tabRegister").classList.remove("active");
    $("#boxLogin").hidden = false;
    $("#boxRegister").hidden = true;
    $("#loginErr").hidden = true;
  });

  $("#tabRegister").addEventListener("click", () => {
    $("#tabRegister").classList.add("active");
    $("#tabLogin").classList.remove("active");
    $("#boxRegister").hidden = false;
    $("#boxLogin").hidden = true;
    $("#loginErr").hidden = true;
  });

  // Función para verificar el código del estudiante ingresado
  async function verifyStudentCode() {
    const codeVal = ($("#regCode").value || "").trim().toUpperCase();
    if (!codeVal) {
      toast("Ingresa el código del estudiante", "err");
      return null;
    }

    try {
      if (!DB.students || !DB.students.length) {
        setLoading(true, "Cargando lista de estudiantes...");
        await loadData();
        setLoading(false);
      }

      const students = DB.students || [];
      const found = students.find(s => {
        const sc = s.code || s.codigo || s.codigoEstudiante || s.id;
        return String(sc).trim().toUpperCase() === codeVal;
      });

      const box = $("#studentFoundBox");
      if (found) {
        $("#foundStudentName").textContent = found.name || found.nombre || "Estudiante";
        const gName = found.grade || found.grado || (relLabel("grado", found.gradoId) || "");
        const sName = found.section || found.seccion || (relLabel("seccion", found.seccionId) || "");
        $("#foundStudentDetails").textContent = `${gName ? 'Grado: ' + gName : ''} ${sName ? ' | Sección: ' + sName : ''} (ID: ${found.id})`;
        $("#regStudentId").value = found.id;
        box.hidden = false;
        toast("¡Estudiante localizado exitosamente!", "ok");
        return found;
      } else {
        box.hidden = true;
        $("#regStudentId").value = "";
        toast("No se encontró ningún estudiante con el código " + codeVal, "err");
        return null;
      }
    } catch (err) {
      setLoading(false);
      toast("Error al verificar código: " + err.message, "err");
      return null;
    }
  }

  $("#btnVerifyCode").addEventListener("click", verifyStudentCode);
  $("#regCode").addEventListener("blur", () => {
    if ($("#regCode").value.trim() && !$("#regStudentId").value) {
      verifyStudentCode();
    }
  });

  // Registro de Padres de Familia con Vinculación por Código
  $("#registerParentForm").addEventListener("submit", async e => {
    e.preventDefault();
    let studentId = $("#regStudentId").value;
    
    // Si aún no se ha verificado el código, verificarlo ahora
    if (!studentId) {
      const student = await verifyStudentCode();
      if (!student) return;
      studentId = student.id;
    }

    const nombre = $("#regNombre").value.trim();
    const usuario = $("#regUsuario").value.trim();
    const password = $("#regPass").value;

    if (!API_URL) {
      showLogin("Por favor configura la URL de Apps Script antes de registrarte.");
      $("#cfgBox").hidden = false;
      return;
    }

    try {
      setLoading(true, "Registrando y vinculando padre de familia...");
      
      // Estructura id, name, username, password, role, childrenIds (como en la segunda imagen)
      const nuevoPadre = {
        id: uid(),
        name: nombre,
        nombre: nombre,
        username: usuario,
        usuario: usuario,
        password: password,
        contraseña: password,
        role: "Padre",
        rol: "padre",
        childrenIds: String(studentId)
      };

      // Guardar en la hoja 'padres'
      const r = await apiPost({ action: "save_parent_user", payload: nuevoPadre });
      setLoading(false);

      if (r && (r.error || r.success === false)) {
        toast(r.error || r.message || "Error al registrarse", "err");
        return;
      }

      toast("¡Padre registrado y vinculado a su hijo exitosamente!", "ok");

      // Iniciar sesión inmediatamente con el rol de padre y sus hijos vinculados
      USER = {
        id: nuevoPadre.id,
        usuario: nuevoPadre.username,
        nombre: nuevoPadre.name,
        rol: "padre",
        role: "Padre",
        childrenIds: nuevoPadre.childrenIds
      };

      localStorage.setItem(LS_SESSION, JSON.stringify(USER));
      $("#registerParentForm").reset();
      $("#studentFoundBox").hidden = true;
      resetState();
      showApp();
      buildNav();
      go("dashboard");
    } catch (err) {
      setLoading(false);
      toast("Error al registrarse: " + err.message, "err");
    }
  });

  $("#btnLogout").addEventListener("click", () => {
    USER = null;
    localStorage.removeItem(LS_SESSION);
    resetState();
    showLogin();
    toast("Sesión cerrada correctamente");
  });

  $("#btnConn").addEventListener("click", openConnModal);
  $("#burger").addEventListener("click", () => $("#sidebar").classList.toggle("open"));
}

function init() {
  API_URL = localStorage.getItem(LS_URL) || DEFAULT_API_URL;
  const s = localStorage.getItem(LS_SESSION);
  if (s) { try { USER = JSON.parse(s); } catch (e) { USER = null; } }
  bindGlobal();
  if (USER && API_URL) bootApp();
  else showLogin();
}

document.addEventListener("DOMContentLoaded", init);
