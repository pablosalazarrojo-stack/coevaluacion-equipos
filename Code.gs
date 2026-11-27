/**
 * Coevaluación de equipo (MT1003B) — el "cerebro" en Google Apps Script.
 * Cada alumno califica a los demás miembros de SU equipo (no a sí mismo)
 * en 5 categorías, del 1 al 4. La lista de alumnos y la clave del profe viven aquí,
 * NO en la página pública (en GitHub van vacías).
 *
 * Instalación: proyecto nuevo → pegar → llenar ALUMNOS y ADMIN_INICIAL →
 * ejecutar setup → Implementar como aplicación web (ejecutar como yo, acceso: cualquiera).
 */

const CATS = ["Aporta ideas y conocimientos útiles", "Cumple con lo asignado o comprometido",
  "Escucha y respeta acuerdos", "Se concentra y colabora en reuniones", "Volvería a hacer equipo"];
const MIN = 1, MAX = 4;
const ADMIN_INICIAL = "CAMBIAR";
/* [matrícula, nombre, grupo ("CCM 515"), equipo ("Equipo 2")] — solo en el editor de Apps Script */
const ALUMNOS = [];

const NC = CATS.length;
const COLS = ["grupo", "equipo", "evaluador_matricula", "evaluador_nombre", "evaluado_matricula", "evaluado_nombre"]
  .concat(CATS, ["promedio", "estado", "actualizado"]);
const I_C = 6, I_PROM = 6 + NC, I_EST = 7 + NC, I_ACT = 8 + NC;

function setup() {
  const p = props();
  const id = p.getProperty("SHEET_ID");
  const ss = id ? SpreadsheetApp.openById(id) : SpreadsheetApp.create("Coevaluación MT1003B - respuestas");
  p.setProperty("SHEET_ID", ss.getId());

  const a = ss.getSheetByName("Alumnos") || ss.insertSheet("Alumnos");
  if (ALUMNOS.length) {
    a.clear();
    a.getRange(1, 1, 1, 4).setValues([["matricula", "nombre", "grupo", "equipo"]]).setFontWeight("bold");
    a.getRange(2, 1, ALUMNOS.length, 4).setNumberFormat("@").setValues(ALUMNOS);
    a.setFrozenRows(1);
    a.autoResizeColumns(1, 4);
  }
  const e = ss.getSheetByName("Evaluaciones") || ss.insertSheet("Evaluaciones", 0);
  e.getRange(1, 1, 1, e.getMaxColumns()).clearContent();
  e.getRange(1, 1, 1, COLS.length).setValues([COLS]).setFontWeight("bold");
  e.setFrozenRows(1);
  e.getRange("A:A").setNumberFormat("@");
  e.getRange("C:C").setNumberFormat("@");
  e.getRange("E:E").setNumberFormat("@");
  ["Sheet1", "Hoja 1", "Hoja1"].forEach(function (n) {
    const s = ss.getSheetByName(n);
    if (s && ss.getSheets().length > 1) ss.deleteSheet(s);
  });
  if (!p.getProperty("ADMIN_PASS")) p.setProperty("ADMIN_PASS", ADMIN_INICIAL);
  CacheService.getScriptCache().remove("alumnos");
  Logger.log(ss.getUrl());
}

function props() { return PropertiesService.getScriptProperties(); }
function libro() { return SpreadsheetApp.openById(props().getProperty("SHEET_ID")); }
function hoja() { return libro().getSheetByName("Evaluaciones"); }
function ronda() { return props().getProperty("RONDA") || "1"; }

function doGet() { return salida({ ok: true, msg: "Coevaluación de equipo activa" }); }

function doPost(e) {
  let out;
  try {
    const d = JSON.parse(e.postData.contents);
    switch (d.action) {
      case "login": out = login(d); break;
      case "draft": out = guardar(d, false); break;
      case "submit": out = guardar(d, true); break;
      case "profe": admin(d); out = tablero(); break;
      case "reset": admin(d); reiniciar(String(d.grupo || "")); out = tablero(); break;
      default: throw new Error("Acción desconocida.");
    }
  } catch (err) {
    out = { ok: false, error: String(err.message || err) };
  }
  return salida(out);
}

function salida(o) { return ContentService.createTextOutput(JSON.stringify(o)).setMimeType(ContentService.MimeType.JSON); }

function admin(d) {
  if (String(d.pass || "") !== props().getProperty("ADMIN_PASS")) {
    Utilities.sleep(1500);
    throw new Error("Contraseña incorrecta.");
  }
}

function norm(m) { return String(m || "").toUpperCase().replace(/[^A-Z0-9]/g, ""); }

function alumnos() {
  const c = CacheService.getScriptCache();
  const hit = c.get("alumnos");
  if (hit) return JSON.parse(hit);
  const sh = libro().getSheetByName("Alumnos");
  const v = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 4).getValues() : [];
  const out = v.filter(function (r) { return r[0]; }).map(function (r) {
    return { m: norm(r[0]), n: String(r[1]).trim(), g: String(r[2]).trim(), e: String(r[3]).trim() };
  });
  c.put("alumnos", JSON.stringify(out), 600);
  return out;
}

function alumno(m) {
  const a = alumnos().filter(function (x) { return x.m === norm(m); })[0];
  if (!a) { Utilities.sleep(700); throw new Error("No encontramos esa matrícula. Revísala: empieza con A0…"); }
  return a;
}

function companeros(a) {
  return alumnos().filter(function (x) { return x.g === a.g && x.e === a.e && x.m !== a.m; });
}

/* "Apellidos, Nombre" o "Nombre Apellidos" (algunas listas vienen así) */
function pila(nombre) { const i = nombre.indexOf(","); return i >= 0 ? (nombre.slice(i + 1).trim() || nombre) : nombre.split(" ")[0]; }
function bonito(nombre) { const i = nombre.indexOf(","); return i >= 0 ? (nombre.slice(i + 1).trim() + " " + nombre.slice(0, i).trim()) : nombre; }

function filas() {
  const sh = hoja();
  if (sh.getLastRow() < 2) return [];
  return sh.getRange(2, 1, sh.getLastRow() - 1, COLS.length).getValues().filter(function (r) { return r[0]; });
}

function deAlumno(m, rows) {
  const ev = {};
  let enviado = false;
  rows.forEach(function (r) {
    if (norm(r[2]) !== m) return;
    const t = r[I_ACT] instanceof Date ? r[I_ACT].getTime() : (Date.parse(r[I_ACT]) || 0);
    ev[norm(r[4])] = { d: r.slice(I_C, I_C + NC).map(Number), t: t };
    if (r[I_EST] === "enviado") enviado = true;
  });
  return { evals: ev, enviado: enviado };
}

function login(d) {
  const a = alumno(d.matricula);
  const yo = deAlumno(a.m, filas());
  return {
    ok: true, matricula: a.m, nombre: pila(a.n), grupo: a.g, equipo: a.e, ronda: ronda(),
    cats: CATS, min: MIN, max: MAX,
    companeros: companeros(a).map(function (x) { return { m: x.m, nombre: bonito(x.n) }; }),
    evals: yo.evals, enviado: yo.enviado
  };
}

function califs(x) {
  if (!x || x.length !== NC) throw new Error("Faltan calificaciones.");
  return x.map(function (v) {
    v = Math.round(Number(v));
    if (!(v >= MIN && v <= MAX)) throw new Error("Calificación fuera de rango.");
    return v;
  });
}

/* draft: un compañero, "borrador".  submit: todos los compañeros, "enviado". */
function guardar(d, final) {
  const a = alumno(d.matricula);
  if (d.ronda && String(d.ronda) !== ronda()) throw new Error("RONDA: El profe reinició la evaluación. Entra de nuevo.");
  const comp = companeros(a);
  const ids = comp.map(function (x) { return x.m; });
  const evals = {};
  if (final) {
    comp.forEach(function (x) { evals[x.m] = califs(d.evals && d.evals[x.m]); });
  } else {
    const ev = norm(d.evaluado);
    if (ids.indexOf(ev) < 0) throw new Error("Solo puedes evaluar a los miembros de tu equipo.");
    evals[ev] = califs(d.d);
  }
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = hoja();
    const rows = filas();
    if (deAlumno(a.m, rows).enviado) {
      if (final) return { ok: true, ya: true };
      throw new Error("Ya enviaste tus calificaciones.");
    }
    const all = sh.getLastRow() > 1 ? sh.getRange(2, 1, sh.getLastRow() - 1, 5).getValues() : [];
    const idx = {};
    all.forEach(function (r, i) { if (r[0]) idx[norm(r[2]) + "|" + norm(r[4])] = i + 2; });
    const ahora = new Date();
    const nuevas = [];
    Object.keys(evals).forEach(function (m) {
      const v = evals[m];
      const ev = comp.filter(function (x) { return x.m === m; })[0];
      const prom = Math.round(v.reduce(function (s, x) { return s + x; }, 0) / v.length * 100) / 100;
      const fila = [a.g, a.e, a.m, a.n, m, ev.n].concat(v, [prom, final ? "enviado" : "borrador", ahora]);
      const at = idx[a.m + "|" + m];
      if (at) sh.getRange(at, 1, 1, fila.length).setValues([fila]);
      else nuevas.push(fila);
    });
    if (nuevas.length) sh.getRange(sh.getLastRow() + 1, 1, nuevas.length, COLS.length).setValues(nuevas);
    return { ok: true, t: ahora.getTime() };
  } finally {
    lock.releaseLock();
  }
}

function reiniciar(grupo) {
  const lock = LockService.getScriptLock();
  lock.waitLock(30000);
  try {
    const sh = hoja();
    /* se borra por la lista del grupo (quién evalúa), no por la columna grupo */
    const ms = {};
    alumnos().forEach(function (x) { if (!grupo || x.g === grupo) ms[x.m] = true; });
    if (sh.getLastRow() > 1) {
      const v = sh.getRange(2, 1, sh.getLastRow() - 1, 3).getValues();
      for (let i = v.length - 1; i >= 0; i--) {
        if (!grupo || ms[norm(v[i][2])]) sh.deleteRow(i + 2);
      }
    }
    props().setProperty("RONDA", String(Date.now()));
  } finally {
    lock.releaseLock();
  }
}

/* Final ecualizado: cada 4 vale 100, cada 3 vale 85, cada 2 vale 70 y cada 1 vale 50; se promedian */
const ECUAL = { 4: 100, 3: 85, 2: 70, 1: 50 };

function tablero() {
  const rows = filas();
  const al = alumnos();
  const recib = {}, hechas = {}, env = {};
  rows.forEach(function (r) {
    const de = norm(r[2]), a = norm(r[4]);
    hechas[de] = (hechas[de] || 0) + 1;
    if (r[I_EST] === "enviado") {
      env[de] = true;
      const t = recib[a] = recib[a] || { n: 0, e: 0, s: CATS.map(function () { return 0; }) };
      t.n++;
      for (let i = 0; i < NC; i++) { const c = Number(r[I_C + i]) || 0; t.s[i] += c; t.e += ECUAL[c] || 0; }
    }
  });
  const tam = {};
  al.forEach(function (x) { const k = x.g + "|" + x.e; tam[k] = (tam[k] || 0) + 1; });
  return {
    ok: true, ronda: ronda(), cats: CATS, min: MIN, max: MAX,
    alumnos: al.map(function (x) {
      const r = recib[x.m];
      const porCat = r ? r.s.map(function (s) { return Math.round(s / r.n * 100) / 100; }) : null;
      return {
        nombre: x.n, matricula: x.m, grupo: x.g, equipo: x.e,
        estado: env[x.m] ? "enviado" : (hechas[x.m] ? "en curso" : "sin empezar"),
        hechas: env[x.m] ? tam[x.g + "|" + x.e] - 1 : (hechas[x.m] || 0),
        porEvaluar: tam[x.g + "|" + x.e] - 1,
        recibidas: r ? r.n : 0,
        porCat: porCat,
        ecual: r ? Math.round(r.e / (r.n * NC) * 100) / 100 : null,
        promedio: r ? Math.round(r.s.reduce(function (s, v) { return s + v; }, 0) / (r.n * NC) * 100) / 100 : null
      };
    })
  };
}
