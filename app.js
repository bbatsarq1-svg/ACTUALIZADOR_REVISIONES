/* app.js — port de subir_revision.py. Mismas funciones y mismas reglas que el Python.
   Requiere JSZip (global). Sin dependencias del navegador salvo DOMParser / XMLSerializer. */

// ---- REGLAS DEL PROYECTO (lo único que se ajusta con el tiempo) ----
const NOMBRE = /^(?<codigo>[A-Z]{4}-[A-Z]{2}-[A-Z]{3}-[A-Z]{3}-[A-Z]{3}-[A-Z]{3}-\d{4})-(?<rev>[A-Z])$/;
const TIPOS_PIE = { CRD: "CRITERIOS DE DISEÑO", ETT: "ESPECIFICACIONES TÉCNICAS", MEM: "MEMORIA" };
const PLACEHOLDER = "CÓDIGO DOCUMENTO";
const FILAS_EXCEL = [22, 23, 24, 25, 26]; // tabla de revisiones en PORTADA (columnas B..F)

const W = "http://schemas.openxmlformats.org/wordprocessingml/2006/main";
const M = "http://schemas.openxmlformats.org/spreadsheetml/2006/main";
const R = "http://schemas.openxmlformats.org/officeDocument/2006/relationships";
const PARTES_WORD = /^word\/(document|header\d*|footer\d*)\.xml$/;

// ---- utilidades ----
const hijos = (el, tag) => Array.from(el.childNodes).filter(n => n.localName === tag);
const todos = (el, ns, tag) => Array.from(el.getElementsByTagNameNS(ns, tag));
const texto = el => [...todos(el, W, "t"), ...todos(el, M, "t")].map(t => t.textContent).join("");
const plano = s => s.toLowerCase().normalize("NFD").replace(/[\u0300-\u036f]/g, "");

const DIA = 864e5, BASE = Date.UTC(1899, 11, 30); // las fechas de Excel cuentan días desde el 30-12-1899
function aFecha(s) { // '02-10-2026' o '46297' -> milisegundos UTC (null si no es fecha)
  s = String(s).trim();
  if (/^\d+$/.test(s)) return BASE + Number(s) * DIA;
  const m = /^(\d{2})-(\d{2})-(\d{4})$/.exec(s);
  if (!m) return null;
  const t = Date.UTC(+m[3], +m[2] - 1, +m[1]), d = new Date(t);
  return d.getUTCDate() === +m[1] && d.getUTCMonth() === +m[2] - 1 ? t : null;
}
const serial = t => Math.round((t - BASE) / DIA);
const fmt = t => { const d = new Date(t), p = n => String(n).padStart(2, "0"); return `${p(d.getUTCDate())}-${p(d.getUTCMonth() + 1)}-${d.getUTCFullYear()}`; };

// ---- paquete: un .docx/.xlsx es un zip de XML ----
async function abrir(bytes) {
  const zip = await JSZip.loadAsync(bytes), datos = {};
  for (const n of Object.keys(zip.files)) if (/\.(xml|rels)$/.test(n)) datos[n] = await zip.file(n).async("string");
  return { zip, datos, arboles: {}, modificados: new Set() };
}
function xml(p, n) {
  p.arboles[n] ||= new DOMParser().parseFromString(p.datos[n], "application/xml");
  return p.arboles[n].documentElement;
}
async function guardar(p) { // solo se reescriben las partes modificadas
  for (const n of p.modificados) {
    const s = n in p.arboles ? new XMLSerializer().serializeToString(p.arboles[n]).replace(/^<\?xml[^>]*\?>\s*/, "") : p.datos[n];
    p.zip.file(n, n in p.arboles ? '<?xml version="1.0" encoding="UTF-8" standalone="yes"?>\r\n' + s : s);
  }
  return p.zip.generateAsync({ type: "uint8array", compression: "DEFLATE" });
}

// ---- tabla de revisiones (misma regla para Word y Excel) ----
function revisarTabla(filas, rev, nueva, fecha) {
  const llenas = filas.map((f, i) => (f[0] ? i : -1)).filter(i => i >= 0);
  if (!llenas.length) return { msgs: [["ERROR", "no encuentro filas con datos en la tabla de revisiones"]] };
  const ult = llenas.at(-1), msgs = [];
  if (filas[ult][0] !== rev) msgs.push(["ERROR", `la última revisión de la tabla es ${filas[ult][0]} pero el nombre dice ${rev}`]);
  const libre = ult + 1 < filas.length ? ult + 1 : null;
  if (libre === null) msgs.push(["ERROR", "la tabla de revisiones no tiene filas libres"]);
  if (nueva <= rev) msgs.push(["ERROR", `la revisión nueva (${nueva}) debe ser posterior a la actual (${rev})`]);
  else if (nueva.charCodeAt(0) !== rev.charCodeAt(0) + 1) msgs.push(["AVISO", `se salta de ${rev} a ${nueva} (no es la letra siguiente)`]);
  const anterior = aFecha(filas[ult][1]);
  if (anterior !== null && fecha <= anterior) msgs.push(["AVISO", `la fecha nueva (${fmt(fecha)}) no es posterior a la de la fila ${filas[ult][0]} (${fmt(anterior)})`]);
  return { ult, libre, msgs };
}

// ---- WORD ----
function ponerTextoParrafo(p, nuevo) { // escribe en el primer trozo y vacía el resto (conserva el formato)
  const ts = todos(p, W, "t");
  if (ts.length) { ts[0].textContent = nuevo; ts.slice(1).forEach(t => (t.textContent = "")); return; }
  const doc = p.ownerDocument, run = doc.createElementNS(W, "w:r"), t = doc.createElementNS(W, "w:t");
  const pPr = hijos(p, "pPr")[0], marca = pPr && hijos(pPr, "rPr")[0];
  if (marca) run.appendChild(marca.cloneNode(true));
  t.textContent = nuevo; run.appendChild(t); p.appendChild(run);
}
const ponerTextoCelda = (tc, nuevo) => ponerTextoParrafo(hijos(tc, "p")[0], nuevo);

function inspeccionarDocx(p, codigo, rev) {
  const partes = [...codigo.split("-"), rev], est = { codigos: [], filaCajas: null, tabla: [], msgs: [] }, pies = {};
  for (const nombre of Object.keys(p.datos).filter(n => PARTES_WORD.test(n)).sort()) {
    const raiz = xml(p, nombre);
    if (nombre.includes("footer")) { pies[nombre] = texto(raiz).trim(); continue; }
    for (const par of todos(raiz, W, "p")) { // solo párrafos que contienen SOLO el código (no las citas a otros documentos)
      const t = texto(par).trim();
      if (t === `${codigo}-${rev}` || plano(t) === plano(PLACEHOLDER)) {
        est.codigos.push([nombre, par]);
        if (plano(t) === plano(PLACEHOLDER)) est.msgs.push(["AVISO", `${nombre}: decía «${t}»; se escribe el código (estándar Rev C)`]);
      }
    }
  }
  for (const tbl of todos(xml(p, "word/document.xml"), W, "tbl")) {
    const filas = hijos(tbl, "tr");
    filas.forEach((tr, i) => {
      const t = hijos(tr, "tc").map(c => texto(c).trim());
      if (t.join("|") === partes.join("|")) est.filaCajas = tr;
      if (t[0] === "REV." && t[1] === "FECHA") est.tabla = filas.slice(i + 2); // cabecera + fila de subtítulos
    });
  }
  if (!est.codigos.length) est.msgs.push(["ERROR", "no encuentro el código en la carátula ni en los encabezados"]);
  if (!est.filaCajas) est.msgs.push(["ERROR", `las casillas de la carátula no coinciden con ${partes.join("-")}`]);
  if (!est.tabla.length) est.msgs.push(["ERROR", "no encuentro la tabla de revisiones"]);
  est.filas = est.tabla.map(tr => [...hijos(tr, "tc").map(c => texto(c).trim()), "", "", "", "", ""].slice(0, 5));
  const clave = TIPOS_PIE[partes[5]]; // el pie debe nombrar el tipo de documento
  for (const [nombre, t] of Object.entries(pies))
    if (clave && t && !plano(t).includes(plano(clave))) est.msgs.push(["AVISO", `${nombre}: el pie dice «${t.slice(0, 45)}» y el tipo ${partes[5]} pide «${clave}»`]);
  return est;
}

function aplicarDocx(p, est, codigo, rev, nueva, fecha, firmas, libre, ult) {
  const nuevoCodigo = `${codigo}-${nueva}`;
  for (const [nombre, par] of est.codigos) { ponerTextoParrafo(par, nuevoCodigo); p.modificados.add(nombre); }
  ponerTextoCelda(hijos(est.filaCajas, "tc").at(-1), nueva);
  const valores = [nueva, fmt(fecha), ...firmas.map((f, i) => f || est.filas[ult][2 + i])];
  hijos(est.tabla[libre], "tc").forEach((c, i) => i < valores.length && ponerTextoCelda(c, valores[i]));
  p.modificados.add("word/document.xml");
  const core = "docProps/core.xml"; // propiedades del documento (descripción con el código)
  if (core in p.datos) { p.datos[core] = p.datos[core].replaceAll(`${codigo}-${rev}`, nuevoCodigo); p.modificados.add(core); }
}

function restantes(p, viejo) { // partes donde aún aparece el código anterior dentro de un texto
  return Object.keys(p.datos).filter(n => PARTES_WORD.test(n) && todos(xml(p, n), W, "p").some(par => texto(par).includes(viejo)));
}

// ---- EXCEL (se edita el XML a mano: así no se dañan nombres definidos, estilos, etc.) ----
function rutaPortada(p) {
  const rid = todos(xml(p, "xl/workbook.xml"), M, "sheet").find(s => s.getAttribute("name") === "PORTADA").getAttributeNS(R, "id");
  const destino = Array.from(xml(p, "xl/_rels/workbook.xml.rels").childNodes).find(r => r.getAttribute && r.getAttribute("Id") === rid).getAttribute("Target");
  return "xl/" + destino.replace(/^\//, "").replace(/^xl\//, "");
}
const celda = (hoja, ref) => todos(hoja, M, "c").find(c => c.getAttribute("r") === ref);
function valor(c, comp) {
  if (!c) return "";
  const t = c.getAttribute("t"), v = todos(c, M, "v")[0];
  if (t === "s") return comp[+v.textContent];
  if (t === "inlineStr") return texto(c);
  return v ? v.textContent : "";
}
const vaciar = c => Array.from(c.childNodes).forEach(h => c.removeChild(h));
function escribirTexto(c, txt) { // texto dentro de la propia celda: no toca los textos compartidos
  vaciar(c); c.setAttribute("t", "inlineStr");
  const d = c.ownerDocument, is = d.createElementNS(M, "is"), t = d.createElementNS(M, "t");
  t.textContent = txt; is.appendChild(t); c.appendChild(is);
}
function escribirNumero(c, n) {
  vaciar(c); c.removeAttribute("t");
  const v = c.ownerDocument.createElementNS(M, "v"); v.textContent = String(n); c.appendChild(v);
}
function copiarValor(destino, origen) {
  vaciar(destino); destino.removeAttribute("t");
  if (origen.getAttribute("t")) destino.setAttribute("t", origen.getAttribute("t"));
  Array.from(origen.childNodes).forEach(h => destino.appendChild(h.cloneNode(true)));
}

function inspeccionarXlsx(p, codigo, rev) {
  const partes = [...codigo.split("-"), rev], ruta = rutaPortada(p), hoja = xml(p, ruta), est = { hoja, ruta, msgs: [] };
  const comp = "xl/sharedStrings.xml" in p.datos ? hijos(xml(p, "xl/sharedStrings.xml"), "si").map(texto) : [];
  const v = ref => valor(celda(hoja, ref), comp).trim();
  if ([..."BCDEFGHI"].map(c => v(c + "6")).join("|") !== partes.join("|")) est.msgs.push(["ERROR", `las casillas de la fila 6 no coinciden con ${partes.join("-")}`]);
  if (v("B4") !== `${codigo}-${rev}`) est.msgs.push(["AVISO", `B4 decía «${v("B4")}»; se escribe ${codigo}-<nueva> (estándar Rev C)`]);
  est.filas = FILAS_EXCEL.map(r => [..."BCDEF"].map(c => v(c + r)));
  return est;
}

function aplicarXlsx(p, est, codigo, rev, nueva, fecha, firmas, libre, ult) {
  const hoja = est.hoja, r = FILAS_EXCEL[libre], ant = FILAS_EXCEL[ult];
  escribirTexto(celda(hoja, "B4"), `${codigo}-${nueva}`);
  escribirTexto(celda(hoja, "I6"), nueva);
  escribirTexto(celda(hoja, `B${r}`), nueva);
  const fec = celda(hoja, `C${r}`);
  escribirNumero(fec, serial(fecha));
  fec.setAttribute("s", celda(hoja, `C${ant}`).getAttribute("s")); // las filas libres no traen formato de fecha: se copia el de la fila anterior
  [..."DEF"].forEach((col, i) => firmas[i] ? escribirTexto(celda(hoja, `${col}${r}`), firmas[i]) : copiarValor(celda(hoja, `${col}${r}`), celda(hoja, `${col}${ant}`)));
  p.modificados.add(est.ruta);
}

// ---- proceso por archivo ----
const FORMATOS = { ".docx": [inspeccionarDocx, aplicarDocx], ".xlsx": [inspeccionarXlsx, aplicarXlsx] };
const hayError = msgs => msgs.some(m => m[0] === "ERROR");

async function verificar(bytes, ext, codigo, rev, nueva, fecha) { // relee el archivo generado
  const p = await abrir(bytes), est = FORMATOS[ext][0](p, codigo, nueva);
  const msgs = est.msgs.filter(m => m[0] === "ERROR").map(m => ["ERROR", "verificación: " + m[1]]);
  const llenas = est.filas.filter(f => f[0]);
  if (!llenas.length || llenas.at(-1)[0] !== nueva || aFecha(llenas.at(-1)[1]) !== fecha) msgs.push(["ERROR", "verificación: la última fila de la tabla no es la esperada"]);
  if (ext === ".docx") for (const n of restantes(p, `${codigo}-${rev}`)) msgs.push(["AVISO", `${n}: todavía aparece ${codigo}-${rev} dentro de un texto; revísalo a mano`]);
  return msgs;
}

// procesar(nombre, bytes, {nueva, fecha(ms UTC), firmas:[elab,rev,acep], simular}) -> {nombre, estado, msgs, salida}
async function procesar(nombre, bytes, { nueva, fecha, firmas = ["", "", ""], simular }) {
  const ext = nombre.slice(nombre.lastIndexOf(".")).toLowerCase(), m = NOMBRE.exec(nombre.slice(0, -ext.length));
  if (!m) return { nombre, estado: "OMITIDO", msgs: [["AVISO", "el nombre no sigue el formato SCAT-XX-XXX-XXX-XXX-XXX-0000-X"]] };
  const { codigo, rev } = m.groups, [inspeccionar, aplicar] = FORMATOS[ext];
  try {
    const p = await abrir(bytes), est = inspeccionar(p, codigo, rev), t = revisarTabla(est.filas, rev, nueva, fecha);
    let msgs = [...est.msgs, ...t.msgs];
    if (hayError(msgs)) return { nombre, estado: "ERROR", msgs };
    if (simular) return { nombre, estado: "OK", msgs };
    aplicar(p, est, codigo, rev, nueva, fecha, firmas, t.libre, t.ult);
    const salida = await guardar(p);
    msgs = [...msgs, ...(await verificar(salida, ext, codigo, rev, nueva, fecha))];
    return { nombre, estado: hayError(msgs) ? "ERROR" : "OK", msgs, salida: { nombre: `${codigo}-${nueva}${ext}`, bytes: salida } };
  } catch (e) {
    return { nombre, estado: "ERROR", msgs: [["ERROR", `no se pudo procesar: ${e}`]] };
  }
}

function informe(resultados, simular) { // mismo formato de texto que el script de Python
  const cuenta = { OK: 0, ERROR: 0, OMITIDO: 0 }, marca = { OK: "✔", ERROR: "✖", OMITIDO: "·" }, lineas = [];
  for (const r of resultados) {
    cuenta[r.estado]++;
    lineas.push(`${marca[r.estado]} ${r.nombre}${simular && r.estado === "OK" ? "  [simulación]" : ""}`);
    r.msgs.forEach(([n, t]) => lineas.push(`     ${n}: ${t}`));
  }
  lineas.push(`\n${cuenta.OK} listos · ${cuenta.ERROR} con error (no se generaron) · ${cuenta.OMITIDO} omitidos`);
  return lineas.join("\n");
}

if (typeof module !== "undefined") module.exports = { procesar, informe, aFecha };
