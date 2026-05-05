// xlsx-loader.js — Carga datos directamente desde mad-map-data-v2.xlsx
// usando SheetJS, sin scripts intermedios ni JSON precomputado.
//
// El archivo .xlsx commiteado en el repo es la fuente única de verdad. Se edita
// a mano (Excel, Numbers, LibreOffice) y al recargar la página la visualización
// refleja los cambios.
//
// Layout: el embedding estructural (PCA/SVD) fue removido. graph.js usa el
// force-directed de D3, que recalcula posiciones al vuelo desde las aristas.
//
// Dependencias:
//   - SheetJS (xlsx.full.min.js) cargado por CDN antes de este script.
//
// API pública:
//   window.MadMapDataLoader.loadFromXlsx(url) -> Promise<DataModel>
//
// La estructura DataModel devuelta es la misma que consumía graph.js cuando
// leía mad-map-data.json o gviz, salvo que ya no incluye posiciones bakeadas.

// ===========================================================================
// Lectura de hojas con SheetJS
// ===========================================================================

// Detecta la fila de encabezados: la primera donde, ignorando celdas vacías
// al final, queden al menos 2 celdas no-vacías sin huecos. Las hojas creadas
// por seed_xlsx.py ponen título y nota arriba; esta heurística las salta.
// El recorte de trailing-empty también tolera columnas borradas que dejaron
// celdas vacías al final del rango.
//
// `aoa` es array-de-arrays (lo que devuelve XLSX.utils.sheet_to_json con
// header:1).
function detectHeaderRow(aoa, maxScan = 8) {
  const isEmpty = c => c === undefined || c === null || String(c).trim() === '';
  for (let r = 0; r < Math.min(maxScan, aoa.length); r++) {
    const raw = aoa[r] || [];
    let row = raw.slice();
    while (row.length && isEmpty(row[row.length - 1])) row.pop();
    if (row.length >= 2 && row.every(c => !isEmpty(c))) {
      return r;
    }
  }
  return 0;
}

// Convierte una hoja del workbook a un array de objetos {col: valor},
// autodetectando la fila de encabezados. Llamado desde buildDataFromTabs
// para cada una de las hojas del modelo (01_Lineas, 02_Sublineas, etc.).
function sheetToObjects(workbook, sheetName) {
  const ws = workbook.Sheets[sheetName];
  if (!ws) return [];
  const aoa = XLSX.utils.sheet_to_json(ws, { header: 1, defval: '', raw: false });
  if (aoa.length === 0) return [];
  const headerIdx = detectHeaderRow(aoa);
  const headers = (aoa[headerIdx] || []).map(h => h == null ? '' : String(h).trim());
  const out = [];
  for (let r = headerIdx + 1; r < aoa.length; r++) {
    const row = aoa[r] || [];
    if (row.every(c => c === undefined || c === null || String(c).trim() === '')) continue;
    const obj = {};
    headers.forEach((h, i) => {
      if (h) obj[h] = row[i] == null ? '' : row[i];
    });
    out.push(obj);
  }
  return out;
}

// Listado de hojas que componen el modelo. Si alguna falta, el loader sigue
// adelante con array vacío (las visualizaciones se degradan elegantemente).
//
// Tras el paso 2 de usabilidad humana, las hojas de relación referencian
// entidades por NOMBRE en lugar de por ID. La hoja 09_Sublinea_Tema fue
// absorbida en la nueva 08_Temas, que ahora contiene filas (sublínea,
// investigador, tema). El loader resuelve los nombres a IDs internos al
// cargar.
const SHEET_NAMES = [
  '01_Lineas',
  '02_Sublineas',
  '03_Areas',
  '04_Modos',
  '05_Salidas',
  '06_Laboratorios',
  '07_Investigadores',
  '08_Temas',
  '10_Lab_Linea',
  '11_Lab_Salida',
  '12_Investigador_Lab',
  '13_Investigador_Modo',
  '14_Linea_Modo',
  '17_Sello',
  '18_Proximidad_Tematica',
];

// Lee todas las hojas relevantes del workbook y devuelve un dict
// {sheetName: [obj, obj, ...]}. Llamado una sola vez por loadFromXlsx().
function readAllSheets(workbook) {
  const tabs = {};
  for (const name of SHEET_NAMES) {
    tabs[name] = sheetToObjects(workbook, name);
  }
  return tabs;
}

// ===========================================================================
// Resolución por nombre
// ===========================================================================

// Construye un mapa nombre->id a partir de un array de entidades. Detecta
// nombres duplicados y los reporta en consola (no rompe la carga, pero el
// editor debería resolverlos). Llamado por buildDataFromTabs para cada
// tipo de entidad antes de procesar las hojas de relación.
function buildNameToIdMap(entities, label) {
  const map = {};
  const seen = new Set();
  const dupes = [];
  for (const e of entities) {
    if (!e.nombre) continue;
    const key = e.nombre.trim();
    if (seen.has(key)) {
      dupes.push(key);
      continue;
    }
    seen.add(key);
    map[key] = e.id;
  }
  if (dupes.length) {
    console.warn(
      `[xlsx-loader] Nombres duplicados en ${label}: ${dupes.join(', ')}. ` +
      `Las relaciones que los referencien usarán la primera ocurrencia.`
    );
  }
  return map;
}

// Resuelve un valor que podría ser un nombre o estar ausente. Si no resuelve,
// emite warning con el contexto y devuelve null. El llamador suele filtrar
// las filas con null para no propagar referencias rotas a graph.js.
function resolveName(value, map, contextLabel) {
  if (value == null || value === '') return null;
  const trimmed = String(value).trim();
  if (trimmed in map) return map[trimmed];
  console.warn(`[xlsx-loader] Referencia no resuelta en ${contextLabel}: ${JSON.stringify(trimmed)}`);
  return null;
}

// ===========================================================================
// Conversión de tabs a estructura de datos del frontend
// ===========================================================================

// Construye el modelo que consume graph.js a partir de los tabs ya parseados.
// Es el mismo contrato JSON que esperaba antes (entidades + relations + edges)
// pero ahora las hojas de relación traen nombres y los resolvemos a IDs aquí.
function buildDataFromTabs(tabs) {
  // 1) Entidades: cada hoja primaria conserva su columna `id` interna.
  const lineas = (tabs['01_Lineas'] || []).map(r => ({
    id: r.id, nombre: r.nombre,
    descripcion: r['descripción'] || r.descripcion,
    estado: r.estado || '',
  })).filter(l => l.id);

  const areas = (tabs['03_Areas'] || []).map(r => ({
    id: r['código'] || r.codigo,
    nombre: r.nombre,
    descripcion: r['descripción'] || r.descripcion,
  })).filter(a => a.id);

  const modos = (tabs['04_Modos'] || []).map(r => ({
    id: r.id, nombre: r.nombre,
    descripcion: r['descripción'] || r.descripcion,
  })).filter(m => m.id);

  const salidas = (tabs['05_Salidas'] || []).map(r => ({
    id: r.id, nombre: r.nombre,
    descripcion: r['descripción'] || r.descripcion,
  })).filter(s => s.id);

  const laboratorios = (tabs['06_Laboratorios'] || []).map(r => ({
    id: r.id, nombre: r.nombre,
    descripcion: r['descripción'] || r.descripcion,
  })).filter(l => l.id);

  const investigadores = (tabs['07_Investigadores'] || []).map(r => ({
    id: r.id, nombre: r.nombre,
    area_principal: r['área_principal'] || r.area_principal,
    perfil_url: r.perfil_casiopea || r.perfil_url || '',
  })).filter(i => i.id);

  // 2) Mapas nombre->id (todas las hojas que se referencian por nombre).
  const lineaN2I = buildNameToIdMap(lineas, '01_Lineas');
  const areaN2I = buildNameToIdMap(areas, '03_Areas');
  const modoN2I = buildNameToIdMap(modos, '04_Modos');
  const salidaN2I = buildNameToIdMap(salidas, '05_Salidas');
  const labN2I = buildNameToIdMap(laboratorios, '06_Laboratorios');
  const invN2I = buildNameToIdMap(investigadores, '07_Investigadores');

  // 3) Sublíneas: las columnas `línea` y `área` traen NOMBRES, hay que
  //    resolverlas a IDs antes de armar el mapa nombre->id de sublíneas.
  const sublineas = (tabs['02_Sublineas'] || []).map(r => ({
    id: r.id,
    nombre: r.nombre,
    linea: resolveName(r['línea'], lineaN2I, `02_Sublineas[${r.id}].línea`),
    area: resolveName(r['área'], areaN2I, `02_Sublineas[${r.id}].área`),
    notas: r.notas || '',
  })).filter(s => s.id);

  const subN2I = buildNameToIdMap(sublineas, '02_Sublineas');

  // 4) 08_Temas en formato nuevo: cada fila es (sublínea, investigador, tema).
  //    Generamos IDs sintéticos para los temas (TEM-NNNN) y reconstruimos
  //    relations.sublinea_tema con los pares (sublinea_id, tema_id) que
  //    consume computeEdges. Los IDs sintéticos viven solo en memoria.
  const temasRaw = (tabs['08_Temas'] || []).filter(r =>
    r['sublínea'] && r.investigador && r.tema
  );
  const temas = [];
  const sublinea_tema = [];
  temasRaw.forEach((r, idx) => {
    const tema_id = `TEM-${String(idx + 1).padStart(4, '0')}`;
    const inv_id = resolveName(r.investigador, invN2I,
      `08_Temas[fila ${idx + 1}].investigador`);
    const sub_id = resolveName(r['sublínea'], subN2I,
      `08_Temas[fila ${idx + 1}].sublínea`);
    if (!inv_id || !sub_id) return;
    temas.push({ id: tema_id, investigador: inv_id, texto: r.tema });
    sublinea_tema.push({ sublinea_id: sub_id, tema_id });
  });

  // 5) Resto de relaciones: cada hoja tiene columnas con nombres a resolver.
  const lab_linea = (tabs['10_Lab_Linea'] || []).map((r, i) => ({
    lab_id: resolveName(r.laboratorio, labN2I, `10_Lab_Linea[${i}].laboratorio`),
    linea_id: resolveName(r['línea'], lineaN2I, `10_Lab_Linea[${i}].línea`),
  })).filter(r => r.lab_id && r.linea_id);

  const lab_salida = (tabs['11_Lab_Salida'] || []).map((r, i) => ({
    lab_id: resolveName(r.laboratorio, labN2I, `11_Lab_Salida[${i}].laboratorio`),
    salida_id: resolveName(r.salida, salidaN2I, `11_Lab_Salida[${i}].salida`),
  })).filter(r => r.lab_id && r.salida_id);

  const inv_lab = (tabs['12_Investigador_Lab'] || []).map((r, i) => ({
    investigador_id: resolveName(r.investigador, invN2I,
      `12_Investigador_Lab[${i}].investigador`),
    lab_id: resolveName(r.laboratorio, labN2I,
      `12_Investigador_Lab[${i}].laboratorio`),
  })).filter(r => r.investigador_id && r.lab_id);

  const inv_modo = (tabs['13_Investigador_Modo'] || []).map((r, i) => ({
    investigador_id: resolveName(r.investigador, invN2I,
      `13_Investigador_Modo[${i}].investigador`),
    modo_id: resolveName(r.modo, modoN2I, `13_Investigador_Modo[${i}].modo`),
  })).filter(r => r.investigador_id && r.modo_id);

  const linea_modo = (tabs['14_Linea_Modo'] || []).map((r, i) => ({
    linea_id: resolveName(r['línea'], lineaN2I, `14_Linea_Modo[${i}].línea`),
    modo_id: resolveName(r.modo, modoN2I, `14_Linea_Modo[${i}].modo`),
    nivel: r.nivel || '',
  })).filter(r => r.linea_id && r.modo_id);

  const proximidad = (tabs['18_Proximidad_Tematica'] || [])
    .filter(p => p.estado !== 'DESCARTADO' && p['sublínea_a'] && p['sublínea_b'])
    .map((p, i) => ({
      sublinea_a_id: resolveName(p['sublínea_a'], subN2I,
        `18_Proximidad_Tematica[${i}].sublínea_a`),
      sublinea_b_id: resolveName(p['sublínea_b'], subN2I,
        `18_Proximidad_Tematica[${i}].sublínea_b`),
      afinidad: parseFloat(p.afinidad) || 0,
      razonamiento: p.razonamiento,
      estado: p.estado,
    }))
    .filter(p => p.sublinea_a_id && p.sublinea_b_id);

  // 6) Sello formativo: variante elegida (la fila cuyo "Foco" contiene "ELEGIDO").
  const sello = (() => {
    const rows = tabs['17_Sello'] || [];
    for (const r of rows) {
      const foco = r.Foco || r.foco || '';
      if (typeof foco === 'string' && foco.includes('ELEGIDO')) {
        return {
          foco: foco.replace(' (ELEGIDO)', ''),
          texto: r.Texto || r.texto || '',
        };
      }
    }
    return { foco: '', texto: '' };
  })();

  const edges = computeEdges({
    lineas, sublineas, investigadores, temas,
    sublinea_tema, lab_linea, linea_modo, proximidad,
  });

  return {
    version: 'xlsx-direct',
    sello,
    lineas, sublineas, areas, modos, salidas, laboratorios,
    investigadores, temas,
    relations: {
      sublinea_tema, lab_linea, lab_salida,
      inv_lab, inv_modo, linea_modo, proximidad,
    },
    edges,
  };
}

// ===========================================================================
// Cómputo de aristas (a..g del spec)
// ===========================================================================

// Las 7 categorías de aristas se derivan en su totalidad de las relaciones
// m:n declaradas en el xlsx. Sin álgebra lineal, solo joins y set
// intersections. Llamado una vez por buildDataFromTabs().
function computeEdges({ lineas, sublineas, investigadores, temas,
                       sublinea_tema, lab_linea, linea_modo, proximidad }) {
  const edges = {
    jerarquica: [],
    coautoria: [],
    coinvestigacion: [],
    sosten_lab: [],
    afinidad_lab: [],
    coincidencia_modo: [],
    proximidad_semantica: [],
  };

  // (a) Jerárquica: cada Sublínea apunta a su Línea-madre.
  for (const s of sublineas) {
    if (s.linea) edges.jerarquica.push({ source: s.id, target: s.linea });
  }

  // (b) Coautoría: Investigador → Sublínea, vía la tabla puente Tema.
  const temaToInv = {};
  for (const t of temas) temaToInv[t.id] = t.investigador;
  const invToSublines = {};
  for (const st of sublinea_tema) {
    const invId = temaToInv[st.tema_id];
    if (!invId) continue;
    if (!invToSublines[invId]) invToSublines[invId] = new Set();
    invToSublines[invId].add(st.sublinea_id);
  }
  for (const [invId, subSet] of Object.entries(invToSublines)) {
    for (const subId of subSet) {
      edges.coautoria.push({ source: invId, target: subId });
    }
  }

  // (c) Coinvestigación: pares Sublínea↔Sublínea con al menos un Investigador
  //     en común. Set intersection sobre subToInvs.
  const subToInvs = {};
  for (const [invId, subSet] of Object.entries(invToSublines)) {
    for (const subId of subSet) {
      if (!subToInvs[subId]) subToInvs[subId] = new Set();
      subToInvs[subId].add(invId);
    }
  }
  const subIds = Object.keys(subToInvs).sort();
  for (let i = 0; i < subIds.length; i++) {
    for (let j = i + 1; j < subIds.length; j++) {
      const a = subIds[i], b = subIds[j];
      const shared = [...subToInvs[a]].filter(x => subToInvs[b].has(x));
      if (shared.length > 0) {
        edges.coinvestigacion.push({ source: a, target: b, weight: shared.length });
      }
    }
  }

  // (d) Sostén Lab: Laboratorio → Línea (copia directa de 10_Lab_Linea).
  for (const ll of lab_linea) {
    edges.sosten_lab.push({ source: ll.lab_id, target: ll.linea_id });
  }

  // (e) Afinidad por Lab: pares de Sublíneas cuyas Líneas-madre comparten
  //     al menos un Lab.
  const lineaToLabs = {};
  for (const ll of lab_linea) {
    if (!lineaToLabs[ll.linea_id]) lineaToLabs[ll.linea_id] = new Set();
    lineaToLabs[ll.linea_id].add(ll.lab_id);
  }
  const subToLabs = {};
  for (const s of sublineas) {
    if (s.linea && lineaToLabs[s.linea]) subToLabs[s.id] = lineaToLabs[s.linea];
  }
  const subIdsWithLab = Object.keys(subToLabs).sort();
  for (let i = 0; i < subIdsWithLab.length; i++) {
    for (let j = i + 1; j < subIdsWithLab.length; j++) {
      const a = subIdsWithLab[i], b = subIdsWithLab[j];
      if ([...subToLabs[a]].some(x => subToLabs[b].has(x))) {
        edges.afinidad_lab.push({ source: a, target: b });
      }
    }
  }

  // (f) Coincidencia Modo: pares de Sublíneas cuyas Líneas-madre comparten
  //     un Modo marcado como "predominante" en 14_Linea_Modo.
  const lineaToModos = {};
  for (const lm of linea_modo) {
    if ((lm.nivel || '').toLowerCase() === 'predominante') {
      if (!lineaToModos[lm.linea_id]) lineaToModos[lm.linea_id] = new Set();
      lineaToModos[lm.linea_id].add(lm.modo_id);
    }
  }
  const subToModos = {};
  for (const s of sublineas) {
    if (s.linea && lineaToModos[s.linea]) subToModos[s.id] = lineaToModos[s.linea];
  }
  const subIdsWithModo = Object.keys(subToModos).sort();
  for (let i = 0; i < subIdsWithModo.length; i++) {
    for (let j = i + 1; j < subIdsWithModo.length; j++) {
      const a = subIdsWithModo[i], b = subIdsWithModo[j];
      if ([...subToModos[a]].some(x => subToModos[b].has(x))) {
        edges.coincidencia_modo.push({ source: a, target: b });
      }
    }
  }

  // (g) Proximidad semántica declarada en 18_Proximidad_Tematica.
  for (const p of proximidad) {
    edges.proximidad_semantica.push({
      source: p.sublinea_a_id,
      target: p.sublinea_b_id,
      weight: p.afinidad,
    });
  }

  return edges;
}

// ===========================================================================
// API pública
// ===========================================================================

// Carga el xlsx desde una URL relativa al HTML (ej: './mad-map-data-v2.xlsx'),
// lo parsea con SheetJS y devuelve el DataModel que graph.js consume.
// Llamado desde graph.js#loadFromXlsx() y desde index.html para el sello.
async function loadFromXlsx(url) {
  if (typeof XLSX === 'undefined') {
    throw new Error('SheetJS (XLSX) no está cargado. Incluye xlsx.full.min.js antes de xlsx-loader.js.');
  }
  const res = await fetch(url);
  if (!res.ok) throw new Error(`xlsx fetch ${url}: HTTP ${res.status}`);
  const buffer = await res.arrayBuffer();
  const workbook = XLSX.read(buffer, { type: 'array' });
  const tabs = readAllSheets(workbook);
  return buildDataFromTabs(tabs);
}

// Compatibilidad: mantenemos el nombre del namespace para no romper graph.js.
window.MadMapDataLoader = {
  loadFromXlsx,
};
