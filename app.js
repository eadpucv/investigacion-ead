// Investigación e[ad] — visualización única: matriz de líneas × polos.
// Lee investigacion-ead.xlsx en el navegador (SheetJS). Sin compilación.
'use strict';

const XLSX_URL = './investigacion-ead.xlsx';
const POLOS = [
  { id: 'teórico', titulo: 'Polo teórico', nota: 'historia, teoría, crítica' },
  { id: 'proyectual', titulo: 'Polo proyectual', nota: 'obra, prototipo, intervención' },
];
const RECIENTE = 2022;

const $ = s => document.querySelector(s);
const esc = s => String(s ?? '').replace(/[&<>"]/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;' }[c]));
const norm = s => String(s ?? '').trim();
const fold = s => norm(s).toLowerCase().normalize('NFD').replace(/[̀-ͯ]/g, '');

const estado = { q: '', profesor: '', sel: null };
let D = null;

// ---------- carga ----------
async function cargar() {
  const res = await fetch(XLSX_URL, { cache: 'no-cache' });
  if (!res.ok) throw new Error(`No se pudo leer ${XLSX_URL} (HTTP ${res.status})`);
  const wb = XLSX.read(await res.arrayBuffer(), { type: 'array' });
  const hoja = n => wb.Sheets[n] ? XLSX.utils.sheet_to_json(wb.Sheets[n], { defval: '' }) : [];
  const avisos = [];

  const programa = Object.fromEntries(hoja('Programa').map(r => [norm(r.campo), norm(r.valor)]));
  const lineas = hoja('Líneas').filter(r => norm(r.nombre)).map(r => ({
    nombre: norm(r.nombre), modo: norm(r.modo), breve: norm(r['definición breve']),
    pregunta: norm(r.pregunta), prolonga: norm(r['prolonga (área del Magíster)']),
    labs: norm(r.laboratorios), descripcion: norm(r['descripción']),
  }));
  const lnames = new Set(lineas.map(l => l.nombre));

  const subs = hoja('Sublíneas').filter(r => norm(r.nombre)).map(r => ({
    nombre: norm(r.nombre), linea: norm(r['línea']), polo: fold(r.polo).startsWith('te') ? 'teórico' : 'proyectual',
    notas: norm(r.notas), profesores: new Map(), proyectos: [],
  }));
  const subBy = new Map(subs.map(s => [s.nombre, s]));
  subs.forEach(s => { if (!lnames.has(s.linea)) avisos.push(`Sublínea «${s.nombre}»: la línea «${s.linea}» no existe en la hoja Líneas.`); });

  const profs = hoja('Profesores').filter(r => norm(r.nombre)).map(r => ({
    nombre: norm(r.nombre), linea: norm(r['línea']), claustro: fold(r.claustro).startsWith('s'),
    perfil: norm(r.perfil), subs: new Map(), proyectos: [],
  }));
  const profBy = new Map(profs.map(p => [p.nombre, p]));
  profs.forEach(p => { if (p.linea && !lnames.has(p.linea)) avisos.push(`Profesor «${p.nombre}»: la línea «${p.linea}» no existe en la hoja Líneas.`); });

  hoja('Temas').forEach((r, i) => {
    const p = profBy.get(norm(r.profesor)), s = subBy.get(norm(r['sublínea']));
    if (!norm(r.profesor) && !norm(r['sublínea'])) return;
    if (!p) { avisos.push(`Temas, fila ${i + 2}: el profesor «${norm(r.profesor)}» no está en la hoja Profesores.`); return; }
    if (!s) { avisos.push(`Temas, fila ${i + 2}: la sublínea «${norm(r['sublínea'])}» no está en la hoja Sublíneas.`); return; }
    const tema = norm(r.tema);
    const lista = s.profesores.get(p.nombre) || []; if (tema) lista.push(tema); s.profesores.set(p.nombre, lista);
    const l2 = p.subs.get(s.nombre) || []; if (tema) l2.push(tema); p.subs.set(s.nombre, l2);
  });

  const proyectos = hoja('Proyectos').filter(r => norm(r['título'])).map((r, i) => {
    const s = subBy.get(norm(r['sublínea']));
    if (!s) avisos.push(`Proyectos, fila ${i + 2}: la sublínea «${norm(r['sublínea'])}» no está en la hoja Sublíneas.`);
    const nombres = norm(r.profesores).split(/[;,]/).map(norm).filter(Boolean);
    const p = { anio: parseInt(r['año'], 10) || null, titulo: norm(r['título']), profesores: nombres, sub: s || null, revisar: fold(r.revisar).startsWith('s') };
    if (s) s.proyectos.push(p);
    nombres.forEach(n => { const pr = profBy.get(n); if (pr) pr.proyectos.push(p); else avisos.push(`Proyectos, fila ${i + 2}: el profesor «${n}» no está en la hoja Profesores.`); });
    return p;
  });
  const porAnio = (a, b) => (b.anio || 0) - (a.anio || 0);
  subs.forEach(s => s.proyectos.sort(porAnio));
  profs.forEach(p => p.proyectos.sort(porAnio));
  return { programa, lineas, subs, subBy, profs, profBy, proyectos, avisos };
}

// ---------- búsqueda / filtro ----------
function coincide(s) {
  if (estado.profesor) {
    const p = D.profBy.get(estado.profesor);
    const enProy = s.proyectos.some(x => x.profesores.includes(estado.profesor));
    if (!(p && (p.subs.has(s.nombre) || enProy))) return false;
  }
  if (estado.q) {
    const q = fold(estado.q);
    const textos = [s.nombre, s.notas, ...s.profesores.keys(), ...[...s.profesores.values()].flat(), ...s.proyectos.map(p => p.titulo)];
    if (!textos.some(t => fold(t).includes(q))) return false;
  }
  return true;
}

// ---------- matriz ----------
function cifras(linea) {
  const S = D.subs.filter(s => s.linea === linea.nombre);
  const profs = new Set(); S.forEach(s => s.profesores.forEach((_, n) => profs.add(n)));
  const pr = new Set(); S.forEach(s => s.proyectos.forEach(p => pr.add(p)));
  const propios = D.profs.filter(p => p.linea === linea.nombre);
  const claustro = propios.filter(p => p.claustro).map(p => p.nombre);
  return { subs: S.length, profs: profs.size, propios: propios.map(p => p.nombre), proyectos: pr.size, claustro };
}

function pintarMatriz() {
  const max = Math.max(1, ...D.subs.map(s => s.proyectos.length));
  const filtro = !!(estado.q || estado.profesor);
  let visibles = 0;
  let html = '<div class="esquina"></div>';
  D.lineas.forEach((l, i) => {
    const c = cifras(l);
    html += `<section class="linea-cab l${i}" aria-label="${esc(l.nombre)}">
      <button type="button" class="linea-nombre" data-linea="${esc(l.nombre)}">${esc(l.nombre)}</button>
      <div class="linea-modo">${esc(l.modo)}</div>
      <p class="linea-breve">${esc(l.breve)}</p>
      <div class="linea-cifras">${c.subs} sublíneas · ${c.propios.length} profesores · ${c.proyectos} proyectos${c.claustro.length ? ` · claustro: ${c.claustro.length}` : ''}</div>
    </section>`;
  });
  POLOS.forEach(polo => {
    html += `<div class="polo-cab"><span class="polo-icono ${polo.id === 'teórico' ? 'teo' : 'pro'}"></span><span class="polo-titulo">${polo.titulo}</span><span class="polo-nota">${polo.nota}</span></div>`;
    D.lineas.forEach((l, i) => {
      const S = D.subs.filter(s => s.linea === l.nombre && s.polo === polo.id)
        .sort((a, b) => b.proyectos.length - a.proyectos.length || a.nombre.localeCompare(b.nombre, 'es'));
      html += `<div class="celda l${i} ${polo.id === 'teórico' ? 'teo' : 'pro'}"><div class="celda-polo">${polo.titulo}</div>`;
      if (!S.length) html += '<p class="vacio">Sin sublíneas</p>';
      S.forEach(s => {
        const ok = coincide(s); if (ok) visibles++;
        const cls = ['tarjeta', filtro && !ok ? 'apagada' : '', filtro && ok ? 'encendida' : '', estado.sel?.tipo === 'sub' && estado.sel.id === s.nombre ? 'elegida' : '', s.notas ? 'con-nota' : ''].join(' ');
        html += `<button type="button" class="${cls}" data-sub="${esc(s.nombre)}">
          <span class="t-nombre">${esc(s.nombre)}</span>
          <span class="t-meta">${s.profesores.size} p · ${s.proyectos.length} pr${s.notas ? ' · <i>nota</i>' : ''}</span>
          <span class="t-barra" style="width:${Math.round(100 * s.proyectos.length / max)}%"></span>
        </button>`;
      });
      html += '</div>';
    });
  });
  $('#matriz').innerHTML = html;
  $('#estado').textContent = filtro ? `${visibles} de ${D.subs.length} sublíneas` : '';
  $('#limpiar').hidden = !filtro;
}

// ---------- panel ----------
const listaProyectos = (ps, conProfes = true) => !ps.length ? '<p class="nada">Sin proyectos registrados.</p>' :
  `<ul class="proyectos">${ps.map(p => `<li><span class="anio">${p.anio || '—'}</span><span class="pt">${esc(p.titulo)}${conProfes && p.profesores.length ? `<span class="pp">${p.profesores.map(n => profLink(n)).join(', ')}</span>` : ''}${p.revisar ? ' <span class="rev" title="Asignación por revisar">por revisar</span>' : ''}</span></li>`).join('')}</ul>`;
const profLink = n => D.profBy.has(n) ? `<button type="button" class="enlace" data-prof="${esc(n)}">${esc(n)}</button>` : esc(n);
const subLink = n => `<button type="button" class="enlace" data-sub="${esc(n)}">${esc(n)}</button>`;

function panelSub(s) {
  const recientes = s.proyectos.filter(p => (p.anio || 0) >= RECIENTE).length;
  return `<div class="p-tipo">Sublínea</div><h2>${esc(s.nombre)}</h2>
    <div class="p-ubica"><button type="button" class="enlace" data-linea="${esc(s.linea)}">${esc(s.linea)}</button> · polo ${s.polo}</div>
    ${s.notas ? `<p class="p-nota">${esc(s.notas)}</p>` : ''}
    <h3>Profesores (${s.profesores.size})</h3>
    ${s.profesores.size ? `<ul class="profes">${[...s.profesores].map(([n, temas]) => `<li>${profLink(n)}${temas.length ? `<span class="tema">${esc(temas.join(' · '))}</span>` : ''}</li>`).join('')}</ul>` : '<p class="nada">Ningún profesor declarado.</p>'}
    <h3>Proyectos 2014-2026 (${s.proyectos.length}${s.proyectos.length ? `; ${recientes} desde ${RECIENTE}` : ''})</h3>
    ${listaProyectos(s.proyectos)}`;
}

function panelProf(p) {
  const porLinea = new Map();
  p.subs.forEach((temas, n) => { const s = D.subBy.get(n); const k = s ? s.linea : '—'; (porLinea.get(k) || porLinea.set(k, []).get(k)).push(s); });
  return `<div class="p-tipo">Profesor${p.claustro ? ' · claustro' : ''}</div><h2>${esc(p.nombre)}</h2>
    ${p.linea ? `<div class="p-ubica">${p.claustro ? 'Claustro de' : 'Aporta a'} <button type="button" class="enlace" data-linea="${esc(p.linea)}">${esc(p.linea)}</button></div>` : ''}
    ${p.perfil ? `<p><a href="${esc(p.perfil)}" target="_blank" rel="noopener">Perfil en Casiopea ↗</a></p>` : ''}
    <h3>Sublíneas que cultiva (${p.subs.size})</h3>
    ${[...porLinea].map(([ln, ss]) => `<div class="grupo"><div class="grupo-t">${esc(ln)}</div><ul class="subs">${ss.filter(Boolean).map(s => `<li>${subLink(s.nombre)} <span class="polo-mini">${s.polo}</span>${(p.subs.get(s.nombre) || []).length ? `<span class="tema">${esc(p.subs.get(s.nombre).join(' · '))}</span>` : ''}</li>`).join('')}</ul></div>`).join('') || '<p class="nada">Sin sublíneas declaradas.</p>'}
    <h3>Proyectos 2014-2026 (${p.proyectos.length})</h3>
    ${listaProyectos(p.proyectos, false)}`;
}

function panelLinea(l) {
  const c = cifras(l);
  return `<div class="p-tipo">Línea de investigación</div><h2>${esc(l.nombre)}</h2>
    <div class="p-modo">${esc(l.modo)}</div>
    <p class="p-breve">${esc(l.breve)}</p>
    ${l.pregunta ? `<h3>Pregunta</h3><p>${esc(l.pregunta)}</p>` : ''}
    ${l.descripcion ? `<h3>Alcance</h3><p>${esc(l.descripcion)}</p>` : ''}
    <h3>En cifras</h3><p>${c.subs} sublíneas · ${c.proyectos} proyectos 2014-2026 · ${c.profs} profesores cultivan alguna de sus sublíneas</p>
    ${c.claustro.length ? `<h3>Claustro (${c.claustro.length})</h3><ul class="profes">${c.claustro.map(n => `<li>${profLink(n)}</li>`).join('')}</ul>` : ''}
    ${c.propios.length ? `<h3>Profesores que aportan a la línea (${c.propios.length})</h3><p class="profes-linea">${c.propios.map(n => profLink(n)).join(', ')}</p>` : ''}
    ${l.prolonga ? `<h3>Prolonga</h3><p>El área ${esc(l.prolonga)} del Magíster.</p>` : ''}
    ${l.labs ? `<h3>Laboratorios</h3><p>${esc(l.labs)}</p>` : ''}`;
}

function abrir(tipo, id) {
  let html = '';
  if (tipo === 'sub') { const s = D.subBy.get(id); if (!s) return; html = panelSub(s); }
  if (tipo === 'prof') { const p = D.profBy.get(id); if (!p) return; html = panelProf(p); }
  if (tipo === 'linea') { const l = D.lineas.find(x => x.nombre === id); if (!l) return; html = panelLinea(l); }
  estado.sel = { tipo, id };
  $('#panel-contenido').innerHTML = html;
  $('#panel').classList.add('abierto'); $('#panel').setAttribute('aria-hidden', 'false');
  $('#velo').classList.add('visible');
  document.body.classList.add('con-panel');
  $('#panel').scrollTop = 0;
  history.replaceState(null, '', '#' + tipo + '=' + encodeURIComponent(id));
  pintarMatriz();
}
function cerrar() {
  estado.sel = null;
  $('#panel').classList.remove('abierto'); $('#panel').setAttribute('aria-hidden', 'true');
  $('#velo').classList.remove('visible');
  document.body.classList.remove('con-panel');
  history.replaceState(null, '', location.pathname);
  pintarMatriz();
}

// ---------- eventos ----------
function cablear() {
  document.addEventListener('click', e => {
    const b = e.target.closest('[data-sub],[data-prof],[data-linea]');
    if (!b) return;
    if (b.dataset.sub) abrir('sub', b.dataset.sub);
    else if (b.dataset.prof) { estado.profesor = b.dataset.prof; $('#profesor').value = b.dataset.prof; abrir('prof', b.dataset.prof); }
    else if (b.dataset.linea) abrir('linea', b.dataset.linea);
  });
  $('#cerrar').addEventListener('click', cerrar);
  $('#velo').addEventListener('click', cerrar);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') cerrar(); });
  $('#buscar').addEventListener('input', e => { estado.q = e.target.value; pintarMatriz(); });
  $('#profesor').addEventListener('change', e => {
    estado.profesor = e.target.value;
    if (estado.profesor) abrir('prof', estado.profesor); else { cerrar(); }
  });
  $('#limpiar').addEventListener('click', () => {
    estado.q = ''; estado.profesor = ''; $('#buscar').value = ''; $('#profesor').value = ''; cerrar();
  });
}

(async function iniciar() {
  try {
    D = await cargar();
  } catch (err) {
    $('#matriz').innerHTML = `<p class="error">No se pudo cargar la planilla: ${esc(err.message)}</p>`;
    console.error(err); return;
  }
  const pg = D.programa;
  $('#pregunta').textContent = pg['pregunta'] || '';
  $('#sello-titulo').textContent = pg['sello (título)'] || 'El sello del programa';
  $('#sello-texto').textContent = pg['sello (texto)'] || '';
  const sel = $('#profesor');
  const grupos = [...D.lineas.map(l => l.nombre), ''];
  grupos.forEach(g => {
    const ps = D.profs.filter(p => (p.linea || '') === g).sort((a, b) => a.nombre.localeCompare(b.nombre, 'es'));
    if (!ps.length) return;
    sel.insertAdjacentHTML('beforeend', `<optgroup label="${esc(g || 'Sin línea asignada')}">${ps.map(p => `<option value="${esc(p.nombre)}">${esc(p.nombre)}${p.claustro ? ' · claustro' : ''}</option>`).join('')}</optgroup>`);
  });
  if (D.avisos.length) {
    const a = $('#avisos'); a.hidden = false;
    a.innerHTML = `<details><summary>${D.avisos.length} aviso(s) de la planilla</summary><ul>${D.avisos.map(x => `<li>${esc(x)}</li>`).join('')}</ul></details>`;
  }
  cablear();
  pintarMatriz();
  const desdeHash = () => {
    const m = location.hash.match(/^#(sub|prof|linea)=(.+)$/);
    if (!m) return;
    const id = decodeURIComponent(m[2]);
    if (m[1] === 'prof') { estado.profesor = id; sel.value = id; }
    abrir(m[1], id);
  };
  window.addEventListener('hashchange', desdeHash);
  desdeHash();
})();
