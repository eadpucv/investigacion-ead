// Investigación e[ad] — visualización única: grafo de fuerzas con envolventes anidadas
// (líneas › polos › ejes temáticos › personas).
// Lee investigacion-ead.xlsx en el navegador (SheetJS). Sin compilación.
'use strict';

const XLSX_URL = './investigacion-ead.xlsx';
const POLOS = [
  { id: 'teórico', titulo: 'Polo teórico' },
  { id: 'proyectual', titulo: 'Polo proyectual' },
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

// ---------- vista: grafo de fuerzas con envolventes anidadas ----------
// líneas (envolvente de color) › polos (envolvente punteada) › ejes temáticos
// (envolvente blanca) › personas (círculos). Una persona que cultiva varios
// ejes aparece en cada uno; al pasar sobre ella se iluminan todas sus apariciones.
const COLOR = ['#a3121a', '#0d5c68'];
const TINTE = ['#f6e7e6', '#e3eff0'];
const RP = 9;                       // radio de una persona
const V = { svg: null, g: null, zoom: null, nodos: [], ejes: [], w: 0, h: 0, t: d3.zoomIdentity };

const lineaIdx = n => Math.max(0, D.lineas.findIndex(l => l.nombre === n));

function construirNodos() {
  const nodos = [], ejes = [];
  D.lineas.forEach((l, li) => {
    POLOS.forEach((po, pi) => {
      D.subs.filter(s => s.linea === l.nombre && s.polo === po.id).forEach(s => {
        const e = { sub: s, li, pi, miembros: [] };
        s.profesores.forEach((temas, n) => {
          const p = D.profBy.get(n);
          const nd = { tipo: 'persona', nombre: n, prof: p, eje: e, li, pi, pli: p && p.linea ? lineaIdx(p.linea) : -1, r: RP };
          e.miembros.push(nd); nodos.push(nd);
        });
        if (!e.miembros.length) { const nd = { tipo: 'vacio', eje: e, li, pi, r: RP * 0.7 }; e.miembros.push(nd); nodos.push(nd); }
        ejes.push(e);
      });
    });
  });
  return { nodos, ejes };
}

// Fuerza propia: cada persona es atraída al centroide de su eje, y los ejes
// se repelen entre sí como burbujas (más separación entre polos y líneas).
function fuerzaEjes(ejes) {
  let nodes;
  function force(alpha) {
    ejes.forEach(e => {
      let x = 0, y = 0; e.miembros.forEach(m => { x += m.x; y += m.y; });
      e.x = x / e.miembros.length; e.y = y / e.miembros.length;
      e.R = Math.sqrt(e.miembros.length) * RP * 1.35 + RP;
      e.miembros.forEach(m => { m.vx += (e.x - m.x) * 0.22 * alpha; m.vy += (e.y - m.y) * 0.22 * alpha; });
    });
    for (let i = 0; i < ejes.length; i++) for (let j = i + 1; j < ejes.length; j++) {
      const a = ejes[i], b = ejes[j];
      const gap = a.li !== b.li ? 90 : a.pi !== b.pi ? 40 : 9;
      let dx = b.x - a.x, dy = b.y - a.y, d = Math.hypot(dx, dy) || 1;
      const min = a.R + b.R + gap;
      if (d < min) {
        const f = (min - d) / d * 0.5 * alpha;
        dx *= f; dy *= f;
        a.miembros.forEach(m => { m.vx -= dx; m.vy -= dy; });
        b.miembros.forEach(m => { m.vx += dx; m.vy += dy; });
      }
    }
  }
  force.initialize = n => { nodes = n; };
  return force;
}

// Contorno suave alrededor de un conjunto de círculos (con margen).
function contorno(puntos, pad) {
  const pts = [];
  puntos.forEach(p => { const r = (p.r || RP) + pad; for (let a = 0; a < 16; a++) pts.push([p.x + r * Math.cos(a * Math.PI / 8), p.y + r * Math.sin(a * Math.PI / 8)]); });
  const h = d3.polygonHull(pts);
  return h ? d3.line().curve(d3.curveCatmullRomClosed.alpha(0.6))(h) : null;
}

function montarVista() {
  const cont = $('#mapa');
  cont.innerHTML = '<svg role="img" aria-label="Líneas, ejes temáticos y personas"></svg><div class="mapa-ayuda">Pasa el cursor para ver nombres · clic para el detalle · rueda o pellizco para acercar</div>';
  V.w = cont.clientWidth; V.h = cont.clientHeight;
  const { nodos, ejes } = construirNodos();
  V.nodos = nodos; V.ejes = ejes;
  // Cada línea se simula por separado y se ubica en su mitad del lienzo,
  // a la misma escala: así sus envolventes nunca se cruzan.
  // Escritorio: líneas lado a lado y polos arriba (teórico) / abajo (proyectual).
  // Pantalla angosta: líneas una sobre otra y polos a izquierda / derecha.
  const vertical = V.w < 700;
  const grupos = D.lineas.map((l, li) => ({ li, nodos: nodos.filter(n => n.li === li), ejes: ejes.filter(e => e.li === li) }));
  grupos.forEach(g => {
    const sim = d3.forceSimulation(g.nodos)
      .force('x', d3.forceX(d => vertical ? (d.pi === 0 ? -150 : 150) : 0).strength(vertical ? 0.1 : 0.06))
      .force('y', d3.forceY(d => vertical ? 0 : (d.pi === 0 ? -110 : 110)).strength(vertical ? 0.06 : 0.1))
      .force('choque', d3.forceCollide(d => d.r + 1.5).strength(0.9))
      .force('ejes', fuerzaEjes(g.ejes))
      .stop();
    for (let i = 0; i < 500; i++) sim.tick();
    g.x0 = d3.min(g.nodos, n => n.x - n.r); g.x1 = d3.max(g.nodos, n => n.x + n.r);
    g.y0 = d3.min(g.nodos, n => n.y - n.r); g.y1 = d3.max(g.nodos, n => n.y + n.r);
  });
  const M = 62, TOP = 46;
  const caja = li => vertical ? { x: 0, y: li * V.h / 2, w: V.w, h: V.h / 2 } : { x: li * V.w / 2, y: 0, w: V.w / 2, h: V.h };
  const escala = Math.min(1.6, ...grupos.map(g => { const c = caja(g.li); return Math.min((c.w - 2 * M) / (g.x1 - g.x0), (c.h - 2 * M - TOP) / (g.y1 - g.y0)); }));
  grupos.forEach(g => {
    const c = caja(g.li), gw = (g.x1 - g.x0) * escala, gh = (g.y1 - g.y0) * escala;
    const ox = c.x + (c.w - gw) / 2, oy = c.y + TOP + (c.h - TOP - gh) / 2;
    g.nodos.forEach(n => { n.x = ox + (n.x - g.x0) * escala; n.y = oy + (n.y - g.y0) * escala; n.r *= Math.min(1.25, Math.max(0.75, escala)); });
  });

  V.svg = d3.select(cont).select('svg').attr('viewBox', [0, 0, V.w, V.h]);
  V.g = V.svg.append('g');
  const gl = V.g.append('g').attr('class', 'capa-lineas');
  const gp = V.g.append('g').attr('class', 'capa-polos');
  const ge = V.g.append('g').attr('class', 'capa-ejes');
  const gu = V.g.append('g').attr('class', 'capa-uniones');
  const gn = V.g.append('g').attr('class', 'capa-personas');
  const gt = V.g.append('g').attr('class', 'capa-titulos');

  D.lineas.forEach((l, li) => {
    const m = nodos.filter(n => n.li === li);
    gl.append('path').attr('class', 'env-linea').attr('d', contorno(m, 46)).attr('fill', TINTE[li]).attr('stroke', COLOR[li])
      .datum({ tipo: 'linea', nombre: l.nombre, li });
    POLOS.forEach((po, pi) => {
      const mp = m.filter(n => n.pi === pi);
      if (mp.length) gp.append('path').attr('class', 'env-polo ' + (pi === 0 ? 'teo' : 'pro')).attr('d', contorno(mp, 24)).attr('stroke', COLOR[li])
        .datum({ tipo: 'polo', nombre: `${po.titulo} · ${l.nombre}`, li });
    });
    const top = d3.min(m, n => n.y - n.r) - 46, cx = d3.mean(m, n => n.x);
    gt.append('text').attr('class', 'titulo-linea').attr('x', cx).attr('y', top - 14).attr('fill', COLOR[li]).text(l.nombre);
  });
  V.envEjes = ge.selectAll('path').data(ejes).join('path').attr('class', 'env-eje').attr('stroke', d => COLOR[d.li])
    .attr('d', d => contorno(d.miembros, 6));
  V.uniones = gu;
  V.circ = gn.selectAll('circle').data(nodos).join('circle')
    .attr('class', d => `per ${d.tipo === 'vacio' ? 'vacio' : ''} ${d.prof && d.prof.claustro ? 'claustro' : ''}`)
    .attr('cx', d => d.x).attr('cy', d => d.y).attr('r', d => d.r)
    .attr('fill', d => d.tipo === 'vacio' ? 'none' : d.pli >= 0 ? COLOR[d.pli] : '#888')
    .attr('stroke', d => d.tipo === 'vacio' ? COLOR[d.li] : null);

  // interacción
  const sobre = (ev, d) => { tooltip(ev, d); resaltar(d); };
  const fuera = () => { $('#tip').hidden = true; resaltar(null); };
  V.svg.selectAll('.env-linea, .env-polo, .env-eje, circle.per')
    .on('mouseenter', sobre).on('mousemove', ev => moverTooltip(ev)).on('mouseleave', fuera)
    .on('click', (ev, d) => { ev.stopPropagation(); clicNodo(d); });
  V.zoom = d3.zoom().scaleExtent([0.6, 8]).on('zoom', ev => { V.t = ev.transform; V.g.attr('transform', ev.transform); });
  V.svg.call(V.zoom).on('dblclick.zoom', null);
  V.svg.on('click', () => cerrar());
}

function enfocar() { V.svg.transition().duration(500).call(V.zoom.transform, d3.zoomIdentity); }

function clicNodo(d) {
  if (d.tipo === 'persona') { estado.profesor = d.nombre; $('#profesor').value = d.nombre; abrir('prof', d.nombre); return; }
  if (d.tipo === 'vacio') { abrir('sub', d.eje.sub.nombre); return; }
  if (d.sub) { abrir('sub', d.sub.nombre); return; }
  if (d.tipo === 'linea') { abrir('linea', d.nombre); return; }
}

// Al pasar sobre una persona: todas sus apariciones y las uniones entre ellas.
function resaltar(d) {
  V.uniones.selectAll('*').remove();
  V.circ.classed('hover', false); V.envEjes.classed('hover', false);
  if (!d) return;
  if (d.tipo === 'persona') {
    const inst = V.nodos.filter(n => n.nombre === d.nombre);
    V.circ.classed('hover', n => n.nombre === d.nombre);
    V.envEjes.classed('hover', e => inst.some(n => n.eje === e));
    for (let i = 1; i < inst.length; i++) V.uniones.append('line').attr('x1', inst[0].x).attr('y1', inst[0].y).attr('x2', inst[i].x).attr('y2', inst[i].y);
  } else if (d.sub) {
    V.envEjes.classed('hover', e => e === d);
  }
}

function tooltip(ev, d) {
  const tip = $('#tip'); let h = '';
  if (d.tipo === 'persona') {
    const temas = d.eje.sub.profesores.get(d.nombre) || [];
    h = `<b>${esc(d.nombre)}</b>${d.prof && d.prof.claustro ? ' · claustro' : ''}<br><span class="tip-sub">${esc(d.eje.sub.nombre)}</span>${temas.length ? `<br><i>${esc(temas.join(' · '))}</i>` : ''}`;
  } else if (d.tipo === 'vacio' || d.sub) {
    const s = (d.sub || d.eje.sub);
    h = `<b>${esc(s.nombre)}</b><br><span class="tip-sub">${s.profesores.size} personas · ${s.proyectos.length} proyectos · polo ${s.polo}</span>`;
  } else if (d.tipo === 'linea') { const l = D.lineas[d.li]; h = `<b>${esc(l.nombre)}</b><br><span class="tip-sub">${esc(l.modo)}</span>`; }
  else h = `<b>${esc(d.nombre)}</b>`;
  tip.innerHTML = h; tip.hidden = false; moverTooltip(ev);
}
function moverTooltip(ev) {
  const r = $('#mapa').getBoundingClientRect(), tip = $('#tip');
  const x = ev.clientX - r.left + 14, y = ev.clientY - r.top + 14;
  tip.style.left = Math.min(x, r.width - 300) + 'px'; tip.style.top = y + 'px';
}

function cifras(linea) {
  const S = D.subs.filter(s => s.linea === linea.nombre);
  const profs = new Set(); S.forEach(s => s.profesores.forEach((_, n) => profs.add(n)));
  const pr = new Set(); S.forEach(s => s.proyectos.forEach(p => pr.add(p)));
  const propios = D.profs.filter(p => p.linea === linea.nombre);
  const claustro = propios.filter(p => p.claustro).map(p => p.nombre);
  return { subs: S.length, profs: profs.size, propios: propios.map(p => p.nombre), proyectos: pr.size, claustro };
}

// Estado visual: búsqueda, persona elegida y selección.
function pintar() {
  if (!V.circ) return;
  const filtro = !!(estado.q || estado.profesor);
  const q = fold(estado.q);
  const okEje = new Map(V.ejes.map(e => [e, coincide(e.sub)]));
  V.envEjes.classed('apagado', e => filtro && !okEje.get(e))
    .classed('elegido', e => estado.sel && estado.sel.tipo === 'sub' && estado.sel.id === e.sub.nombre);
  V.circ.classed('apagado', d => {
    if (!filtro) return false;
    if (d.tipo === 'vacio') return !okEje.get(d.eje);
    if (estado.profesor) return d.nombre !== estado.profesor;
    return !(fold(d.nombre).includes(q) || okEje.get(d.eje));
  }).classed('encendido', d => !!estado.profesor && d.nombre === estado.profesor);
  const visibles = V.ejes.filter(e => okEje.get(e)).length;
  $('#estado').textContent = filtro ? `${visibles} de ${D.subs.length} ejes temáticos` : '';
  $('#limpiar').hidden = !filtro;
}

// ---------- panel ----------
const listaProyectos = (ps, conProfes = true) => !ps.length ? '<p class="nada">Sin proyectos registrados.</p>' :
  `<ul class="proyectos">${ps.map(p => `<li><span class="anio">${p.anio || '—'}</span><span class="pt">${esc(p.titulo)}${conProfes && p.profesores.length ? `<span class="pp">${p.profesores.map(n => profLink(n)).join(', ')}</span>` : ''}${p.revisar ? ' <span class="rev" title="Asignación por revisar">por revisar</span>' : ''}</span></li>`).join('')}</ul>`;
const profLink = n => D.profBy.has(n) ? `<button type="button" class="enlace" data-prof="${esc(n)}">${esc(n)}</button>` : esc(n);
const subLink = n => `<button type="button" class="enlace" data-sub="${esc(n)}">${esc(n)}</button>`;

function panelSub(s) {
  const recientes = s.proyectos.filter(p => (p.anio || 0) >= RECIENTE).length;
  return `<div class="p-tipo">Eje temático</div><h2>${esc(s.nombre)}</h2>
    <div class="p-ubica"><button type="button" class="enlace" data-linea="${esc(s.linea)}">${esc(s.linea)}</button> · polo ${s.polo}</div>
    ${s.notas ? `<p class="p-nota">${esc(s.notas)}</p>` : ''}
    <h3>Personas (${s.profesores.size})</h3>
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
    <h3>Ejes temáticos que cultiva (${p.subs.size})</h3>
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
    <h3>En cifras</h3><p>${c.subs} ejes temáticos · ${c.proyectos} proyectos 2014-2026 · ${c.profs} personas cultivan alguno de sus ejes</p>
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
  pintar();
}
function cerrar() {
  estado.sel = null;
  $('#panel').classList.remove('abierto'); $('#panel').setAttribute('aria-hidden', 'true');
  $('#velo').classList.remove('visible');
  document.body.classList.remove('con-panel');
  history.replaceState(null, '', location.pathname);
  pintar();
}

// ---------- eventos ----------
function cablear() {
  document.addEventListener('click', e => {
    if (e.target.closest('#mapa svg')) return;
    const b = e.target.closest('[data-sub],[data-prof],[data-linea]');
    if (!b) return;
    if (b.dataset.sub) abrir('sub', b.dataset.sub);
    else if (b.dataset.prof) { estado.profesor = b.dataset.prof; $('#profesor').value = b.dataset.prof; abrir('prof', b.dataset.prof); }
    else if (b.dataset.linea) abrir('linea', b.dataset.linea);
  });
  $('#cerrar').addEventListener('click', cerrar);
  $('#velo').addEventListener('click', cerrar);
  document.addEventListener('keydown', e => { if (e.key === 'Escape') cerrar(); });
  $('#buscar').addEventListener('input', e => { estado.q = e.target.value; pintar(); });
  $('#profesor').addEventListener('change', e => {
    estado.profesor = e.target.value;
    if (estado.profesor) { abrir('prof', estado.profesor); } else { cerrar(); }
  });
  $('#limpiar').addEventListener('click', () => {
    estado.q = ''; estado.profesor = ''; $('#buscar').value = ''; $('#profesor').value = ''; cerrar();
  });
}

(async function iniciar() {
  try {
    D = await cargar();
  } catch (err) {
    $('#mapa').innerHTML = `<p class="error">No se pudo cargar la planilla: ${esc(err.message)}</p>`;
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
  montarVista();
  pintar();
  let rz; window.addEventListener('resize', () => { clearTimeout(rz); rz = setTimeout(() => { montarVista(); pintar(); }, 250); });
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
