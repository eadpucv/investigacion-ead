// graph.js — visualización compartida por las 3 superficies de Investigación e[ad].
//
// Layout: D3 force-directed. Las posiciones se recalculan al vuelo desde las
// aristas leídas en mad-map-data-v2.xlsx (la fuente única de verdad). No hay
// embedding ni posiciones bakeadas; el layout es organico por construcción.
//
// Uso desde un HTML:
//   const graph = new MadMapGraph({
//     surface: 'cartografia' | 'narrativa' | 'exploracion',
//     allowedEdges: ['jerarquica', 'proximidad_semantica'],
//     defaultEdges: ['jerarquica', 'proximidad_semantica'],
//     allowPerfiles: false,
//     defaultEnvolventes: ['area'],
//     allowedFilters: ['linea','area','modo'],   // qué filtros mostrar
//   });
//   graph.loadFromXlsx('./mad-map-data-v2.xlsx');

class MadMapGraph {
  constructor(config) {
    this.config = Object.assign({
      surface: 'exploracion',
      allowedEdges: ['jerarquica','coautoria','coinvestigacion','sosten_lab',
                     'afinidad_lab','coincidencia_modo','proximidad_semantica'],
      defaultEdges: ['jerarquica'],
      allowPerfiles: true,
      defaultShowPerfiles: false,
      defaultEnvolventes: ['area'],
      allowedFilters: ['linea','area','modo','salida','lab','investigador'],
    }, config);

    this.data = null;
    this.state = {
      activeEdges: new Set(this.config.defaultEdges),
      activeEnvolventes: new Set(this.config.defaultEnvolventes),
      showPerfiles: this.config.defaultShowPerfiles,
      filters: {
        linea: new Set(), area: new Set(), modo: new Set(),
        salida: new Set(), lab: new Set(), investigador: new Set(),
      },
      searchQuery: '',
      selectedNode: null,
    };

    this.svg = null;
    this.simulation = null;
    this.zoomBehavior = null;
    this.nodes = [];
    this.edges = [];
  }

  // Carga el modelo desde un archivo .xlsx local (fuente única de verdad).
  // Requiere que SheetJS y xlsx-loader.js esten cargados antes que graph.js.
  async loadFromXlsx(xlsxUrl) {
    this.data = await window.MadMapDataLoader.loadFromXlsx(xlsxUrl);
    this._initWithData();
  }

  _initWithData() {
    this._setupDOM();
    this._buildNodesAndEdges();
    this._renderControls();
    this._populateFilters();
    this._initSimulation();
    this._render();
    this._hideLoading();
  }

  _setupDOM() {
    this.svgEl = document.querySelector('main.canvas svg');
    this.svg = d3.select(this.svgEl);
    this.tooltip = d3.select('.tooltip');
    this.detailPanel = document.querySelector('.detail-panel');
    this.emptyState = document.querySelector('.empty-state');

    // Grupos en orden de profundidad (de atrás hacia adelante)
    this.gZoom = this.svg.append('g').attr('class', 'zoom-root');
    this.gEnvolventes = this.gZoom.append('g').attr('class', 'envolventes');
    this.gEdges = this.gZoom.append('g').attr('class', 'edges');
    this.gNodes = this.gZoom.append('g').attr('class', 'nodes');

    // Zoom y pan
    this.zoomBehavior = d3.zoom()
      .scaleExtent([0.2, 5])
      .on('zoom', (event) => {
        this.gZoom.attr('transform', event.transform);
      });
    this.svg.call(this.zoomBehavior);

    // Click fuera de un nodo cierra el panel de detalle
    this.svg.on('click', (event) => {
      if (event.target === this.svgEl) {
        this._selectNode(null);
      }
    });
  }

  _buildNodesAndEdges() {
    const d = this.data;
    const nodes = [];
    const nodeIndex = {};

    // Líneas
    for (const l of d.lineas) {
      const n = { id: l.id, kind: 'linea', data: l, label: l.nombre };
      nodes.push(n); nodeIndex[l.id] = n;
    }
    // Sublíneas
    for (const s of d.sublineas) {
      const n = { id: s.id, kind: 'sublinea', data: s, label: s.nombre };
      nodes.push(n); nodeIndex[s.id] = n;
    }
    // Investigadores
    for (const i of d.investigadores) {
      const n = { id: i.id, kind: 'investigador', data: i, label: i.nombre };
      nodes.push(n); nodeIndex[i.id] = n;
    }

    // Aristas (conservamos sólo las del tipo permitido por la superficie)
    const edges = [];
    for (const kind of this.config.allowedEdges) {
      for (const e of (d.edges[kind] || [])) {
        const sn = nodeIndex[e.source];
        const tn = nodeIndex[e.target];
        if (sn && tn) {
          edges.push({ source: sn, target: tn, kind, weight: e.weight ?? 1 });
        }
      }
    }

    // Construir mapas auxiliares para filtros e info de detalle
    const subToInvs = {};
    for (const e of d.edges.coautoria || []) {
      if (!subToInvs[e.target]) subToInvs[e.target] = new Set();
      subToInvs[e.target].add(e.source);
    }
    const invToSubs = {};
    for (const e of d.edges.coautoria || []) {
      if (!invToSubs[e.source]) invToSubs[e.source] = new Set();
      invToSubs[e.source].add(e.target);
    }
    const labToLineas = {};
    const lineaToLabs = {};
    for (const r of d.relations.lab_linea || []) {
      if (!labToLineas[r.lab_id]) labToLineas[r.lab_id] = new Set();
      labToLineas[r.lab_id].add(r.linea_id);
      if (!lineaToLabs[r.linea_id]) lineaToLabs[r.linea_id] = new Set();
      lineaToLabs[r.linea_id].add(r.lab_id);
    }
    const labToSalidas = {};
    for (const r of d.relations.lab_salida || []) {
      if (!labToSalidas[r.lab_id]) labToSalidas[r.lab_id] = new Set();
      labToSalidas[r.lab_id].add(r.salida_id);
    }
    const lineaToModos = {};
    for (const r of d.relations.linea_modo || []) {
      if (!lineaToModos[r.linea_id]) lineaToModos[r.linea_id] = new Set();
      lineaToModos[r.linea_id].add(r.modo_id);
    }
    const invToModos = {};
    for (const r of d.relations.inv_modo || []) {
      if (!invToModos[r.investigador_id]) invToModos[r.investigador_id] = new Set();
      invToModos[r.investigador_id].add(r.modo_id);
    }
    const invToLabs = {};
    for (const r of d.relations.inv_lab || []) {
      if (!invToLabs[r.investigador_id]) invToLabs[r.investigador_id] = new Set();
      invToLabs[r.investigador_id].add(r.lab_id);
    }

    // Anotar cada nodo con sus categorías para filtrado y envolventes
    for (const n of nodes) {
      n.lineas = new Set();
      n.areas = new Set();
      n.modos = new Set();
      n.salidas = new Set();
      n.labs = new Set();
      n.investigadores = new Set();

      if (n.kind === 'linea') {
        n.lineas.add(n.id);
        if (lineaToModos[n.id]) lineaToModos[n.id].forEach(m => n.modos.add(m));
        if (lineaToLabs[n.id]) lineaToLabs[n.id].forEach(l => {
          n.labs.add(l);
          if (labToSalidas[l]) labToSalidas[l].forEach(s => n.salidas.add(s));
        });
      } else if (n.kind === 'sublinea') {
        n.lineas.add(n.data.linea);
        n.areas.add(n.data.area);
        if (lineaToModos[n.data.linea]) lineaToModos[n.data.linea].forEach(m => n.modos.add(m));
        if (lineaToLabs[n.data.linea]) {
          lineaToLabs[n.data.linea].forEach(l => {
            n.labs.add(l);
            if (labToSalidas[l]) labToSalidas[l].forEach(s => n.salidas.add(s));
          });
        }
        if (subToInvs[n.id]) subToInvs[n.id].forEach(i => n.investigadores.add(i));
      } else if (n.kind === 'investigador') {
        n.investigadores.add(n.id);
        n.areas.add(n.data.area_principal);
        if (invToModos[n.id]) invToModos[n.id].forEach(m => n.modos.add(m));
        if (invToLabs[n.id]) {
          invToLabs[n.id].forEach(l => {
            n.labs.add(l);
            if (labToLineas[l]) labToLineas[l].forEach(ll => n.lineas.add(ll));
            if (labToSalidas[l]) labToSalidas[l].forEach(s => n.salidas.add(s));
          });
        }
        if (invToSubs[n.id]) invToSubs[n.id].forEach(s => n.sublineas = (n.sublineas || new Set()).add(s));
      }
    }

    this.nodes = nodes;
    this.nodeIndex = nodeIndex;
    this.edges = edges;
    this._aux = { subToInvs, invToSubs, labToLineas, lineaToLabs, labToSalidas,
                  lineaToModos, invToModos, invToLabs };
  }

  _initSimulation() {
    // Force-directed orgánico con calibración por tipo.
    //
    // Lógica de los parámetros:
    //   - Las 4 líneas tienen carga moderada (no son anclas equidistantes:
    //     su posición emerge de las relaciones reales — sublíneas compartidas,
    //     proximidad cruzada entre sus sublíneas, labs en común). Líneas con
    //     muchas afinidades cruzadas se acercarán; líneas con pocas se
    //     alejarán naturalmente.
    //   - Las sublíneas tienen carga media; los investigadores carga ligera
    //     (orbitan alrededor de las sublíneas que cultivan sin dominar).
    //   - La arista jerárquica es muy fuerte y corta (sublínea pegada a su
    //     línea-madre); coautoría es media; afinidad-lab y coincidencia-modo
    //     casi nulas (ruido informativo).
    //   - La arista de proximidad semántica usa el valor de afinidad
    //     declarado en la planilla como modulación de la fuerza: pares con
    //     afinidad 0.9 se atraen casi tanto como una jerárquica; pares con
    //     0.5 se atraen apenas. Así la curaduría de la matriz se traduce
    //     en distancia visual real.
    //   - Collide previene overlap; alpha decay lento da tiempo a estabilizar.

    const { width, height } = this._currentSize();
    const cx = width / 2;
    const cy = height / 2;
    const radius = Math.min(width, height) * 0.3;

    // Posiciones iniciales en círculo — evita explosión inicial y da spread base.
    this.nodes.forEach((n, i) => {
      const angle = (i / this.nodes.length) * Math.PI * 2;
      n.x = cx + radius * Math.cos(angle);
      n.y = cy + radius * Math.sin(angle);
    });

    const linkDistance = (e) => ({
      jerarquica: 45,
      coautoria: 70,
      proximidad_semantica: 90,
      coinvestigacion: 130,
      sosten_lab: 130,
      afinidad_lab: 200,
      coincidencia_modo: 220,
    })[e.kind] ?? 100;

    // linkStrength por tipo de arista. Caso especial proximidad_semantica:
    // se calcula a partir del peso (afinidad declarada 0..1) para que la
    // curaduría de la matriz se exprese visualmente.
    //   weight 0.0 -> strength 0.20 (apenas une)
    //   weight 0.5 -> strength 0.45 (igual al valor plano anterior, baseline)
    //   weight 1.0 -> strength 0.70 (casi tan fuerte como una jerárquica)
    const linkStrength = (e) => {
      if (e.kind === 'proximidad_semantica') {
        const w = typeof e.weight === 'number' ? e.weight : 0.5;
        return 0.20 + 0.50 * Math.max(0, Math.min(1, w));
      }
      return ({
        jerarquica: 0.85,
        coautoria: 0.35,
        sosten_lab: 0.35,
        coinvestigacion: 0.15,
        afinidad_lab: 0.04,
        coincidencia_modo: 0.02,
      })[e.kind] ?? 0.1;
    };

    // Carga por tipo de nodo. Las líneas bajaron de -1000 a -350 para que
    // su posición no sea un polo equidistante artificial, sino el resultado
    // de las atracciones que reciben de sus sublíneas. Aún tienen suficiente
    // repulsión para no solaparse y para mantener separación legible.
    const chargeStrength = (n) => ({
      linea: -350,
      sublinea: -180,
      investigador: -60,
    })[n.kind] ?? -200;

    this.simulation = d3.forceSimulation(this.nodes)
      .force('link', d3.forceLink(this.edges).id(d => d.id)
        .distance(linkDistance)
        .strength(linkStrength))
      .force('charge', d3.forceManyBody()
        .strength(chargeStrength)
        .distanceMax(450))
      .force('center', d3.forceCenter(cx, cy).strength(0.04))
      .force('x', d3.forceX(cx).strength(0.015))
      .force('y', d3.forceY(cy).strength(0.015))
      .force('collide', d3.forceCollide()
        .radius(d => this._nodeRadius(d) + 6)
        .strength(0.9))
      .alphaDecay(0.018)
      .alphaMin(0.003);

    // Recalibrar al cambiar tamaño
    window.addEventListener('resize', () => this._recenter());
    requestAnimationFrame(() => this._recenter());
  }

  _currentSize() {
    const r = this.svgEl.getBoundingClientRect();
    return {
      width: r.width > 100 ? r.width : 1000,
      height: r.height > 100 ? r.height : 600,
    };
  }

  _recenter() {
    if (!this.simulation) return;
    const { width, height } = this._currentSize();
    const cx = width / 2;
    const cy = height / 2;
    this.simulation
      .force('center', d3.forceCenter(cx, cy).strength(0.04))
      .force('x', d3.forceX(cx).strength(0.015))
      .force('y', d3.forceY(cy).strength(0.015));
    this.simulation.alpha(0.5).restart();
  }

  _nodeRadius(n) {
    if (n.kind === 'linea') return 14;
    if (n.kind === 'investigador') return 8;
    // sublineas: tamaño según número de relaciones
    return 6 + Math.min(8, (n.investigadores ? n.investigadores.size : 0));
  }

  _render() {
    // Aristas
    const visibleEdges = this.edges.filter(e =>
      this.state.activeEdges.has(e.kind) &&
      this._isNodeVisible(e.source) && this._isNodeVisible(e.target)
    );
    const edgeSel = this.gEdges.selectAll('line')
      .data(visibleEdges, d => `${d.source.id}-${d.target.id}-${d.kind}`);
    edgeSel.exit().remove();
    edgeSel.enter().append('line')
      .attr('class', d => `edge ${d.kind}`);
    this.gEdges.selectAll('line').attr('class', d => {
      const att = this._isAttenuated(d.source) || this._isAttenuated(d.target);
      return `edge ${d.kind}${att ? ' attenuated' : ''}`;
    });

    // Nodos
    const visibleNodes = this.nodes.filter(n => this._isNodeVisible(n));
    const nodeSel = this.gNodes.selectAll('g.node')
      .data(visibleNodes, d => d.id);
    nodeSel.exit().remove();

    const nodeEnter = nodeSel.enter().append('g')
      .attr('class', d => `node ${d.kind}`)
      .call(d3.drag()
        .on('start', (event, d) => {
          if (!event.active) this.simulation.alphaTarget(0.3).restart();
          d.fx = d.x; d.fy = d.y;
        })
        .on('drag', (event, d) => { d.fx = event.x; d.fy = event.y; })
        .on('end', (event, d) => {
          if (!event.active) this.simulation.alphaTarget(0);
          // Al soltar, dejamos que las fuerzas reacomoden el grafo;
          // no fijamos posición para preservar el carácter orgánico.
          d.fx = null; d.fy = null;
        }))
      .on('click', (event, d) => { event.stopPropagation(); this._selectNode(d); })
      .on('mouseenter', (event, d) => this._showTooltip(event, d))
      .on('mousemove', (event) => this._moveTooltip(event))
      .on('mouseleave', () => this._hideTooltip());

    nodeEnter.each(function (d) {
      const g = d3.select(this);
      if (d.kind === 'investigador') {
        // Investigadores como cuadrados para diferenciarlos
        const r = 7;
        g.append('rect')
          .attr('x', -r).attr('y', -r)
          .attr('width', r * 2).attr('height', r * 2);
      } else {
        g.append('circle');
      }
    });

    this.gNodes.selectAll('g.node circle')
      .attr('r', d => this._nodeRadius(d));

    this.gNodes.selectAll('g.node')
      .attr('class', d => {
        let cls = `node ${d.kind}`;
        if (this._isAttenuated(d)) cls += ' attenuated';
        if (this._isHighlighted(d)) cls += ' highlighted';
        if (this.state.selectedNode && this.state.selectedNode.id === d.id) cls += ' selected';
        return cls;
      });

    // Envolventes (convex hulls) — se redibujan en cada tick
    this._renderEnvolventes();

    // Empty state
    if (visibleNodes.length === 0) {
      this.emptyState.classList.add('visible');
    } else {
      this.emptyState.classList.remove('visible');
    }

    // Re-energizar la simulación
    this.simulation.nodes(this.nodes);
    this.simulation.force('link').links(this.edges.filter(e =>
      this._isNodeVisible(e.source) && this._isNodeVisible(e.target)
    ));
    this.simulation.alpha(0.5).restart();

    this.simulation.on('tick', () => {
      this.gEdges.selectAll('line')
        .attr('x1', d => d.source.x).attr('y1', d => d.source.y)
        .attr('x2', d => d.target.x).attr('y2', d => d.target.y);
      this.gNodes.selectAll('g.node')
        .attr('transform', d => `translate(${d.x},${d.y})`);
      this._renderEnvolventes();
    });
  }

  _renderEnvolventes() {
    const groups = [];
    if (this.state.activeEnvolventes.has('area')) {
      for (const a of this.data.areas) {
        const members = this.nodes.filter(n => this._isNodeVisible(n) && n.areas.has(a.id));
        if (members.length >= 3) groups.push({ kind: 'area', cat: a.id, label: a.codigo || a.id, members });
      }
    }
    if (this.state.activeEnvolventes.has('modo')) {
      for (const m of this.data.modos) {
        const members = this.nodes.filter(n => this._isNodeVisible(n) && n.modos.has(m.id));
        if (members.length >= 3) groups.push({ kind: 'modo', cat: m.id, label: m.nombre, members });
      }
    }
    if (this.state.activeEnvolventes.has('salida')) {
      for (const s of this.data.salidas) {
        const members = this.nodes.filter(n => this._isNodeVisible(n) && n.salidas.has(s.id));
        if (members.length >= 3) groups.push({ kind: 'salida', cat: s.id, label: s.nombre, members });
      }
    }

    const sel = this.gEnvolventes.selectAll('g.envolvente-group')
      .data(groups, d => `${d.kind}-${d.cat}`);
    sel.exit().remove();
    const enter = sel.enter().append('g').attr('class', 'envolvente-group');
    enter.append('path');
    enter.append('text').attr('class', 'envolvente-label');

    this.gEnvolventes.selectAll('g.envolvente-group').each(function (d) {
      const points = d.members.map(n => [n.x || 0, n.y || 0]);
      const hull = d3.polygonHull(points);
      const path = d3.select(this).select('path');
      path.attr('class', d => `envolvente ${d.kind === 'area' ? `area-${d.cat}` : d.kind}`);
      if (hull) {
        // expandir el hull para que rodee con margen
        const cx = d3.mean(points, p => p[0]);
        const cy = d3.mean(points, p => p[1]);
        const expanded = hull.map(([x, y]) => {
          const dx = x - cx, dy = y - cy;
          const len = Math.sqrt(dx*dx + dy*dy) || 1;
          const margin = 28;
          return [x + (dx/len)*margin, y + (dy/len)*margin];
        });
        path.attr('d', 'M' + expanded.join('L') + 'Z');
        // Etiqueta en el centroide
        d3.select(this).select('text')
          .attr('x', cx).attr('y', cy)
          .text(d.label);
      } else {
        path.attr('d', null);
      }
    });
  }

  _isNodeVisible(n) {
    // Capa: investigadores sólo si showPerfiles
    if (n.kind === 'investigador' && !this.state.showPerfiles) return false;
    // Filtros: si un filtro tiene valores, el nodo debe coincidir en al menos uno
    const f = this.state.filters;
    if (f.linea.size && !this._setIntersect(n.lineas, f.linea)) return false;
    if (f.area.size && !this._setIntersect(n.areas, f.area)) return false;
    if (f.modo.size && !this._setIntersect(n.modos, f.modo)) return false;
    if (f.salida.size && !this._setIntersect(n.salidas, f.salida)) return false;
    if (f.lab.size && !this._setIntersect(n.labs, f.lab)) return false;
    if (f.investigador.size && !this._setIntersect(n.investigadores, f.investigador)) return false;
    return true;
  }

  _setIntersect(a, b) {
    if (!a) return false;
    for (const x of a) if (b.has(x)) return true;
    return false;
  }

  _isHighlighted(n) {
    if (!this.state.searchQuery) return false;
    return this._matchesSearch(n);
  }

  _isAttenuated(n) {
    return this.state.searchQuery && !this._matchesSearch(n);
  }

  _matchesSearch(n) {
    if (!this.state.searchQuery) return true;
    const q = this.state.searchQuery.toLowerCase();
    if (n.label && n.label.toLowerCase().includes(q)) return true;
    if (n.data.descripcion && n.data.descripcion.toLowerCase().includes(q)) return true;
    return false;
  }

  // ====== Operaciones públicas ======
  toggleEdge(kind) {
    if (this.state.activeEdges.has(kind)) this.state.activeEdges.delete(kind);
    else this.state.activeEdges.add(kind);
    this._render();
  }

  toggleEnvolvente(kind) {
    if (this.state.activeEnvolventes.has(kind)) this.state.activeEnvolventes.delete(kind);
    else this.state.activeEnvolventes.add(kind);
    this._render();
  }

  togglePerfiles() {
    this.state.showPerfiles = !this.state.showPerfiles;
    // Cuando los perfiles se encienden, asegurar que sus conexiones a sublíneas
    // (coautoría temática) sean visibles — un perfil sin aristas a su trabajo
    // queda flotando sin sentido.
    if (this.state.showPerfiles && this.config.allowedEdges.includes('coautoria')) {
      this.state.activeEdges.add('coautoria');
    }
    this._syncControlState();
    this._render();
  }

  setFilter(facet, ids) {
    this.state.filters[facet] = new Set(ids);
    this._render();
  }

  setSearch(q) {
    this.state.searchQuery = q.trim();
    this._render();
  }

  clearFilters() {
    Object.keys(this.state.filters).forEach(k => this.state.filters[k] = new Set());
    this.state.searchQuery = '';
    document.querySelectorAll('header.topbar input[type="search"]').forEach(el => el.value = '');
    document.querySelectorAll('aside.controls select[multiple]').forEach(el => {
      Array.from(el.options).forEach(o => o.selected = false);
    });
    this._render();
  }

  applyPreset(name) {
    if (name === 'cobertura_por_linea') {
      this.state.activeEnvolventes = new Set(['area']);
      this.state.activeEdges = new Set(['jerarquica']);
      this.state.showPerfiles = false;
    } else if (name === 'perfiles_por_area') {
      this.state.activeEnvolventes = new Set(['area']);
      this.state.activeEdges = new Set(['jerarquica','coinvestigacion','coautoria']);
      this.state.showPerfiles = true;
    }
    this._syncControlState();
    this._render();
  }

  // ====== Controles laterales ======
  _renderControls() {
    // No-op: cada HTML cablea su propia barra lateral.
    // Aquí sólo nos aseguramos que los toggles ya activos coincidan con el estado.
    this._syncControlState();
  }

  _syncControlState() {
    document.querySelectorAll('input[data-toggle-edge]').forEach(el => {
      el.checked = this.state.activeEdges.has(el.dataset.toggleEdge);
    });
    document.querySelectorAll('input[data-toggle-envolvente]').forEach(el => {
      el.checked = this.state.activeEnvolventes.has(el.dataset.toggleEnvolvente);
    });
    const perfilEl = document.querySelector('input[data-toggle-perfiles]');
    if (perfilEl) perfilEl.checked = this.state.showPerfiles;
  }

  _populateFilters() {
    const fillSelect = (sel, items, getId, getLabel) => {
      if (!sel) return;
      sel.innerHTML = '';
      for (const it of items) {
        const opt = document.createElement('option');
        opt.value = getId(it); opt.textContent = getLabel(it);
        sel.appendChild(opt);
      }
    };
    fillSelect(document.querySelector('select[data-filter="linea"]'),
               this.data.lineas, l => l.id, l => l.nombre);
    fillSelect(document.querySelector('select[data-filter="area"]'),
               this.data.areas, a => a.id, a => a.nombre);
    fillSelect(document.querySelector('select[data-filter="modo"]'),
               this.data.modos, m => m.id, m => m.nombre);
    fillSelect(document.querySelector('select[data-filter="salida"]'),
               this.data.salidas, s => s.id, s => s.nombre);
    fillSelect(document.querySelector('select[data-filter="lab"]'),
               this.data.laboratorios, l => l.id, l => l.nombre);
    fillSelect(document.querySelector('select[data-filter="investigador"]'),
               this.data.investigadores, i => i.id, i => i.nombre);
  }

  // ====== Tooltip y panel de detalle ======
  _showTooltip(event, n) {
    if (!this.tooltip) return;
    this.tooltip.style('display', 'block').text(n.label);
    this._moveTooltip(event);
  }
  _moveTooltip(event) {
    if (!this.tooltip) return;
    // El tooltip está dentro de main.canvas (position:relative), así que
    // las coordenadas deben ser relativas a ese contenedor, no a la página.
    const rect = this.svgEl.getBoundingClientRect();
    const x = event.clientX - rect.left + 14;
    const y = event.clientY - rect.top + 14;
    this.tooltip.style('left', x + 'px').style('top', y + 'px');
  }
  _hideTooltip() {
    if (this.tooltip) this.tooltip.style('display', 'none');
  }

  _selectNode(n) {
    this.state.selectedNode = n;
    if (!n) {
      this.detailPanel.classList.remove('visible');
    } else {
      this._renderDetailPanel(n);
      this.detailPanel.classList.add('visible');
    }
    this.gNodes.selectAll('g.node')
      .classed('selected', d => n && d.id === n.id);
  }

  _renderDetailPanel(n) {
    const d = n.data;
    const kindLabel = { linea: 'Línea troncal', sublinea: 'Sublínea', investigador: 'Investigador/a' }[n.kind];
    let html = `<button class="close" aria-label="Cerrar">×</button>
                <div class="kind">${kindLabel}</div>
                <h2>${d.nombre}</h2>`;
    if (d.descripcion) html += `<div class="descripcion">${d.descripcion}</div>`;

    html += '<dl>';
    if (n.kind === 'sublinea') {
      const linea = this.data.lineas.find(l => l.id === d.linea);
      const area = this.data.areas.find(a => a.id === d.area);
      if (linea) html += `<dt>Línea madre</dt><dd>${linea.nombre}</dd>`;
      if (area) html += `<dt>Área</dt><dd>${area.nombre}</dd>`;
      if (n.investigadores && n.investigadores.size) {
        const names = Array.from(n.investigadores)
          .map(id => this.data.investigadores.find(i => i.id === id)?.nombre).filter(Boolean);
        if (names.length) html += `<dt>Investigadores que la cultivan</dt><dd><ul>${names.map(x=>`<li>${x}</li>`).join('')}</ul></dd>`;
      }
    } else if (n.kind === 'linea') {
      const subs = this.data.sublineas.filter(s => s.linea === d.id);
      html += `<dt>Sublíneas</dt><dd><ul>${subs.map(s=>`<li>${s.nombre}</li>`).join('')}</ul></dd>`;
      const labs = Array.from(n.labs).map(id => this.data.laboratorios.find(l => l.id === id)?.nombre).filter(Boolean);
      if (labs.length) html += `<dt>Laboratorios que la sostienen</dt><dd>${labs.join(', ')}</dd>`;
    } else if (n.kind === 'investigador') {
      if (d.area_principal) {
        const area = this.data.areas.find(a => a.id === d.area_principal);
        html += `<dt>Área principal</dt><dd>${area ? area.nombre : d.area_principal}</dd>`;
      }
      if (n.sublineas && n.sublineas.size) {
        const names = Array.from(n.sublineas)
          .map(id => this.data.sublineas.find(s => s.id === id)?.nombre).filter(Boolean);
        html += `<dt>Sublíneas que cultiva</dt><dd><ul>${names.map(x=>`<li>${x}</li>`).join('')}</ul></dd>`;
      }
      if (d.perfil_url) html += `<dt>Perfil Casiopea</dt><dd><a href="${d.perfil_url}" target="_blank" rel="noopener">${d.perfil_url}</a></dd>`;
    }
    html += '</dl>';

    this.detailPanel.innerHTML = html;
    this.detailPanel.querySelector('.close').addEventListener('click', () => this._selectNode(null));
  }

  _hideLoading() {
    const ld = document.querySelector('.loading');
    if (ld) ld.style.display = 'none';
  }
}

window.MadMapGraph = MadMapGraph;
