'use strict';

/* ================= Utilidades ================= */
const MESES = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
const DIA_MS = 864e5;
const HORA_MS = 36e5;
const $ = s => document.querySelector(s);

const esc = s => String(s ?? '').replace(/[&<>"']/g, c => ({ '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;' }[c]));
const norm = s => String(s ?? '').normalize('NFD').replace(/[̀-ͯ]/g, '').toLowerCase().replace(/\s+/g, ' ').trim();
const pad2 = n => String(n).padStart(2, '0');
const monthKey = d => d.getFullYear() + '-' + pad2(d.getMonth() + 1);
const monthLabel = k => { const [y, m] = k.split('-'); return MESES[+m - 1] + ' ' + y; };
const monthShort = k => MESES[+k.split('-')[1] - 1];
function addMonths(k, n) {
  let [y, m] = k.split('-').map(Number);
  m += n;
  while (m < 1) { m += 12; y--; }
  while (m > 12) { m -= 12; y++; }
  return y + '-' + pad2(m);
}
const endOfMonth = k => { const [y, m] = k.split('-').map(Number); return new Date(y, m, 0, 23, 59, 59); };
const fmtDate = d => d ? pad2(d.getDate()) + '/' + pad2(d.getMonth() + 1) + '/' + d.getFullYear() : '';
const fmtDateTime = d => d ? fmtDate(d) + ' ' + pad2(d.getHours()) + ':' + pad2(d.getMinutes()) : '';
const fmtNum = (v, dec = 1) => v == null || !isFinite(v) ? '' : v.toLocaleString('es-CO', { minimumFractionDigits: dec, maximumFractionDigits: dec });
const isoDate = d => d.getFullYear() + '-' + pad2(d.getMonth() + 1) + '-' + pad2(d.getDate());
const mean = a => a.length ? a.reduce((s, x) => s + x, 0) / a.length : null;
function median(a) {
  if (!a.length) return null;
  const s = [...a].sort((x, y) => x - y), h = s.length >> 1;
  return s.length % 2 ? s[h] : (s[h - 1] + s[h]) / 2;
}
function num(v) {
  if (typeof v === 'number' && isFinite(v)) return v;
  if (typeof v === 'string') { const x = parseFloat(v.replace(',', '.')); return isFinite(x) ? x : null; }
  return null;
}

// Fecha desde celda de Excel: número serial, Date o texto "d/m/aaaa hh:mm[:ss] a. m./p. m."
function parseDate(v) {
  if (v instanceof Date) return isNaN(v) ? null : v;
  if (typeof v === 'number' && v > 1000 && v < 100000) {
    const d = new Date(Math.round((v - 25569) * DIA_MS));
    return new Date(d.getUTCFullYear(), d.getUTCMonth(), d.getUTCDate(), d.getUTCHours(), d.getUTCMinutes(), d.getUTCSeconds());
  }
  if (typeof v === 'string') {
    const s = v.replace(/ | /g, ' ').trim();
    const m = s.match(/^(\d{1,2})[\/\-.](\d{1,2})[\/\-.](\d{2,4})(?:[ T]+(\d{1,2}):(\d{2})(?::(\d{2}))?)?/);
    if (!m) return null;
    let [, d, mo, y, h = 0, mi = 0, se = 0] = m;
    y = +y < 100 ? 2000 + +y : +y;
    h = +h;
    const resto = s.slice(m[0].length).toLowerCase().replace(/[\s.]/g, '');
    if (resto.startsWith('pm') && h < 12) h += 12;
    if (resto.startsWith('am') && h === 12) h = 0;
    const f = new Date(y, +mo - 1, +d, h, +mi, +se);
    return isNaN(f) || f.getDate() !== +d ? null : f;
  }
  return null;
}

// "PH3" y "PH03" son el mismo cluster: letras + número a 2 dígitos + sufijo
function normCluster(c) {
  const s = String(c ?? '').toUpperCase().replace(/\s+/g, '');
  const m = s.match(/^([A-Z]+)0*(\d+)([A-Z]*)$/);
  return m ? m[1] + pad2(m[2]) + m[3] : s;
}
// "GIRASOL-H0712" y "GIRASOL H0712" son el mismo pozo
const wellKey = p => String(p ?? '').toUpperCase().replace(/[^A-Z0-9Ñ]/g, '');

function campoDesdeArchivo(nombre, primerPozo) {
  const m = nombre.match(/CAMPO\s+(.+?)\.xls/i);
  if (m) return m[1].trim().toUpperCase();
  if (primerPozo) return String(primerPozo).trim().split(/[\s\-]/)[0].toUpperCase();
  return nombre.replace(/\.xls.*$/i, '');
}

/* ================= Estado ================= */
const state = {
  archivos: new Map(),   // nombre -> { campo, filas: [...], alertas: [...] }
  tests: [], alerts: [], wells: new Map(), campoStats: {},
  result: null, abiertos: new Set(), tab: 'clusters', neqAuto: 1,
  sort: { pozos: null, clusters: null }
};

/* ================= Lectura de Excel ================= */
const COLS = {
  cluster: n => n === 'cluster',
  pozo: n => n === 'pozo',
  equipo: n => n.startsWith('equipo'),
  alin: n => n.includes('alineacion'),
  ini: n => n.includes('inicio'),
  fin: n => n.includes('finalizacion') || n === 'fecha fin',
  dur: n => n.startsWith('duracion'),
  liq: n => n.startsWith('total liquido'),
  bsw: n => n.startsWith('bs&w') || n.startsWith('bsw'),
  crudo: n => n.startsWith('crudo'),
  agua: n => n.startsWith('agua'),
  gas: n => n.startsWith('gas'),
  valida: n => n.includes('valida')
};

function leerLibro(nombre, buf) {
  const wb = XLSX.read(buf, { type: 'array' });
  const hojas = [...wb.SheetNames].sort((a, b) => (norm(b) === 'pruebas') - (norm(a) === 'pruebas'));
  for (const hoja of hojas) {
    const rows = XLSX.utils.sheet_to_json(wb.Sheets[hoja], { header: 1, raw: true, defval: null });
    for (let i = 0; i < Math.min(rows.length, 40); i++) {
      const h = (rows[i] || []).map(norm);
      if (!h.includes('cluster') || !h.includes('pozo')) continue;
      const ix = {};
      for (const [k, f] of Object.entries(COLS)) ix[k] = h.findIndex(f);
      const filas = [];
      for (let r = i + 1; r < rows.length; r++) {
        const row = rows[r] || [];
        if (row[ix.pozo] == null || String(row[ix.pozo]).trim() === '') continue;
        const g = k => ix[k] >= 0 ? row[ix[k]] : null;
        filas.push({
          fila: r + 1, archivo: nombre,
          clusterRaw: g('cluster'), pozo: String(g('pozo')).trim().replace(/\s+/g, ' '),
          equipo: g('equipo') ? String(g('equipo')).trim().toUpperCase() : '',
          alinRaw: g('alin'), iniRaw: g('ini'), finRaw: g('fin'),
          dur: num(g('dur')), liq: num(g('liq')), bsw: num(g('bsw')), crudo: num(g('crudo')),
          agua: num(g('agua')), gas: num(g('gas')), valida: g('valida')
        });
      }
      return { campo: campoDesdeArchivo(nombre, filas[0] && filas[0].pozo), filas };
    }
  }
  throw new Error('No se encontró una hoja con columnas "CLUSTER" y "Pozo".');
}

/* ================= Procesamiento ================= */
function procesar() {
  const tests = [], alerts = [];
  const hoy = new Date(), limiteFuturo = new Date(hoy.getTime() + 2 * DIA_MS);
  const mapCluster = new Map(), nombresPozo = new Map();
  const alerta = (t, tipo, detalle) => alerts.push({ campo: t.campo, archivo: t.archivo, fila: t.fila, pozo: t.pozo, tipo, detalle });

  // 1) Fechas y normalización
  const crudos = [];
  for (const [, a] of state.archivos) {
    for (const f of a.filas) {
      const t = { ...f, campo: a.campo };
      t.fAlin = parseDate(f.alinRaw); t.fIni = parseDate(f.iniRaw); t.fFin = parseDate(f.finRaw);
      t.fecha = t.fIni || t.fAlin || t.fFin;
      if (typeof f.iniRaw === 'string' && t.fIni) alerta(t, 'Fecha como texto', `Fecha de inicio escrita como texto ("${f.iniRaw}"); se interpretó como ${fmtDateTime(t.fIni)}.`);
      if (!t.fecha) { alerta(t, 'Sin fecha', 'La prueba no tiene fecha de inicio válida; no se tuvo en cuenta.'); continue; }
      const raw = f.clusterRaw == null ? '' : String(f.clusterRaw).trim();
      t.cluster = raw ? normCluster(raw) : 'SIN CLUSTER';
      if (!raw) alerta(t, 'Sin cluster', 'La prueba no tiene cluster; se agrupó en "SIN CLUSTER".');
      else if (t.cluster !== raw.toUpperCase()) mapCluster.set(a.campo + '|' + raw + '→' + t.cluster, { campo: a.campo, archivo: f.archivo, raw, nuevo: t.cluster });
      t.key = a.campo + '|' + wellKey(f.pozo);
      if (!nombresPozo.has(t.key)) nombresPozo.set(t.key, new Set());
      nombresPozo.get(t.key).add(f.pozo);
      const v = norm(f.valida);
      if (['no', 'n', 'invalida', 'no valida', 'false'].includes(v)) { alerta(t, 'Prueba no válida', 'Marcada como no válida; no se tuvo en cuenta.'); continue; }
      crudos.push(t);
    }
  }

  // 2) Fechas fuera de rango (por campo): más de 2 años antes del dato más reciente o en el futuro
  const ultimaPorCampo = {};
  for (const t of crudos) if (t.fecha <= limiteFuturo && (!ultimaPorCampo[t.campo] || t.fecha > ultimaPorCampo[t.campo])) ultimaPorCampo[t.campo] = t.fecha;
  for (const t of crudos) {
    const ult = ultimaPorCampo[t.campo];
    if (t.fecha > limiteFuturo) { alerta(t, 'Fecha futura', `Fecha ${fmtDateTime(t.fecha)} posterior a hoy; no se tuvo en cuenta.`); continue; }
    if (ult && t.fecha < new Date(ult.getTime() - 730 * DIA_MS)) { alerta(t, 'Fecha fuera de rango', `Fecha ${fmtDateTime(t.fecha)} muy anterior al resto del historial (¿error de digitación?); no se tuvo en cuenta.`); continue; }
    if (t.fFin && t.fIni && t.fFin < t.fIni) alerta(t, 'Fechas invertidas', 'La fecha de finalización es anterior a la de inicio.');
    t.mes = monthKey(t.fecha);
    tests.push(t);
  }

  for (const m of mapCluster.values()) alerts.push({ campo: m.campo, archivo: m.archivo, fila: '', pozo: '', tipo: 'Cluster unificado', detalle: `"${m.raw}" se tomó como "${m.nuevo}".` });
  for (const [k, s] of nombresPozo) if (s.size > 1) alerts.push({ campo: k.split('|')[0], archivo: '', fila: '', pozo: [...s][0], tipo: 'Pozo unificado', detalle: 'Mismo pozo escrito de varias formas: ' + [...s].map(x => `"${x}"`).join(', ') + '.' });

  // 3) Tiempos típicos por campo: preparación (alineación→inicio) y espera entre pruebas del mismo equipo
  const campoStats = {};
  const porCampo = groupBy(tests, t => t.campo);
  for (const [campo, ts] of porCampo) {
    const prep = ts.filter(t => t.fAlin && t.fIni).map(t => (t.fIni - t.fAlin) / HORA_MS).filter(h => h >= 0 && h <= 12);
    const gaps = [];
    for (const [, te] of groupBy(ts.filter(t => t.equipo && t.fAlin && t.fFin), t => t.equipo)) {
      te.sort((a, b) => a.fAlin - b.fAlin);
      for (let i = 1; i < te.length; i++) {
        const g = (te[i].fAlin - te[i - 1].fFin) / HORA_MS;
        if (g >= 0 && g <= 12) gaps.push(g);
      }
    }
    campoStats[campo] = { prep: median(prep) ?? 1, gap: median(gaps) ?? 0.5, ultima: ultimaPorCampo[campo] };
  }

  // 4) Ciclo de cada prueba = (fin − alineación) + espera entre pruebas
  for (const t of tests) {
    const cs = campoStats[t.campo];
    const real = t.fAlin && t.fFin ? (t.fFin - t.fAlin) / HORA_MS : null;
    if (real != null && real > 0 && real <= 72) t.ciclo = real + cs.gap;
    else if (t.dur != null && t.dur > 0 && t.dur <= 72) t.ciclo = t.dur + cs.prep + cs.gap;
    else t.ciclo = null;
  }

  // 5) Pozos
  const wells = new Map();
  for (const [key, ts] of groupBy(tests, t => t.key)) {
    ts.sort((a, b) => a.fecha - b.fecha);
    const ult = ts[ts.length - 1];
    const clusters = new Set(ts.map(t => t.cluster));
    if (clusters.size > 1) alerts.push({ campo: ult.campo, archivo: '', fila: '', pozo: ult.pozo, tipo: 'Pozo en varios clusters', detalle: `Aparece en ${[...clusters].join(', ')}; se usa el más reciente (${ult.cluster}).` });
    const ciclos = ts.map(t => t.ciclo).filter(c => c != null);
    wells.set(key, {
      key, campo: ult.campo, pozo: ult.pozo, cluster: ult.cluster, tests: ts,
      ciclo: mean(ciclos), nCiclos: ciclos.length
    });
  }

  state.tests = tests; state.alerts = alerts; state.wells = wells; state.campoStats = campoStats;
}

function groupBy(arr, fn) {
  const m = new Map();
  for (const x of arr) { const k = fn(x); if (!m.has(k)) m.set(k, []); m.get(k).push(x); }
  return m;
}

/* ================= Análisis ================= */
function params() {
  const n = id => { const v = parseFloat($(id).value); return isFinite(v) ? v : null; };
  return {
    campo: $('#f-campo').value, mes: $('#f-mes').value, ref: $('#f-ref').value,
    inact: Math.max(0, n('#f-inact') || 0), hpd: Math.min(24, Math.max(1, n('#f-hpd') || 24)),
    neq: Math.max(1, Math.round(n('#f-neq') || state.neqAuto)), mov: Math.max(0, n('#f-mov') || 0),
    ini: $('#f-ini').value ? new Date($('#f-ini').value + 'T00:00:00') : new Date(), prio: $('#f-prio').value
  };
}

function analizar(p) {
  const enCampo = x => p.campo === '__all' || x.campo === p.campo;
  const tests = state.tests.filter(enCampo);
  const cur = p.mes;
  const refIni = p.ref === 'all' ? '0000-00' : addMonths(cur, -Number(p.ref));
  const ultima = tests.reduce((m, t) => t.fecha > m ? t.fecha : m, new Date(0));
  const corte = new Date(Math.min(endOfMonth(cur).getTime(), ultima.getTime()));

  // Promedios de ciclo por cluster y campo (fallback para pozos sin datos)
  const wellsCampo = [...state.wells.values()].filter(enCampo);
  const cicloCluster = new Map(), cicloCampo = new Map();
  for (const [k, ws] of groupBy(wellsCampo.filter(w => w.ciclo != null), w => w.campo + '|' + w.cluster)) cicloCluster.set(k, mean(ws.map(w => w.ciclo)));
  for (const [k, ws] of groupBy(wellsCampo.filter(w => w.ciclo != null), w => w.campo)) cicloCampo.set(k, mean(ws.map(w => w.ciclo)));

  const pozos = [];
  for (const w of wellsCampo) {
    const tRef = w.tests.filter(t => t.mes >= refIni && t.mes < cur);
    const tCur = w.tests.filter(t => t.mes === cur);
    if (!tRef.length && !tCur.length) continue;
    const previas = w.tests.filter(t => t.fecha <= corte);
    const ult = previas[previas.length - 1] || tCur[0];
    const ultRef = tRef[tRef.length - 1];
    const dias = Math.max(0, Math.floor((corte - (tCur.length ? tCur[tCur.length - 1].fecha : ult.fecha)) / DIA_MS));
    let estado = 'medido';
    if (!tCur.length) estado = p.inact > 0 && dias > p.inact ? 'inactivo' : 'pendiente';
    let ciclo = w.ciclo, fuente = 'pozo';
    if (ciclo == null) { ciclo = cicloCluster.get(w.campo + '|' + w.cluster); fuente = 'cluster'; }
    if (ciclo == null) { ciclo = cicloCampo.get(w.campo); fuente = 'campo'; }
    if (ciclo == null) { ciclo = 8; fuente = 'defecto'; }
    const ref = ultRef || ult;
    pozos.push({
      w, campo: w.campo, cluster: w.cluster, pozo: w.pozo, estado, dias, ciclo, fuente,
      ultima: ult.fecha, nPruebas: w.tests.filter(t => t.mes <= cur).length,
      meses: new Set(w.tests.map(t => t.mes)), equipo: ult.equipo,
      crudo: ref.crudo, liq: ref.liq, bsw: ref.bsw
    });
  }

  const clusters = [];
  for (const [k, ps] of groupBy(pozos, x => x.campo + '|' + x.cluster)) {
    const pend = ps.filter(x => x.estado === 'pendiente');
    const horas = pend.reduce((s, x) => s + x.ciclo, 0);
    clusters.push({
      key: k, campo: ps[0].campo, cluster: ps[0].cluster, pozos: ps,
      total: ps.length, medidos: ps.filter(x => x.estado === 'medido').length,
      pendientes: pend.length, inactivos: ps.filter(x => x.estado === 'inactivo').length,
      horas, dias: horas / p.hpd,
      crudo: pend.reduce((s, x) => s + (x.crudo || 0), 0),
      maxDias: pend.reduce((m, x) => Math.max(m, x.dias), 0)
    });
  }

  // Plan: reparto de clusters pendientes entre equipos (al menos cargado primero)
  const orden = {
    largo: (a, b) => b.horas - a.horas,
    crudo: (a, b) => b.crudo - a.crudo || b.horas - a.horas,
    antiguo: (a, b) => b.maxDias - a.maxDias || b.horas - a.horas,
    nombre: (a, b) => (a.campo + a.cluster).localeCompare(b.campo + b.cluster, 'es', { numeric: true })
  }[p.prio];
  const conPend = clusters.filter(c => c.pendientes > 0).sort(orden);
  const equipos = Array.from({ length: p.neq }, (_, i) => ({ n: i + 1, horas: 0, items: [] }));
  const aFecha = h => new Date(p.ini.getTime() + (h / p.hpd) * DIA_MS);
  for (const c of conPend) {
    const e = equipos.reduce((a, b) => b.horas < a.horas ? b : a);
    const h0 = e.horas + (e.items.length ? p.mov : 0);
    e.horas = h0 + c.horas;
    e.items.push({ c, ini: aFecha(h0), fin: aFecha(e.horas) });
  }
  const horasTot = conPend.reduce((s, c) => s + c.horas, 0) + Math.max(0, conPend.length - 1) * p.mov;
  const diasN = Math.max(0, ...equipos.map(e => e.horas)) / p.hpd;

  // Avance mensual
  const meses = [...new Set(tests.map(t => t.mes))].sort();
  const avance = meses.map(m => ({ mes: m, pozos: new Set(tests.filter(t => t.mes === m).map(t => t.key)).size, pruebas: tests.filter(t => t.mes === m).length }));

  return {
    p, cur, refIni, corte, pozos, clusters, equipos, conPend, avance,
    dias1: horasTot / p.hpd, diasN, finEst: aFecha(Math.max(0, ...equipos.map(e => e.horas))),
    finMes: endOfMonth(cur)
  };
}

/* ================= Render ================= */
function render() {
  if (!state.tests.length) { $('#panel').hidden = true; return; }
  $('#panel').hidden = false;
  const r = state.result = analizar(params());
  guardarConfig();
  renderKpis(r);
  renderClusters(r);
  renderPozos(r);
  renderPlan(r);
  renderAvance(r);
  renderAlertas();
}

function renderKpis(r) {
  const pend = r.pozos.filter(x => x.estado === 'pendiente').length;
  const med = r.pozos.filter(x => x.estado === 'medido').length;
  const inac = r.pozos.filter(x => x.estado === 'inactivo').length;
  const total = pend + med;
  const diasMes = Math.max(0, (r.finMes - r.p.ini) / DIA_MS);
  const excede = r.finEst > r.finMes;
  const refTxt = r.p.ref === 'all' ? 'todo el historial' : `últimos ${r.p.ref} mes(es)`;
  $('#kpis').innerHTML = [
    kpi(pend, 'Pozos pendientes en ' + monthLabel(r.cur), `de ${total} pozos con historial (${refTxt})`, pend ? 'bad' : 'ok'),
    kpi(r.conPend.length, 'Clusters con pendientes', `de ${r.clusters.length} clusters`, r.conPend.length ? 'warn' : 'ok'),
    kpi(fmtNum(total ? med / total * 100 : 0, 0) + '%', 'Avance del mes', `${med} pozos medidos`, ''),
    kpi(fmtNum(r.dias1), 'Días con 1 equipo', `${fmtNum(r.dias1 * r.p.hpd, 0)} h a ${r.p.hpd} h/día`, ''),
    kpi(fmtNum(r.diasN), `Días con ${r.p.neq} equipo(s)`, `Termina aprox. ${fmtDate(r.finEst)}`, excede ? 'bad' : 'ok'),
    kpi(fmtNum(diasMes, 0), 'Días restantes del mes', excede ? `Faltan ~${fmtNum((r.finEst - r.finMes) / DIA_MS)} días de capacidad` : 'Alcanza dentro del mes', excede ? 'bad' : 'ok'),
    inac ? kpi(inac, 'Ocultos por inactividad', `sin medir hace más de ${r.p.inact} días`, 'warn') : ''
  ].join('');
}
const kpi = (v, l, s, cls) => `<div class="kpi ${cls}"><div class="l">${esc(l)}</div><div class="v">${esc(v)}</div><div class="s">${esc(s)}</div></div>`;

function sortRows(rows, cfg, getters) {
  if (!cfg) return rows;
  const g = getters[cfg.col];
  return [...rows].sort((a, b) => {
    const x = g(a), y = g(b);
    const c = typeof x === 'string' ? x.localeCompare(y, 'es', { numeric: true }) : (x ?? -Infinity) - (y ?? -Infinity);
    return cfg.dir * c;
  });
}
function th(tabla, col, label, cls = '') {
  const s = state.sort[tabla];
  const flecha = s && s.col === col ? (s.dir > 0 ? ' ▲' : ' ▼') : '';
  return `<th class="${cls}" data-sort="${tabla}:${col}">${esc(label)}${flecha}</th>`;
}

const GET_CL = {
  campo: c => c.campo, cluster: c => c.cluster, total: c => c.total, medidos: c => c.medidos, pendientes: c => c.pendientes,
  avance: c => c.medidos / c.total, horas: c => c.horas, dias: c => c.dias, maxDias: c => c.maxDias, crudo: c => c.crudo
};
function renderClusters(r) {
  const multi = r.p.campo === '__all';
  let rows = r.clusters.filter(c => !$('#soloPend').checked || c.pendientes > 0);
  rows = state.sort.clusters ? sortRows(rows, state.sort.clusters, GET_CL) : rows.sort((a, b) => b.dias - a.dias || a.cluster.localeCompare(b.cluster, 'es', { numeric: true }));
  if (!rows.length) { $('#tblClusters').innerHTML = '<p class="vacio">No hay clusters con pozos pendientes. 🎉</p>'; return; }
  const cols = (multi ? 10 : 9);
  let h = '<table><thead><tr>' + (multi ? th('clusters', 'campo', 'Campo') : '') +
    th('clusters', 'cluster', 'Cluster') + th('clusters', 'total', 'Pozos', 'num') + th('clusters', 'medidos', 'Medidos', 'num') +
    th('clusters', 'pendientes', 'Pendientes', 'num') + th('clusters', 'avance', 'Avance') + th('clusters', 'horas', 'Horas est.', 'num') +
    th('clusters', 'dias', 'Días est.', 'num') + th('clusters', 'maxDias', 'Máx. días sin medir', 'num') + th('clusters', 'crudo', 'Crudo pend. (BOPD)', 'num') +
    '</tr></thead><tbody>';
  for (const c of rows) {
    const abierto = state.abiertos.has(c.key);
    const pct = c.total ? c.medidos / c.total * 100 : 0;
    h += `<tr class="cl" data-cl="${esc(c.key)}">` + (multi ? `<td>${esc(c.campo)}</td>` : '') +
      `<td><span class="flecha">${abierto ? '▾' : '▸'}</span><b>${esc(c.cluster)}</b></td><td class="num">${c.total}</td><td class="num">${c.medidos}</td>` +
      `<td class="num">${c.pendientes ? `<b>${c.pendientes}</b>` : 0}</td><td><span class="prog"><i style="width:${pct}%"></i></span>${fmtNum(pct, 0)}%</td>` +
      `<td class="num">${fmtNum(c.horas)}</td><td class="num"><b>${fmtNum(c.dias)}</b></td><td class="num">${c.pendientes ? c.maxDias : ''}</td><td class="num">${fmtNum(c.crudo)}</td></tr>`;
    if (abierto) {
      const ps = [...c.pozos].sort((a, b) => ({ pendiente: 0, inactivo: 1, medido: 2 }[a.estado] - { pendiente: 0, inactivo: 1, medido: 2 }[b.estado]) || a.pozo.localeCompare(b.pozo, 'es', { numeric: true }));
      h += `<tr class="det"><td colspan="${cols}"><div class="pozos-lista">` + ps.map(x =>
        `<div><span><span class="tag ${x.estado}">${x.estado}</span> ${esc(x.pozo)}</span><span>${x.estado === 'medido' ? 'medido ' + fmtDate(x.ultima) : fmtNum(x.ciclo) + ' h' + (x.fuente !== 'pozo' ? ` <span class="estim">(prom. ${x.fuente})</span>` : '')}</span></div>`
      ).join('') + '</div></td></tr>';
    }
  }
  $('#tblClusters').innerHTML = h + '</tbody></table>';
}

const GET_PZ = {
  campo: x => x.campo, cluster: x => x.cluster, pozo: x => x.pozo, estado: x => x.estado, ultima: x => x.ultima.getTime(), dias: x => x.dias,
  nPruebas: x => x.nPruebas, ciclo: x => x.ciclo, crudo: x => x.crudo, liq: x => x.liq, equipo: x => x.equipo
};
function renderPozos(r) {
  const multi = r.p.campo === '__all';
  const q = norm($('#buscar').value), est = $('#f-estado').value;
  let rows = r.pozos.filter(x => (est === 'todos' || x.estado === est) && (!q || norm(x.pozo + ' ' + x.cluster).includes(q)));
  rows = state.sort.pozos ? sortRows(rows, state.sort.pozos, GET_PZ) : rows.sort((a, b) => b.dias - a.dias || (b.crudo || 0) - (a.crudo || 0));
  if (!rows.length) { $('#tblPozos').innerHTML = '<p class="vacio">No hay pozos con este filtro.</p>'; return; }
  // Últimos meses a mostrar como "chips" (máx. 8)
  const mesesVis = [];
  for (let k = r.cur, i = 0; i < 8 && k >= (r.avance[0] ? r.avance[0].mes : r.cur); i++, k = addMonths(k, -1)) mesesVis.unshift(k);
  let h = '<table><thead><tr>' + (multi ? th('pozos', 'campo', 'Campo') : '') + th('pozos', 'cluster', 'Cluster') + th('pozos', 'pozo', 'Pozo') +
    th('pozos', 'estado', 'Estado') + th('pozos', 'ultima', 'Última medición') + th('pozos', 'dias', 'Días sin medir', 'num') +
    '<th>Meses medidos</th>' + th('pozos', 'nPruebas', 'N° pruebas', 'num') + th('pozos', 'ciclo', 'Tiempo est. (h)', 'num') +
    th('pozos', 'crudo', 'Crudo (BOPD)', 'num') + th('pozos', 'liq', 'Líquido (BFPD)', 'num') + th('pozos', 'equipo', 'Último equipo') + '</tr></thead><tbody>';
  for (const x of rows) {
    h += '<tr>' + (multi ? `<td>${esc(x.campo)}</td>` : '') + `<td>${esc(x.cluster)}</td><td><b>${esc(x.pozo)}</b></td><td><span class="tag ${x.estado}">${x.estado}</span></td>` +
      `<td>${fmtDate(x.ultima)}</td><td class="num">${x.estado === 'medido' ? '' : x.dias}</td>` +
      `<td><span class="meses">${mesesVis.map(m => `<b class="${x.meses.has(m) ? 'si' : ''} ${m === r.cur ? 'cur' : ''}" title="${monthLabel(m)}">${monthShort(m)}</b>`).join('')}</span></td>` +
      `<td class="num">${x.nPruebas}</td><td class="num">${fmtNum(x.ciclo)}${x.fuente !== 'pozo' ? ' <span class="estim">*</span>' : ''}</td>` +
      `<td class="num">${fmtNum(x.crudo)}</td><td class="num">${fmtNum(x.liq)}</td><td>${esc(x.equipo)}</td></tr>`;
  }
  $('#tblPozos').innerHTML = h + '</tbody></table><p class="hint">* Tiempo estimado con el promedio del cluster o del campo (el pozo no tiene fechas suficientes).</p>';
}

function renderPlan(r) {
  if (!r.conPend.length) { $('#plan').innerHTML = '<p class="vacio">No hay clusters pendientes para planificar.</p>'; return; }
  const excede = r.finEst > r.finMes;
  let h = `<div class="aviso ${excede ? '' : 'ok'}">Con <b>${r.p.neq}</b> equipo(s) a <b>${r.p.hpd} h/día</b>, iniciando el <b>${fmtDate(r.p.ini)}</b>, se terminaría aprox. el <b>${fmtDate(r.finEst)}</b>` +
    (excede ? ` — después del cierre de ${monthLabel(r.cur)} (${fmtDate(r.finMes)}).` : ` — dentro de ${monthLabel(r.cur)}.`) + '</div>';
  if (r.p.campo === '__all') h += '<p class="hint">Estás viendo todos los campos: el plan reparte los equipos entre campos distintos. Para un plan realista filtra un solo campo.</p>';
  h += '<div class="equipos">';
  for (const e of r.equipos) {
    h += `<div class="equipo"><h3><span>Equipo ${e.n}</span><span>${fmtNum(e.horas / r.p.hpd)} días</span></h3>`;
    if (!e.items.length) h += '<p class="hint">Sin asignación.</p>';
    else h += '<ol>' + e.items.map(it => `<li><b>${esc((r.p.campo === '__all' ? it.c.campo + ' · ' : '') + it.c.cluster)}</b> — ${it.c.pendientes} pozo(s), ${fmtNum(it.c.dias)} días<small>${fmtDateTime(it.ini)} → ${fmtDateTime(it.fin)}</small></li>`).join('') + '</ol>';
    h += '</div>';
  }
  $('#plan').innerHTML = h + '</div><p class="hint">El reparto asigna cada cluster al equipo que queda libre primero. No considera distancias entre clusters; ajusta las horas de movilización si aplica.</p>';
}

function renderAvance(r) {
  const max = Math.max(1, ...r.avance.map(a => a.pozos));
  let h = '<h2>Pozos distintos medidos por mes</h2><div class="barras">' + r.avance.map(a =>
    `<div class="col ${a.mes === r.cur ? 'cur' : ''}" title="${a.pruebas} pruebas"><em>${a.pozos}</em><i style="height:${a.pozos / max * 170}px"></i><span>${monthLabel(a.mes)}</span></div>`).join('') + '</div>';
  h += '<div class="tabla"><table><thead><tr><th>Mes</th><th class="num">Pozos medidos</th><th class="num">Pruebas</th></tr></thead><tbody>' +
    r.avance.map(a => `<tr><td>${monthLabel(a.mes)}${a.mes === r.cur ? ' (mes revisado)' : ''}</td><td class="num">${a.pozos}</td><td class="num">${a.pruebas}</td></tr>`).join('') + '</tbody></table></div>';
  $('#avance').innerHTML = h;
}

function renderAlertas() {
  const p = $('#f-campo').value;
  const al = state.alerts.filter(a => p === '__all' || a.campo === p);
  $('#nAlertas').textContent = al.length || '';
  if (!al.length) { $('#tblAlertas').innerHTML = '<p class="vacio">Sin alertas de calidad de datos.</p>'; return; }
  $('#tblAlertas').innerHTML = '<table><thead><tr><th>Campo</th><th>Tipo</th><th>Pozo</th><th>Fila</th><th>Detalle</th></tr></thead><tbody>' +
    al.map(a => `<tr><td>${esc(a.campo)}</td><td>${esc(a.tipo)}</td><td>${esc(a.pozo)}</td><td class="num">${esc(a.fila)}</td><td style="white-space:normal">${esc(a.detalle)}</td></tr>`).join('') + '</tbody></table>';
}

/* ================= Controles ================= */
function poblarControles() {
  const campos = [...new Set(state.tests.map(t => t.campo))].sort();
  const selC = $('#f-campo'), prevC = selC.value || cfg.campo;
  selC.innerHTML = (campos.length > 1 ? '<option value="__all">Todos los campos</option>' : '') + campos.map(c => `<option>${esc(c)}</option>`).join('');
  selC.value = campos.includes(prevC) || (prevC === '__all' && campos.length > 1) ? prevC : campos[0];
  poblarMeses();
}
function poblarMeses() {
  const c = $('#f-campo').value;
  const meses = [...new Set(state.tests.filter(t => c === '__all' || t.campo === c).map(t => t.mes))].sort().reverse();
  const sel = $('#f-mes'), prev = sel.value;
  sel.innerHTML = meses.map(m => `<option value="${m}">${monthLabel(m)}</option>`).join('');
  sel.value = meses.includes(prev) ? prev : meses[0];
  actualizarEquiposAuto();
}
// Equipos por defecto = equipos distintos usados en el mes revisado
function actualizarEquiposAuto() {
  const c = $('#f-campo').value, m = $('#f-mes').value;
  const eq = new Set(state.tests.filter(t => (c === '__all' || t.campo === c) && t.mes === m && t.equipo).map(t => t.equipo));
  state.neqAuto = Math.max(1, eq.size);
  $('#f-neq').placeholder = `Auto: ${state.neqAuto}`;
}

let cfg = {};
try { cfg = JSON.parse(localStorage.getItem('mpfm-config') || '{}'); } catch (e) { cfg = {}; }
function guardarConfig() {
  try {
    localStorage.setItem('mpfm-config', JSON.stringify({
      campo: $('#f-campo').value, ref: $('#f-ref').value, inact: $('#f-inact').value, hpd: $('#f-hpd').value,
      mov: $('#f-mov').value, prio: $('#f-prio').value
    }));
  } catch (e) { /* almacenamiento no disponible */ }
}
function aplicarConfig() {
  if (cfg.ref) $('#f-ref').value = cfg.ref;
  if (cfg.inact) $('#f-inact').value = cfg.inact;
  if (cfg.hpd) $('#f-hpd').value = cfg.hpd;
  if (cfg.mov) $('#f-mov').value = cfg.mov;
  if (cfg.prio) $('#f-prio').value = cfg.prio;
  $('#f-ini').value = isoDate(new Date());
}

async function cargarArchivos(files) {
  const errores = [];
  for (const f of files) {
    try {
      const a = leerLibro(f.name, await f.arrayBuffer());
      if (!a.filas.length) throw new Error('No tiene pruebas.');
      state.archivos.set(f.name, a);
    } catch (e) { errores.push(`${f.name}: ${e.message}`); }
  }
  if (errores.length) alert('No se pudieron leer algunos archivos:\n\n' + errores.join('\n'));
  refrescarArchivos();
}
function refrescarArchivos() {
  $('#archivos').innerHTML = [...state.archivos].map(([n, a]) =>
    `<span class="archivo" title="${esc(n)}"><b>${esc(a.campo)}</b> · ${a.filas.length} pruebas<button type="button" data-quitar="${esc(n)}" aria-label="Quitar">×</button></span>`).join('');
  procesar();
  if (state.tests.length) poblarControles();
  render();
}

function exportar() {
  const r = state.result; if (!r) return;
  const wb = XLSX.utils.book_new();
  const hoja = (nombre, filas) => XLSX.utils.book_append_sheet(wb, XLSX.utils.json_to_sheet(filas.length ? filas : [{ '': 'Sin datos' }]), nombre);
  hoja('Por cluster', r.clusters.filter(c => c.pendientes > 0).sort((a, b) => b.dias - a.dias).map(c => ({
    Campo: c.campo, Cluster: c.cluster, 'Pozos': c.total, Medidos: c.medidos, Pendientes: c.pendientes,
    'Avance %': +(c.medidos / c.total * 100).toFixed(1), 'Horas estimadas': +c.horas.toFixed(1), 'Días estimados': +c.dias.toFixed(2),
    'Máx. días sin medir': c.maxDias, 'Crudo pendiente (BOPD)': +c.crudo.toFixed(1),
    'Pozos pendientes': c.pozos.filter(x => x.estado === 'pendiente').map(x => x.pozo).join(', ')
  })));
  hoja('Pozos pendientes', r.pozos.filter(x => x.estado === 'pendiente').sort((a, b) => b.dias - a.dias).map(filaPozo));
  hoja('Todos los pozos', r.pozos.map(filaPozo));
  const plan = [];
  for (const e of r.equipos) for (const it of e.items) plan.push({ Equipo: e.n, Campo: it.c.campo, Cluster: it.c.cluster, Pozos: it.c.pendientes, 'Días': +it.c.dias.toFixed(2), Inicio: fmtDateTime(it.ini), Fin: fmtDateTime(it.fin) });
  hoja('Plan por equipo', plan);
  hoja('Alertas', state.alerts.filter(a => r.p.campo === '__all' || a.campo === r.p.campo).map(a => ({ Campo: a.campo, Tipo: a.tipo, Pozo: a.pozo, Archivo: a.archivo, Fila: a.fila, Detalle: a.detalle })));
  const campo = r.p.campo === '__all' ? 'TODOS' : r.p.campo;
  XLSX.writeFile(wb, `Pendientes_${campo}_${r.cur}.xlsx`);
}
const filaPozo = x => ({
  Campo: x.campo, Cluster: x.cluster, Pozo: x.pozo, Estado: x.estado, 'Última medición': fmtDate(x.ultima),
  'Días sin medir': x.estado === 'medido' ? '' : x.dias, 'N° pruebas': x.nPruebas, 'Tiempo estimado (h)': +x.ciclo.toFixed(2),
  'Base del tiempo': x.fuente, 'Crudo (BOPD)': x.crudo, 'Líquido (BFPD)': x.liq, 'BSW %': x.bsw, 'Último equipo': x.equipo
});

/* ================= Eventos ================= */
document.addEventListener('DOMContentLoaded', () => {
  aplicarConfig();
  const drop = $('#drop');
  $('#file').addEventListener('change', e => { cargarArchivos([...e.target.files]); e.target.value = ''; });
  ['dragenter', 'dragover'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.add('over'); }));
  ['dragleave', 'drop'].forEach(ev => drop.addEventListener(ev, e => { e.preventDefault(); drop.classList.remove('over'); }));
  drop.addEventListener('drop', e => cargarArchivos([...e.dataTransfer.files]));
  $('#archivos').addEventListener('click', e => {
    const n = e.target.dataset && e.target.dataset.quitar;
    if (n) { state.archivos.delete(n); refrescarArchivos(); }
  });

  $('#f-campo').addEventListener('change', () => { state.abiertos.clear(); poblarMeses(); render(); });
  $('#f-mes').addEventListener('change', () => { actualizarEquiposAuto(); render(); });
  ['#f-ref', '#f-inact', '#f-hpd', '#f-neq', '#f-mov', '#f-ini', '#f-prio', '#soloPend', '#f-estado'].forEach(id => $(id).addEventListener('change', render));
  $('#buscar').addEventListener('input', () => state.result && renderPozos(state.result));
  $('#exportar').addEventListener('click', exportar);

  document.querySelector('.tabs').addEventListener('click', e => {
    const b = e.target.closest('button[data-tab]'); if (!b) return;
    document.querySelectorAll('.tabs button').forEach(x => x.classList.toggle('on', x === b));
    document.querySelectorAll('.tab').forEach(t => { t.hidden = t.id !== 'tab-' + b.dataset.tab; });
  });
  $('#tblClusters').addEventListener('click', e => {
    const tr = e.target.closest('tr.cl'); if (!tr) return;
    const k = tr.dataset.cl;
    state.abiertos.has(k) ? state.abiertos.delete(k) : state.abiertos.add(k);
    renderClusters(state.result);
  });
  document.addEventListener('click', e => {
    const t = e.target.closest('th[data-sort]'); if (!t) return;
    const [tabla, col] = t.dataset.sort.split(':');
    const s = state.sort[tabla];
    state.sort[tabla] = s && s.col === col ? (s.dir < 0 ? { col, dir: 1 } : null) : { col, dir: -1 };
    tabla === 'pozos' ? renderPozos(state.result) : renderClusters(state.result);
  });
});

// Para pruebas: window.cargarArchivos(files)
window.cargarArchivos = cargarArchivos;
