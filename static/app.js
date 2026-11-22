/* ============================================================
   Puerto A&M · Business Intelligence
   El frontend NO contiene datos: todo llega desde /api/dashboard,
   que calcula las cifras leyendo el archivo Excel con openpyxl.
   ============================================================ */
'use strict';

const $ = (id) => document.getElementById(id);

const state = {
  filters: { fecha_desde: '', fecha_hasta: '', producto: '', categoria: '', canal: '' },
  data: null,
  charts: {},
  page: 1,
  pageSize: 10,
  sort: { key: 'fecha', dir: 'desc' },
  search: '',
  loading: false,
  firstLoad: true,
  chat: { busy: false }
};

const CHAT_GREETING = (
  'Hola. Soy A&M Copiloto, analista de datos de Puerto A&M. Respondo únicamente con lo que ' +
  'calculamos desde el Excel cargado. Pregúntame por ventas, pedidos, metas, categorías o productos.'
);

const fmtMoney0 = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', maximumFractionDigits: 0 });
const fmtMoney2 = new Intl.NumberFormat('es-MX', { style: 'currency', currency: 'MXN', minimumFractionDigits: 2, maximumFractionDigits: 2 });
const fmtNumber = new Intl.NumberFormat('es-MX', { maximumFractionDigits: 0 });
const fmtPct = new Intl.NumberFormat('es-MX', { minimumFractionDigits: 1, maximumFractionDigits: 1 });

const cssVar = (name) => getComputedStyle(document.documentElement).getPropertyValue(name).trim();
const COLORS = () => ({
  petroleum: cssVar('--petroleum-3') || '#14607a',
  deep: cssVar('--petroleum') || '#0b3948',
  gold: cssVar('--gold') || '#c7a03c',
  gold2: cssVar('--gold-2') || '#e0bd63',
  green: cssVar('--green') || '#1d8a67',
  red: cssVar('--red') || '#c45d45',
  grid: cssVar('--chart-grid') || '#e9eef1',
  text: cssVar('--chart-text') || '#83939c'
});

const esc = (value) => String(value ?? '').replace(/[&<>"']/g, (ch) => ({
  '&': '&amp;', '<': '&lt;', '>': '&gt;', '"': '&quot;', "'": '&#39;'
}[ch]));

const money = (value) => fmtMoney0.format(Number(value) || 0);
const money2 = (value) => fmtMoney2.format(Number(value) || 0);
const pct = (value) => value === null || value === undefined || !Number.isFinite(Number(value)) ? '—' : `${fmtPct.format(Number(value))}%`;

/* ---------- API ---------- */
function filterQuery() {
  return new URLSearchParams(Object.entries(state.filters).filter(([, value]) => value !== '' && value !== null));
}

async function api(path, options) {
  const response = await fetch(path, options);
  const payload = await response.json().catch(() => ({}));
  if (!response.ok) throw new Error(payload.detail || `Error HTTP ${response.status}`);
  return payload;
}

async function loadData(force = false) {
  if (state.loading) return;
  state.loading = true;
  const button = $('refreshButton');
  if (button) {
    button.disabled = true;
    button.querySelector('span').textContent = force ? 'Actualizando…' : 'Cargando…';
  }

  try {
    if (force) await api('/api/refresh', { method: 'POST' });
    const query = filterQuery();
    const payload = await api(`/api/dashboard?${query.toString()}`);
    state.data = payload;

    if (state.firstLoad) {
      populateFilters(payload.options);
      // La primera petición no tenía filtros: mostramos el rango completo ya aplicado.
      payload.filters_applied.fecha_desde = payload.options.fecha_min;
      payload.filters_applied.fecha_hasta = payload.options.fecha_max;
      state.firstLoad = false;
    }
    $('errorBanner').hidden = true;
    renderAll();
    renderChatScope();
    if (force && !state.firstLoad) {
      resetChat(true);
      showToast(`Excel leído de nuevo · ${payload.meta.modified_label}`, 'success');
    }
  } catch (error) {
    $('errorDetail').textContent = error.message;
    $('errorBanner').hidden = false;
    setStatus('error', 'Sin conexión con el Excel');
  } finally {
    state.loading = false;
    if (button) {
      button.disabled = false;
      button.querySelector('span').textContent = 'Actualizar datos';
    }
  }
}

function populateFilters(options) {
  $('dateFrom').min = options.fecha_min;
  $('dateFrom').max = options.fecha_max;
  $('dateTo').min = options.fecha_min;
  $('dateTo').max = options.fecha_max;

  const fill = (id, values, label) => {
    const select = $(id);
    select.innerHTML = `<option value="">${label}</option>`;
    values.forEach((value) => {
      const option = document.createElement('option');
      option.value = value;
      option.textContent = value;
      select.appendChild(option);
    });
  };
  fill('productFilter', options.productos, 'Todos los productos');
  fill('categoryFilter', options.categorias, 'Todas las categorías');
  fill('channelFilter', options.canales, 'Todos los canales');

  state.filters.fecha_desde = options.fecha_min;
  state.filters.fecha_hasta = options.fecha_max;
  $('dateFrom').value = options.fecha_min;
  $('dateTo').value = options.fecha_max;
}

/* ---------- Render ---------- */
function renderAll() {
  const data = state.data;
  if (!data) return;
  renderStatus(data.meta);
  renderFacts(data);
  renderKPIs(data.kpis);
  renderAttainment(data);
  renderCharts(data.charts);
  renderTable();
  renderSource(data.meta);
}

function setStatus(kind, text) {
  const status = $('sourceStatus');
  status.classList.remove('ok', 'error');
  if (kind) status.classList.add(kind);
  $('statusDot').style.background = '';
  $('sourceLabel').textContent = text;
}

function renderStatus(meta) {
  setStatus('ok', `Excel conectado · ${meta.modified_label}`);
  $('sourceStatus').title = `Archivo ${meta.name || meta.file} · leído ${meta.server_time} · ${meta.cache_hit ? 'caché vigente' : 'lectura nueva'}`;
}

function renderFacts(data) {
  const f = data.filters_applied;
  const formatDate = (iso) => {
    if (!iso) return '—';
    const [y, m, d] = iso.split('-').map(Number);
    return `${String(d).padStart(2, '0')}/${String(m).padStart(2, '0')}/${y}`;
  };
  $('factPeriod').textContent = `${formatDate(f.fecha_desde)} — ${formatDate(f.fecha_hasta)}`;
  $('factRecords').textContent = fmtNumber.format(data.kpis.filas_filtradas);
}

function renderKPIs(k) {
  $('kpiIngresos').textContent = money(k.ingresos);
  $('kpiIngresosSub').textContent = `${fmtNumber.format(k.filas_filtradas)} registros del Excel`;
  $('kpiPedidos').textContent = fmtNumber.format(k.pedidos);
  $('kpiTicket').textContent = money(k.ticket_promedio);
  $('kpiUnidades').textContent = fmtNumber.format(k.unidades);
  $('kpiCostos').textContent = money(k.costos);
  $('kpiUtilidad').textContent = money(k.utilidad);
  $('kpiMargen').textContent = pct(k.margen);

  const growth = $('kpiCrecimiento');
  growth.textContent = pct(k.crecimiento);
  growth.classList.toggle('positive', k.crecimiento > 0);
  growth.classList.toggle('negative', k.crecimiento < 0);
  $('growthTag').textContent = k.crecimiento > 0 ? '▲ Crecimiento' : k.crecimiento < 0 ? '▼ Decrecimiento' : 'Mgto.';
  $('kpiCrecimientoSub').textContent = k.crecimiento === null
    ? 'Se necesitan dos meses en el filtro'
    : 'Último mes vs. mes anterior';
}

function renderAttainment(data) {
  const metas = data.charts.metas || [];
  const real = metas.reduce((sum, item) => sum + item.real, 0);
  const target = metas.reduce((sum, item) => sum + item.meta, 0);
  const value = data.kpis.cumplimiento_metas;
  $('attainmentValue').textContent = pct(value);
  $('attainmentTitle').textContent = target > 0
    ? `${money(real)} de ${money(target)} planificados`
    : 'Sin metas en el periodo filtrado';
  $('attainmentSub').textContent = target > 0
    ? `${metas.length} mes${metas.length === 1 ? '' : 'es'} con meta en la hoja Metas del Excel.`
    : 'Agrega metas mensuales al Excel para ver el cumplimiento.';
  const width = value === null ? 0 : Math.min(100, Math.max(0, (value / 150) * 100));
  $('attainmentFill').style.width = `${width}%`;
}

function ensureChart(id, config) {
  const canvas = $(id);
  const empty = canvas.parentElement.querySelector('.chart-empty');
  const hasData = config.data.labels && config.data.labels.length > 0;
  canvas.style.display = hasData ? '' : 'none';
  if (empty) empty.hidden = hasData;
  if (!hasData) {
    if (state.charts[id]) { state.charts[id].destroy(); delete state.charts[id]; }
    return;
  }
  if (state.charts[id]) state.charts[id].destroy();
  state.charts[id] = new Chart(canvas, config);
}

function baseOptions(c = COLORS()) {
  return {
    responsive: true,
    maintainAspectRatio: false,
    animation: { duration: 550, easing: 'easeOutQuart' },
    interaction: { mode: 'index', intersect: false },
    plugins: {
      legend: { display: false },
      tooltip: {
        backgroundColor: c.deep,
        titleColor: '#fff',
        bodyColor: '#e6f1f4',
        padding: 11,
        cornerRadius: 8,
        displayColors: true,
        boxWidth: 8,
        boxHeight: 8,
        titleFont: { size: 11, weight: '600' },
        bodyFont: { size: 11 }
      }
    },
    scales: {
      x: {
        grid: { display: false },
        border: { color: c.grid },
        ticks: { color: c.text, font: { size: 10 } }
      },
      y: {
        grid: { color: c.grid, drawTicks: false },
        border: { display: false },
        ticks: { color: c.text, font: { size: 10 }, padding: 8 }
      }
    }
  };
}

function renderCharts(charts) {
  const c = COLORS();
  if (window.Chart) {
    Chart.defaults.font.family = getComputedStyle(document.body).fontFamily;
    Chart.defaults.color = c.text;
  } else {
    return;
  }

  /* 1. Evolución de ventas (ingresos + pedidos) */
  const trend = charts.tendencia || [];
  const trendOptions = baseOptions(c);
  trendOptions.scales.y.ticks.callback = (value) => money(value);
  trendOptions.scales.y1 = {
    position: 'right', grid: { display: false }, border: { display: false },
    ticks: { color: c.text, font: { size: 10 }, precision: 0 }
  };
  trendOptions.plugins.tooltip.callbacks = {
    label: (ctx) => ctx.dataset.yAxisID === 'y1'
      ? ` Pedidos: ${fmtNumber.format(ctx.parsed.y)}`
      : ` Ingresos: ${money(ctx.parsed.y)}`
  };
  ensureChart('trendChart', {
    data: {
      labels: trend.map((item) => item.label),
      datasets: [
        {
          type: 'bar', label: 'Pedidos', data: trend.map((item) => item.pedidos),
          yAxisID: 'y1', backgroundColor: c.petroleum + '38', hoverBackgroundColor: c.petroleum + '66',
          borderRadius: 5, barPercentage: .45, order: 2
        },
        {
          type: 'line', label: 'Ingresos', data: trend.map((item) => item.ingresos),
          yAxisID: 'y', borderColor: c.gold, backgroundColor: c.gold + '22',
          pointBackgroundColor: c.gold, pointBorderColor: '#fff', pointBorderWidth: 2,
          pointRadius: 4, pointHoverRadius: 6, borderWidth: 2.5, tension: .38, fill: true, order: 1
        }
      ]
    },
    options: trendOptions
  });

  /* 2. Ingresos vs. costos */
  const rc = charts.ingresos_costos || [];
  const rcOptions = baseOptions(c);
  rcOptions.scales.y.ticks.callback = (value) => money(value);
  rcOptions.plugins.tooltip.callbacks = { label: (ctx) => ` ${ctx.dataset.label}: ${money(ctx.parsed.y)}` };
  ensureChart('revenueCostChart', {
    type: 'bar',
    data: {
      labels: rc.map((item) => item.label),
      datasets: [
        { label: 'Ingresos', data: rc.map((item) => item.ingresos), backgroundColor: c.deep, borderRadius: 5, maxBarThickness: 40 },
        { label: 'Costos', data: rc.map((item) => item.costos), backgroundColor: c.gold, borderRadius: 5, maxBarThickness: 40 }
      ]
    },
    options: rcOptions
  });

  /* 3. Ventas por categoría */
  const cats = charts.categorias || [];
  const catOptions = baseOptions(c);
  delete catOptions.scales;
  catOptions.cutout = '68%';
  catOptions.plugins.legend = {
    display: true, position: 'right',
    labels: { color: c.text, boxWidth: 9, boxHeight: 9, usePointStyle: true, pointStyle: 'rectRounded', font: { size: 10 }, padding: 14 }
  };
  catOptions.plugins.tooltip.callbacks = { label: (ctx) => ` ${ctx.label}: ${money(ctx.parsed)}` };
  ensureChart('categoryChart', {
    type: 'doughnut',
    data: {
      labels: cats.map((item) => item.label),
      datasets: [{
        data: cats.map((item) => item.value),
        backgroundColor: [c.deep, c.gold, c.petroleum, c.gold2, '#5b8b9c', '#a7c3cc', '#8a7330'],
        borderColor: cssVar('--surface'), borderWidth: 3, hoverOffset: 6
      }]
    },
    options: catOptions
  });

  /* 4. Productos destacados */
  const products = (charts.productos || []).slice(0, 8);
  const prodOptions = baseOptions(c);
  delete prodOptions.scales.y1;
  prodOptions.indexAxis = 'y';
  prodOptions.scales.x.ticks.callback = (value) => money(value);
  prodOptions.plugins.tooltip.callbacks = {
    label: (ctx) => ` Ingresos: ${money(ctx.parsed.x)} · ${fmtNumber.format(products[ctx.dataIndex].unidades)} uds.`
  };
  ensureChart('productsChart', {
    type: 'bar',
    data: {
      labels: products.map((item) => item.label.length > 28 ? item.label.slice(0, 27) + '…' : item.label),
      datasets: [{
        label: 'Ingresos', data: products.map((item) => item.value),
        backgroundColor: products.map((_, index) => index === 0 ? c.gold : index === 1 ? c.gold2 : c.petroleum),
        borderRadius: 5, maxBarThickness: 20
      }]
    },
    options: prodOptions
  });

  /* 5. Metas vs. resultados */
  const goals = charts.metas || [];
  const goalOptions = baseOptions(c);
  goalOptions.scales.y.ticks.callback = (value) => money(value);
  goalOptions.plugins.tooltip.callbacks = {
    afterBody: (items) => {
      const item = goals[items[0].dataIndex];
      return item && item.cumplimiento !== null ? `Cumplimiento: ${pct(item.cumplimiento)}` : '';
    },
    label: (ctx) => ` ${ctx.dataset.label}: ${money(ctx.parsed.y)}`
  };
  ensureChart('goalsChart', {
    type: 'bar',
    data: {
      labels: goals.map((item) => item.label),
      datasets: [
        { label: 'Meta', data: goals.map((item) => item.meta), backgroundColor: '#9db3bd', borderRadius: 5, maxBarThickness: 46 },
        { label: 'Real', data: goals.map((item) => item.real), backgroundColor: c.gold, borderRadius: 5, maxBarThickness: 46 }
      ]
    },
    options: goalOptions
  });

  /* 6. Ventas por canal */
  const channels = charts.canales || [];
  const channelOptions = baseOptions(c);
  delete channelOptions.scales.y1;
  channelOptions.indexAxis = 'y';
  channelOptions.scales.x.ticks.callback = (value) => money(value);
  channelOptions.plugins.tooltip.callbacks = { label: (ctx) => ` ${money(ctx.parsed.x)}` };
  ensureChart('channelsChart', {
    type: 'bar',
    data: {
      labels: channels.map((item) => item.label),
      datasets: [{
        data: channels.map((item) => item.value),
        backgroundColor: [c.deep, c.petroleum, c.gold, c.gold2, '#5b8b9c'],
        borderRadius: 5, maxBarThickness: 26
      }]
    },
    options: channelOptions
  });
}

/* ---------- Tabla ---------- */
function tableRows() {
  let rows = [...(state.data?.table || [])];
  const search = state.search.trim().toLocaleLowerCase('es');
  if (search) {
    rows = rows.filter((row) => [row.pedido, row.producto, row.categoria, row.canal]
      .some((value) => String(value).toLocaleLowerCase('es').includes(search)));
  }
  const { key, dir } = state.sort;
  const factor = dir === 'asc' ? 1 : -1;
  rows.sort((a, b) => {
    const av = a[key], bv = b[key];
    if (typeof av === 'number' && typeof bv === 'number') return (av - bv) * factor;
    return String(av).localeCompare(String(bv), 'es') * factor;
  });
  return rows;
}

function renderTable() {
  const rows = tableRows();
  const totalPages = Math.max(1, Math.ceil(rows.length / state.pageSize));
  state.page = Math.min(state.page, totalPages);
  const start = (state.page - 1) * state.pageSize;
  const visible = rows.slice(start, start + state.pageSize);

  $('salesBody').innerHTML = visible.length ? visible.map((row) => `
    <tr>
      <td>${esc(row.fecha.split('-').reverse().join('/'))}</td>
      <td>${esc(row.pedido)}</td>
      <td title="${esc(row.producto)}">${esc(row.producto)}</td>
      <td>${esc(row.categoria)}</td>
      <td class="num">${fmtNumber.format(row.cantidad)}</td>
      <td class="num">${money2(row.ingreso)}</td>
      <td class="num">${money2(row.costo)}</td>
      <td class="num ${row.utilidad < 0 ? 'negative' : 'positive'}">${money2(row.utilidad)}</td>
      <td>${esc(row.canal)}</td>
    </tr>`).join('') : '<tr><td colspan="9" class="loading-row">No hay ventas que coincidan con los filtros.</td></tr>';

  $('tableSummary').textContent = `${fmtNumber.format(rows.length)} registros después de aplicar filtros y búsqueda.`;
  $('tableInfo').textContent = rows.length
    ? `Mostrando ${fmtNumber.format(start + 1)}–${fmtNumber.format(Math.min(start + state.pageSize, rows.length))} de ${fmtNumber.format(rows.length)}`
    : 'Sin resultados';
  $('pageLabel').textContent = `${state.page} / ${totalPages}`;
  $('prevPage').disabled = state.page <= 1;
  $('nextPage').disabled = state.page >= totalPages;
  document.querySelectorAll('thead th[data-sort]').forEach((th) => {
    th.classList.toggle('sorted', th.dataset.sort === state.sort.key);
  });
}

function renderSource(meta) {
  $('fileName').textContent = meta.file;
  $('sourceFile').textContent = meta.file;
  $('sourceModified').textContent = meta.modified_label;
  $('sourceRows').textContent = fmtNumber.format(meta.rows_total);
  $('sourceCounts').textContent = `${meta.products_total} / ${meta.goals_total}`;
  $('sourceCache').textContent = meta.cache_hit ? 'Caché vigente' : 'Lectura nueva';
}

/* ---------- Asistente IA (siempre fundamentado en el Excel) ---------- */
function appendChatMessage(role, text, meta = {}) {
  const wrap = document.createElement('div');
  wrap.className = `chat-message${role === 'user' ? ' mine' : ''}`;
  wrap.innerHTML = `<div class="chat-bubble-avatar"><svg class="icon"><use href="#i-${role === 'user' ? 'chat' : 'spark'}"/></svg></div>`;

  const bubble = document.createElement('div');
  bubble.className = `chat-bubble${meta.missing ? ' is-missing' : ''}`;
  const author = document.createElement('span');
  author.className = 'chat-author';
  author.textContent = role === 'user' ? 'TÚ' : 'A&M COPILOTO';
  const body = document.createElement('div');
  body.textContent = text; // CSS white-space: pre-wrap conserva los saltos de línea.
  bubble.append(author, body);

  if (meta.sources && meta.sources.length) {
    const sources = document.createElement('div');
    sources.className = 'chat-sources';
    meta.sources.forEach((source) => {
      const chip = document.createElement('span');
      chip.className = 'chat-source';
      chip.textContent = source;
      sources.appendChild(chip);
    });
    bubble.appendChild(sources);
  }

  wrap.appendChild(bubble);
  const log = $('chatMessages');
  log.appendChild(wrap);
  log.scrollTop = log.scrollHeight;
}

function setChatTyping(visible) {
  const previous = $('chatMessages').querySelector('.chat-typing-message');
  if (previous) previous.remove();
  if (!visible) return;
  const wrap = document.createElement('div');
  wrap.className = 'chat-message chat-typing-message';
  wrap.innerHTML = '<div class="chat-bubble-avatar"><svg class="icon"><use href="#i-spark"/></svg></div>'
    + '<div class="chat-bubble"><span class="chat-author">A&M COPILOTO</span>'
    + '<div class="chat-typing"><i></i><i></i><i></i></div></div>';
  const log = $('chatMessages');
  log.appendChild(wrap);
  log.scrollTop = log.scrollHeight;
}

function resetChat(notifyRefresh = false) {
  $('chatMessages').innerHTML = '';
  appendChatMessage('assistant', notifyRefresh
    ? 'Los datos del Excel se actualizaron y reinicié la conversación. Soy A&M Copiloto: vuelve a preguntar, las cifras son las nuevas.'
    : CHAT_GREETING);
  setChatTyping(false);
  $('chatInput').value = '';
  $('chatInput').style.height = 'auto';
}

function renderChatScope() {
  const f = state.data?.filters_applied || {};
  const parts = [];
  if (f.fecha_desde || f.fecha_hasta) parts.push(`${f.fecha_desde || '…'} → ${f.fecha_hasta || '…'}`);
  if (f.producto) parts.push(f.producto);
  if (f.categoria) parts.push(f.categoria);
  if (f.canal) parts.push(f.canal);
  $('chatScope').textContent = parts.length
    ? `Analista de datos · filtros: ${parts.join(' · ')}`
    : 'Analista de datos · respuesta fundamentada';
  $('chatFileName').textContent = state.data?.meta?.file || 'Excel';
}

async function sendChat(question) {
  const text = String(question || '').trim();
  if (!text || state.chat.busy) return;

  appendChatMessage('user', text);
  $('chatInput').value = '';
  $('chatInput').style.height = 'auto';

  state.chat.busy = true;
  $('chatSend').disabled = true;
  document.querySelectorAll('#chatSuggestions button').forEach((button) => { button.disabled = true; });
  setChatTyping(true);

  try {
    const payload = await api('/api/chat', {
      method: 'POST',
      headers: { 'Content-Type': 'application/json' },
      body: JSON.stringify({ mensaje: text, ...state.filters }),
    });
    appendChatMessage('assistant', payload.respuesta, {
      sources: payload.fuentes,
      missing: /No tengo ese dato|No hay registros|No puedo calcular|No hay metas/.test(payload.respuesta),
    });
    $('chatMode').textContent = payload.verificado
      ? 'Motor local · cifras verificadas'
      : 'Modelo externo · contexto cerrado';
  } catch (error) {
    appendChatMessage('assistant', `No pude consultar el Excel en este momento: ${error.message}`, { missing: true });
  } finally {
    state.chat.busy = false;
    $('chatSend').disabled = false;
    document.querySelectorAll('#chatSuggestions button').forEach((button) => { button.disabled = false; });
    setChatTyping(false);
    $('chatInput').focus({ preventScroll: true });
  }
}

function initChat() {
  resetChat(false);

  $('chatForm').addEventListener('submit', (event) => {
    event.preventDefault();
    sendChat($('chatInput').value);
  });

  $('chatInput').addEventListener('keydown', (event) => {
    if (event.key === 'Enter' && !event.shiftKey) {
      event.preventDefault();
      $('chatForm').requestSubmit();
    }
  });

  $('chatInput').addEventListener('input', (event) => {
    const input = event.target;
    input.style.height = 'auto';
    input.style.height = `${Math.min(120, input.scrollHeight)}px`;
  });

  document.querySelectorAll('#chatSuggestions button').forEach((button) => {
    button.addEventListener('click', () => sendChat(button.dataset.question));
  });

  $('chatReset').addEventListener('click', () => resetChat(false));

  $('chatFab').addEventListener('click', () => {
    $('asistente').scrollIntoView({ behavior: 'smooth', block: 'start' });
    setTimeout(() => $('chatInput').focus({ preventScroll: true }), 550);
  });
}

/* ---------- Acciones ---------- */
function showToast(message, type = '') {
  const toast = $('toast');
  toast.className = `toast ${type}`;
  toast.textContent = message;
  toast.hidden = false;
  clearTimeout(showToast.timer);
  showToast.timer = setTimeout(() => { toast.hidden = true; }, 4200);
}

function readFiltersFromInputs() {
  state.filters.fecha_desde = $('dateFrom').value;
  state.filters.fecha_hasta = $('dateTo').value;
  state.filters.producto = $('productFilter').value;
  state.filters.categoria = $('categoryFilter').value;
  state.filters.canal = $('channelFilter').value;
  state.page = 1;
}

function applyFilters() {
  readFiltersFromInputs();
  const all = state.data?.options;
  if (all) {
    const active = Object.entries(state.filters)
      .filter(([key, value]) => !['producto', 'categoria', 'canal'].includes(key) && value)
      .every(([key, value]) => value === all[key === 'fecha_desde' ? 'fecha_min' : 'fecha_max']);
    document.querySelectorAll('#quickRanges button').forEach((btn) => {
      btn.classList.toggle('is-active', btn.dataset.range === 'all' && active);
    });
  }
  loadData(false);
}

async function exportReport(format) {
  try {
    const query = filterQuery();
    query.set('formato', format);
    const response = await fetch(`/api/export?${query.toString()}`);
    if (!response.ok) {
      const error = await response.json().catch(() => ({}));
      throw new Error(error.detail || 'No se pudo generar el reporte');
    }
    const blob = await response.blob();
    const url = URL.createObjectURL(blob);
    const anchor = document.createElement('a');
    anchor.href = url;
    anchor.download = `PuertoAM_Reporte_Ventas.${format}`;
    document.body.appendChild(anchor);
    anchor.click();
    anchor.remove();
    URL.revokeObjectURL(url);
    showToast(`Reporte ${format.toUpperCase()} generado desde el Excel.`, 'success');
  } catch (error) {
    showToast(error.message, 'error');
  }
}

function toggleTheme() {
  const next = document.documentElement.dataset.theme === 'dark' ? 'light' : 'dark';
  document.documentElement.dataset.theme = next;
  localStorage.setItem('puertoam.theme', next);
  $('themeButton').querySelector('use').setAttribute('href', next === 'dark' ? '#i-sun' : '#i-moon');
  if (state.data) renderCharts(state.data.charts);
}

function bindEvents() {
  $('refreshButton').addEventListener('click', () => loadData(true));
  $('retryButton').addEventListener('click', () => loadData(false));
  $('applyFilters').addEventListener('click', applyFilters);
  $('clearFilters').addEventListener('click', () => {
    const options = state.data?.options;
    $('dateFrom').value = options?.fecha_min || '';
    $('dateTo').value = options?.fecha_max || '';
    $('productFilter').value = '';
    $('categoryFilter').value = '';
    $('channelFilter').value = '';
    state.search = '';
    $('tableSearch').value = '';
    document.querySelectorAll('#quickRanges button').forEach((b) => b.classList.toggle('is-active', b.dataset.range === 'all'));
    applyFilters();
  });

  ['dateFrom', 'dateTo', 'productFilter', 'categoryFilter', 'channelFilter'].forEach((id) => {
    $(id).addEventListener('change', applyFilters);
  });

  document.querySelectorAll('#quickRanges button').forEach((button) => {
    button.addEventListener('click', () => {
      const options = state.data?.options;
      if (!options) return;
      const range = button.dataset.range;
      if (range === 'all') {
        $('dateFrom').value = options.fecha_min;
        $('dateTo').value = options.fecha_max;
      } else {
        const from = `${range}-01`;
        const lastDay = new Date(Number(range.slice(0, 4)), Number(range.slice(5, 7)), 0).getDate();
        $('dateFrom').value = from;
        $('dateTo').value = `${range}-${String(lastDay).padStart(2, '0')}`;
      }
      document.querySelectorAll('#quickRanges button').forEach((b) => b.classList.toggle('is-active', b === button));
      applyFilters();
    });
  });

  $('tableSearch').addEventListener('input', (event) => {
    state.search = event.target.value;
    state.page = 1;
    renderTable();
  });

  document.querySelectorAll('thead th[data-sort]').forEach((th) => {
    th.addEventListener('click', () => {
      const key = th.dataset.sort;
      if (state.sort.key === key) state.sort.dir = state.sort.dir === 'asc' ? 'desc' : 'asc';
      else state.sort = { key, dir: 'asc' };
      state.page = 1;
      renderTable();
    });
  });

  $('prevPage').addEventListener('click', () => { state.page--; renderTable(); });
  $('nextPage').addEventListener('click', () => { state.page++; renderTable(); });
  $('exportCsv').addEventListener('click', () => exportReport('csv'));
  $('exportXlsx').addEventListener('click', () => exportReport('xlsx'));
  $('themeButton').addEventListener('click', toggleTheme);

  $('menuButton').addEventListener('click', () => {
    $('sidebar').classList.add('is-open');
    $('sidebarBackdrop').hidden = false;
  });
  const closeSidebar = () => {
    $('sidebar').classList.remove('is-open');
    $('sidebarBackdrop').hidden = true;
  };
  $('sidebarClose').addEventListener('click', closeSidebar);
  $('sidebarBackdrop').addEventListener('click', closeSidebar);

  document.querySelectorAll('.nav-item').forEach((link) => {
    link.addEventListener('click', () => {
      document.querySelectorAll('.nav-item').forEach((item) => item.classList.toggle('is-active', item === link));
      $('crumbTitle').textContent = link.querySelector('span').textContent;
      closeSidebar();
    });
  });

  const sections = ['resumen', 'ventas', 'productos', 'metas', 'asistente', 'fuente'];
  const observer = new IntersectionObserver((entries) => {
    const visible = entries.filter((entry) => entry.isIntersecting).sort((a, b) => b.intersectionRatio - a.intersectionRatio)[0];
    if (!visible) return;
    const link = document.querySelector(`.nav-item[data-section="${visible.target.id}"]`);
    if (!link) return;
    document.querySelectorAll('.nav-item').forEach((item) => item.classList.toggle('is-active', item === link));
    $('crumbTitle').textContent = link.querySelector('span').textContent;
  }, { rootMargin: '-80px 0px -60% 0px', threshold: [0, .1, .5] });
  sections.forEach((id) => { const node = $(id); if (node) observer.observe(node); });
}

function initTheme() {
  const saved = localStorage.getItem('puertoam.theme');
  const theme = saved === 'light' || saved === 'dark'
    ? saved
    : (window.matchMedia('(prefers-color-scheme: dark)').matches ? 'dark' : 'light');
  document.documentElement.dataset.theme = theme;
  $('themeButton').querySelector('use').setAttribute('href', theme === 'dark' ? '#i-sun' : '#i-moon');
}

initTheme();
bindEvents();
initChat();
loadData(false);
