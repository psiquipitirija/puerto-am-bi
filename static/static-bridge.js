/* Puerto A&M · puente para GitHub Pages.
   Habla con el mismo Excel exportado a data/dashboard.json y calcula en el
   navegador todo lo que antes calculaba FastAPI: filtros, KPIs, gráficas,
   exportaciones y respuestas del chat. No hay backend ni claves aquí. */
(() => {
  const realFetch = window.fetch.bind(window);
  let sourcePromise = null;

  const MONTHS = {
    enero: '01', febrero: '02', marzo: '03', abril: '04', mayo: '05', junio: '06',
    julio: '07', agosto: '08', septiembre: '09', octubre: '10', noviembre: '11', diciembre: '12',
    ene: '01', feb: '02', mar: '03', abr: '04', may: '05', jun: '06',
    jul: '07', ago: '08', sep: '09', oct: '10', nov: '11', dic: '12',
  };
  const MONTH_LABELS = ['Ene', 'Feb', 'Mar', 'Abr', 'May', 'Jun', 'Jul', 'Ago', 'Sep', 'Oct', 'Nov', 'Dic'];
  const SIN_DATO = 'No tengo ese dato en el Excel cargado. Solo respondo con información de las hojas Ventas, Productos y Metas. Puedo ayudarte con ventas, pedidos, ticket promedio, unidades, costos, utilidad, margen, crecimiento, metas, categorías, productos, canales, métodos de pago y con la venta más grande o más pequeña calculada de cada fila.';

  const norm = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase().replace(/\s+/g, '');
  const plain = (value) => String(value ?? '').normalize('NFD').replace(/[\u0300-\u036f]/g, '').toLowerCase();
  const has = (text, phrase) => text.includes(norm(phrase));
  const hasAny = (text, phrases) => phrases.some((phrase) => has(text, phrase));
  const moneyFmt = new Intl.NumberFormat('es-MX', { maximumFractionDigits: 0 });
  const numFmt = new Intl.NumberFormat('es-MX', { maximumFractionDigits: 0 });
  const pctFmt = new Intl.NumberFormat('es-MX', { minimumFractionDigits: 1, maximumFractionDigits: 1 });
  const money = (value) => '$' + moneyFmt.format(Number(value) || 0);
  const number = (value) => numFmt.format(Number(value) || 0);
  const pct = (value) => value === null || value === undefined || !Number.isFinite(Number(value)) ? '—' : pctFmt.format(Number(value)) + '%';
  const round2 = (value) => Math.round((Number(value) + Number.EPSILON) * 100) / 100;
  const sum = (rows, key) => round2(rows.reduce((acc, row) => acc + Number(row[key] || 0), 0));

  function getSource() {
    if (!sourcePromise) {
      sourcePromise = realFetch('data/dashboard.json').then((response) => {
        if (!response.ok) throw new Error('No se pudo leer data/dashboard.json');
        return response.json();
      });
    }
    return sourcePromise;
  }

  function filteredRows(src, params) {
    const from = params.get('fecha_desde') || '';
    const to = params.get('fecha_hasta') || '';
    const product = params.get('producto') || '';
    const category = params.get('categoria') || '';
    const channel = params.get('canal') || '';
    return src.rows.filter((row) =>
      (!from || row.fecha >= from) &&
      (!to || row.fecha <= to) &&
      (!product || norm(row.producto) === norm(product)) &&
      (!category || norm(row.categoria) === norm(category)) &&
      (!channel || norm(row.canal) === norm(channel)));
  }

  function buildDashboard(src, params) {
    const rows = filteredRows(src, params);
    const ingresos = sum(rows, 'ingreso');
    const costos = sum(rows, 'costo');
    const utilidad = round2(ingresos - costos);
    const unidades = rows.reduce((acc, row) => acc + Number(row.cantidad || 0), 0);
    const pedidos = new Set(rows.map((row) => row.pedido)).size;
    const margen = ingresos > 0 ? utilidad / ingresos * 100 : null;

    const byMonth = new Map();
    for (const row of rows) {
      const key = row.fecha.slice(0, 7);
      if (!byMonth.has(key)) {
        const month = Number(key.slice(5, 7));
        byMonth.set(key, {
          key, year: Number(key.slice(0, 4)), month,
          label: `${MONTH_LABELS[month - 1]} ${key.slice(0, 4)}`,
          ingresos: 0, costos: 0, utilidad: 0, unidades: 0, orders: new Set(),
        });
      }
      const bucket = byMonth.get(key);
      bucket.ingresos += Number(row.ingreso || 0);
      bucket.costos += Number(row.costo || 0);
      bucket.utilidad += Number(row.utilidad || 0);
      bucket.unidades += Number(row.cantidad || 0);
      bucket.orders.add(row.pedido);
    }
    const monthly = [...byMonth.values()].sort((a, b) => a.key.localeCompare(b.key)).map((m) => ({
      key: m.key, label: m.label,
      ingresos: round2(m.ingresos), pedidos: m.orders.size,
      costos: round2(m.costos), utilidad: round2(m.utilidad), unidades: m.unidades,
    }));
    const growth = monthly.length >= 2 && monthly[monthly.length - 2].ingresos > 0
      ? (monthly[monthly.length - 1].ingresos / monthly[monthly.length - 2].ingresos - 1) * 100
      : null;

    const byProduct = new Map();
    const byCategory = new Map();
    const byChannel = new Map();
    const byPayment = new Map();
    for (const row of rows) {
      const p = byProduct.get(row.producto) || { producto: row.producto, ingresos: 0, unidades: 0, utilidad: 0 };
      p.ingresos += Number(row.ingreso || 0); p.unidades += Number(row.cantidad || 0); p.utilidad += Number(row.utilidad || 0);
      byProduct.set(row.producto, p);
      const c = byCategory.get(row.categoria) || { categoria: row.categoria, ingresos: 0, unidades: 0 };
      c.ingresos += Number(row.ingreso || 0); c.unidades += Number(row.cantidad || 0);
      byCategory.set(row.categoria, c);
      byChannel.set(row.canal, (byChannel.get(row.canal) || 0) + Number(row.ingreso || 0));
      const pay = byPayment.get(row.pago) || { label: row.pago, ingresos: 0, unidades: 0, orders: new Set() };
      pay.ingresos += Number(row.ingreso || 0); pay.unidades += Number(row.cantidad || 0); pay.orders.add(row.pedido);
      byPayment.set(row.pago, pay);
    }
    const productsSorted = [...byProduct.values()].map((p) => ({ ...p, ingresos: round2(p.ingresos), utilidad: round2(p.utilidad) })).sort((a, b) => b.ingresos - a.ingresos);
    const categoriesSorted = [...byCategory.values()].map((c) => ({ ...c, ingresos: round2(c.ingresos) })).sort((a, b) => b.ingresos - a.ingresos);

    const goalMap = new Map(src.goals.map((g) => [`${g.year}-${String(g.month).padStart(2, '0')}`, g]));
    let totalReal = 0;
    let totalGoal = 0;
    const goals = [];
    for (const month of monthly) {
      const goal = goalMap.get(month.key);
      if (!goal) continue;
      totalReal += month.ingresos;
      totalGoal += Number(goal.meta_ingresos || 0);
      goals.push({
        key: month.key, label: month.label,
        real: month.ingresos, meta: round2(goal.meta_ingresos),
        cumplimiento: goal.meta_ingresos > 0 ? month.ingresos / goal.meta_ingresos * 100 : null,
        pedidos: month.pedidos, meta_pedidos: round2(goal.meta_pedidos),
      });
    }

    const table = rows.slice().sort((a, b) => b.fecha.localeCompare(a.fecha) || b.pedido.localeCompare(a.pedido));
    const get = (name) => params.get(name) || null;

    return {
      meta: { ...src.meta, cache_hit: true },
      filters_applied: {
        fecha_desde: get('fecha_desde'), fecha_hasta: get('fecha_hasta'),
        producto: get('producto'), categoria: get('categoria'), canal: get('canal'),
      },
      options: src.options,
      kpis: {
        ingresos, pedidos, ticket_promedio: pedidos ? round2(ingresos / pedidos) : 0,
        unidades, costos, utilidad, margen, crecimiento: growth,
        cumplimiento_metas: totalGoal > 0 ? totalReal / totalGoal * 100 : null,
        filas_filtradas: rows.length,
      },
      charts: {
        tendencia: monthly.map((m) => ({ key: m.key, label: m.label, ingresos: m.ingresos, pedidos: m.pedidos })),
        ingresos_costos: monthly.map((m) => ({ key: m.key, label: m.label, ingresos: m.ingresos, costos: m.costos, utilidad: m.utilidad })),
        categorias: categoriesSorted.map((c) => ({ label: c.categoria, value: c.ingresos, unidades: c.unidades })),
        productos: productsSorted.slice(0, 8).map((p) => ({ label: p.producto, value: p.ingresos, unidades: p.unidades, utilidad: p.utilidad })),
        metas: goals,
        canales: [...byChannel.entries()].map(([label, value]) => ({ label, value: round2(value) })).sort((a, b) => b.value - a.value),
        metodos_pago: [...byPayment.values()].map((p) => ({ label: p.label, value: round2(p.ingresos), pedidos: p.orders.size, unidades: p.unidades })).sort((a, b) => b.value - a.value),
      },
      table,
    };
  }

  const BIG = [
    'venta mas grande', 'venta mayor', 'mayor venta', 'pedido mas grande', 'pedido mayor',
    'pedido mas caro', 'pedido mas alto', 'monto mas alto', 'monto mas grande', 'importe mas alto',
    'venta mas alta', 'venta mas cara', 'facturo mas', 'venta facturo mas', 'ingreso mas alto',
    'venta maxima', 'record de ventas',
  ].map(norm);
  const SMALL = [
    'venta mas pequena', 'venta mas baja', 'menor venta', 'venta minima',
    'pedido mas pequeno', 'monto mas bajo', 'ingreso mas bajo',
  ].map(norm);
  const TOP = [
    'top ventas', 'top de ventas', 'top pedidos', 'top de pedidos', 'top 5 ventas',
    'top 5 de ventas', 'las 5 ventas mas altas', 'cinco ventas mas grandes',
    'mejores ventas', 'ventas mas altas', 'ventas mas grandes',
  ].map(norm);
  const METRICS = {
    ingresos: ['ingresos', 'ventas totales', 'total vendido', 'facturacion', 'cuanto vendimos', 'cuanto vendio', 'cuanto se vendio', 'vendimos', 'vendio'],
    pedidos: ['pedidos', 'ordenes', 'numero de pedidos', 'cuantas ventas', 'cuantos pedidos'],
    ticket: ['ticket', 'promedio por pedido', 'venta promedio', 'valor promedio'],
    unidades: ['unidades', 'cantidad vendida', 'piezas', 'articulos vendidos'],
    costos: ['costos', 'costo total', 'cuanto costo', 'gasto'],
    utilidad: ['utilidad', 'ganancia', 'utilidad bruta', 'margen en pesos'],
    margen: ['margen', 'rentabilidad', 'porcentaje de ganancia'],
    crecimiento: ['crecimiento', 'variacion', 'mes anterior', 'evolucion', 'subio', 'bajo'],
    metas: ['meta', 'metas', 'cumplimiento', 'objetivo', 'planificado', 'plan'],
  };

  function monthsOf(message) {
    const text = plain(message);
    const keys = [];
    for (const name of Object.keys(MONTHS).sort((a, b) => b.length - a.length)) {
      if (new RegExp('\\b' + name + '\\b').test(text)) {
        const key = '2026-' + MONTHS[name];
        if (!keys.includes(key)) keys.push(key);
      }
    }
    for (const match of text.match(/\b(20\d{2})-(0[1-9]|1[0-2])\b/g) || []) {
      if (!keys.includes(match)) keys.push(match);
    }
    return keys;
  }

  function entities(text, options) {
    const groups = [
      ['producto', options.productos], ['categoria', options.categorias],
      ['canal', options.canales], ['metodo de pago', options.metodos_pago],
    ];
    const found = [];
    for (const [kind, values] of groups) {
      const matches = values.filter((value) => norm(value) && text.includes(norm(value)));
      if (matches.length) found.push([kind, matches.reduce((a, b) => norm(b).length > norm(a).length ? b : a)]);
    }
    return found;
  }

  function scopedRows(payload, message) {
    let rows = payload.table.slice();
    const scope = [];
    const months = monthsOf(message);
    if (months.length) {
      rows = rows.filter((row) => months.some((key) => row.fecha.startsWith(key)));
      if (!rows.length) return { rows: [], scope: '', missing: true };
      const labels = Object.fromEntries(payload.charts.tendencia.map((item) => [item.key, item.label]));
      scope.push('en ' + months.map((key) => labels[key] || key).join(', '));
    }
    if (/\b(de|del|en|por|para)\b/.test(plain(message))) {
      const found = entities(norm(message), payload.options);
      if (found.length) {
        const [kind, value] = found.reduce((a, b) => norm(b[1]).length > norm(a[1]).length ? b : a);
        const field = { producto: 'producto', categoria: 'categoria', canal: 'canal', 'metodo de pago': 'pago' }[kind];
        const filtered = rows.filter((row) => norm(row[field]) === norm(value));
        if (!filtered.length) return { rows: [], scope: '', missing: true };
        rows = filtered;
        scope.push(`de ${kind} «${value}»`);
      }
    }
    return { rows, scope: scope.length ? ' ' + scope.join(' ') : '', missing: false };
  }

  function extreme(biggest) {
    return (message, payload) => {
      const scoped = scopedRows(payload, message);
      if (scoped.missing || !scoped.rows.length) return { respuesta: SIN_DATO, fuentes: [] };
      const row = biggest
        ? scoped.rows.reduce((a, b) => Number(b.ingreso) > Number(a.ingreso) ? b : a)
        : scoped.rows.reduce((a, b) => Number(b.ingreso) < Number(a.ingreso) ? b : a);
      return {
        respuesta: `La venta ${biggest ? 'mayor' : 'menor'} del periodo${scoped.scope}: ${money(row.ingreso)} (pedido ${row.pedido} del ${row.fecha}, ${row.producto}, categoría ${row.categoria}, ${number(row.cantidad)} unidades, canal ${row.canal}, pago ${row.pago}; hoja Ventas).`,
        fuentes: ['Ventas: Pedido, Fecha, Producto, Categoría y Cantidad', 'Ventas: Cantidad × Precio unitario'],
      };
    };
  }

  function topSalesAnswer(message, payload) {
    const scoped = scopedRows(payload, message);
    if (scoped.missing || !scoped.rows.length) return { respuesta: SIN_DATO, fuentes: [] };
    const top = scoped.rows.slice().sort((a, b) => Number(b.ingreso) - Number(a.ingreso)).slice(0, 5);
    const lines = [`Top ${top.length} ventas por ingreso${scoped.scope} (hoja Ventas):`];
    top.forEach((row, index) => lines.push(`${index + 1}. ${row.pedido} · ${row.fecha} · ${money(row.ingreso)} — ${row.producto}`));
    return { respuesta: lines.join('\n'), fuentes: ['Ventas: Pedido, Fecha, Producto y Cantidad × Precio unitario'] };
  }

  function monthAnswer(months, payload, message) {
    const trend = Object.fromEntries(payload.charts.tendencia.map((item) => [item.key, item]));
    const goals = Object.fromEntries(payload.charts.metas.map((item) => [item.key, item]));
    const text = norm(message);
    const lines = [];
    const sources = new Set(['Ventas: Fecha, Cantidad × Precio unitario']);
    let found = 0;
    for (const key of months) {
      const item = trend[key];
      if (!item) { lines.push(`No hay registros de ${key} con los filtros actuales.`); continue; }
      found += 1;
      const costs = sum(payload.table.filter((row) => row.fecha.startsWith(key)), 'costo');
      let line = `${item.label}: ingresos ${money(item.ingresos)}, ${number(item.pedidos)} pedidos, utilidad ${money(item.ingresos - costs)}.`;
      const goal = goals[key];
      if (goal) {
        line += ` Meta ${money(goal.meta)}: cumplimiento ${pct(goal.cumplimiento)}.`;
        sources.add('Metas: Meta de ingresos');
      }
      lines.push(line);
    }
    if (found >= 2 && hasAny(text, ['comparar', 'comparado', 'diferencia', 'versus'])) {
      const values = months.filter((key) => trend[key]).map((key) => trend[key].ingresos);
      if (values.length >= 2 && Math.min(...values) > 0) {
        lines.push(`El mes mayor supera al menor por ${money(Math.max(...values) - Math.min(...values))} (${pct((Math.max(...values) / Math.min(...values) - 1) * 100)}).`);
      }
    }
    if (!found) lines.push(SIN_DATO);
    return { respuesta: lines.join('\n'), fuentes: [...sources].sort() };
  }

  function aggregate(rows, field, value) {
    const filtered = rows.filter((row) => norm(row[field]) === norm(value));
    if (!filtered.length) return null;
    const ingresos = sum(filtered, 'ingreso');
    const costos = sum(filtered, 'costo');
    const utilidad = round2(ingresos - costos);
    return {
      ingresos, pedidos: new Set(filtered.map((row) => row.pedido)).size,
      unidades: filtered.reduce((acc, row) => acc + Number(row.cantidad || 0), 0),
      utilidad, margen: ingresos > 0 ? utilidad / ingresos * 100 : 0,
    };
  }

  function entityAnswer(kind, value, payload) {
    const field = { producto: 'producto', categoria: 'categoria', canal: 'canal', 'metodo de pago': 'pago' }[kind];
    const result = aggregate(payload.table, field, value);
    if (!result) return { respuesta: SIN_DATO, fuentes: [] };
    const label = { producto: 'Producto', categoria: 'Categoría', canal: 'Canal', 'metodo de pago': 'Método de pago' }[kind];
    return {
      respuesta: `${label} «${value}»: ingresos ${money(result.ingresos)}, ${number(result.pedidos)} pedidos, ${number(result.unidades)} unidades, utilidad ${money(result.utilidad)} y margen ${pct(result.margen)}.`,
      fuentes: ['Ventas: ' + label, 'Ventas: Cantidad × Precio unitario y Costo unitario'],
    };
  }

  function listingAnswer(kind, payload) {
    const key = { categoria: 'categorias', canal: 'canales', 'metodo de pago': 'metodos_pago', producto: 'productos' }[kind];
    const items = payload.charts[key] || [];
    if (!items.length) return { respuesta: SIN_DATO, fuentes: [] };
    const title = { categoria: 'categorías', canal: 'canales', 'metodo de pago': 'métodos de pago', producto: 'productos' }[kind];
    const top = items.slice(0, 5);
    return {
      respuesta: [`Top ${top.length} ${title} por ingresos (hoja Ventas):`, ...top.map((item, index) => `${index + 1}. ${item.label} — ${money(item.value)}`)].join('\n'),
      fuentes: ['Ventas: Cantidad × Precio unitario'],
    };
  }

  function chatAnswer(payload, message) {
    const text = norm(message);
    if (!text) return { respuesta: 'Escribe una pregunta sobre los datos del dashboard y respondo solo con el Excel.', fuentes: [], verificado: true };
    const k = payload.kpis;
    const trend = payload.charts.tendencia;

    if (hasAny(text, ['hola', 'buenos dias', 'buenas tardes', 'que tal']) && text.length < 24) {
      return { respuesta: 'Hola, soy A&M Copiloto. Respondo únicamente con los datos del Excel cargado. Pregúntame por ventas, pedidos, metas, categorías, productos o canales.', fuentes: [], verificado: true };
    }
    if (hasAny(text, ['ayuda', 'que puedes hacer', 'como funciona', 'que sabes'])) {
      return { respuesta: 'Puedo responder sobre ingresos, pedidos, ticket promedio, unidades, costos, utilidad, margen, crecimiento mensual, metas, categorías, productos, canales, métodos de pago y la venta más grande o más pequeña. Si el dato no está en el Excel, te lo digo sin inventarlo.', fuentes: [], verificado: true };
    }
    if (hasAny(text, BIG)) return { ...extreme(true)(message, payload), verificado: true };
    if (hasAny(text, SMALL)) return { ...extreme(false)(message, payload), verificado: true };
    if (hasAny(text, TOP)) return { ...topSalesAnswer(message, payload), verificado: true };

    const months = monthsOf(message);
    if (months.length) return { ...monthAnswer(months, payload, message), verificado: true };

    const found = entities(text, payload.options);
    if (found.length) {
      const [kind, value] = found.reduce((a, b) => norm(b[1]).length > norm(a[1]).length ? b : a);
      return { ...entityAnswer(kind, value, payload), verificado: true };
    }

    if (hasAny(text, ['top', 'mas vendido', 'destacado', 'ranking', 'mejor vendido'])) return { ...listingAnswer('producto', payload), verificado: true };
    if (has(text, 'categoria') || has(text, 'categorias')) return { ...listingAnswer('categoria', payload), verificado: true };
    if (has(text, 'canal') || has(text, 'canales')) return { ...listingAnswer('canal', payload), verificado: true };
    if (hasAny(text, ['metodo de pago', 'metodos de pago', 'forma de pago'])) return { ...listingAnswer('metodo de pago', payload), verificado: true };

    if (hasAny(text, ['cual mes', 'que mes', 'mejor mes', 'peor mes'])) {
      if (!trend.length) return { respuesta: SIN_DATO, fuentes: [], verificado: true };
      const best = trend.reduce((a, b) => Number(b.ingresos) > Number(a.ingresos) ? b : a);
      const worst = trend.reduce((a, b) => Number(b.ingresos) < Number(a.ingresos) ? b : a);
      return {
        respuesta: `El mes con más ingresos fue ${best.label} con ${money(best.ingresos)}. El de menos fue ${worst.label} con ${money(worst.ingresos)}. Diferencia: ${money(best.ingresos - worst.ingresos)} (hoja Ventas).`,
        fuentes: ['Ventas: Fecha, Cantidad × Precio unitario'], verificado: true,
      };
    }

    if (hasAny(text, ['cliente', 'empleado', 'sueldo', 'impuesto', 'predice', 'pronostico', 'proyeccion', 'futuro', 'invertir', 'opinion', 'recomiendas'])) {
      return { respuesta: SIN_DATO, fuentes: [], verificado: true };
    }

    for (const [metric, words] of Object.entries(METRICS)) {
      if (!hasAny(text, words)) continue;
      if (metric === 'ingresos') return { respuesta: `Ingresos totales: ${money(k.ingresos)} en el periodo filtrado (${number(k.filas_filtradas)} registros de la hoja Ventas).`, fuentes: ['Ventas: Cantidad × Precio unitario'], verificado: true };
      if (metric === 'pedidos') return { respuesta: `Pedidos: ${number(k.pedidos)} en el periodo filtrado (hoja Ventas).`, fuentes: ['Ventas: Pedido'], verificado: true };
      if (metric === 'ticket') return { respuesta: `Ticket promedio: ${money(k.ticket_promedio)} por pedido (ingresos ÷ pedidos, hoja Ventas).`, fuentes: ['Ventas: Cantidad × Precio unitario', 'Ventas: Pedido'], verificado: true };
      if (metric === 'unidades') return { respuesta: `Unidades vendidas: ${number(k.unidades)} (hoja Ventas, Cantidad).`, fuentes: ['Ventas: Cantidad'], verificado: true };
      if (metric === 'costos') return { respuesta: `Costos totales: ${money(k.costos)} (Cantidad × Costo unitario, hoja Ventas).`, fuentes: ['Ventas: Cantidad × Costo unitario'], verificado: true };
      if (metric === 'utilidad') return { respuesta: `Utilidad bruta: ${money(k.utilidad)} (ingresos menos costos, hoja Ventas).`, fuentes: ['Ventas: Cantidad × Precio unitario y Costo unitario'], verificado: true };
      if (metric === 'margen') return { respuesta: `Margen bruto: ${pct(k.margen)} (utilidad ÷ ingresos, hoja Ventas).`, fuentes: ['Ventas: Cantidad × Precio unitario y Costo unitario'], verificado: true };
      if (metric === 'crecimiento') {
        if (k.crecimiento === null) return { respuesta: 'No puedo calcular el crecimiento con los filtros actuales: se necesitan al menos dos meses en el periodo (hoja Ventas).', fuentes: [], verificado: true };
        const detail = trend.map((item) => `${item.label}: ${money(item.ingresos)}`).join('; ');
        return { respuesta: `Crecimiento del último mes frente al anterior: ${pct(k.crecimiento)}. ${detail}.`, fuentes: ['Ventas: Fecha, Cantidad × Precio unitario'], verificado: true };
      }
      if (metric === 'metas') {
        const goals = payload.charts.metas;
        if (!goals.length || k.cumplimiento_metas === null) return { respuesta: 'No hay metas mensuales en el Excel para el periodo filtrado.', fuentes: [], verificado: true };
        const detail = goals.map((item) => `${item.label}: ${pct(item.cumplimiento)}`).join('; ');
        return { respuesta: `Cumplimiento global de metas: ${pct(k.cumplimiento_metas)}. Por mes — ${detail}.`, fuentes: ['Metas: Meta de ingresos', 'Ventas: Cantidad × Precio unitario'], verificado: true };
      }
    }
    return { respuesta: SIN_DATO, fuentes: [], verificado: true };
  }

  function csvBlob(rows) {
    const cols = ['fecha', 'pedido', 'producto', 'categoria', 'cantidad', 'precio', 'ingreso', 'costo', 'utilidad', 'canal', 'pago'];
    const quote = (value) => '"' + String(value ?? '').replace(/"/g, '""') + '"';
    const text = '\uFEFF' + [cols.join(','), ...rows.map((row) => cols.map((col) => quote(row[col])).join(','))].join('\r\n');
    return new Blob([text], { type: 'text/csv;charset=utf-8' });
  }

  function exportBlob(format, rows) {
    if (format === 'csv' || typeof window.XLSX === 'undefined') return csvBlob(rows);
    const sheet = window.XLSX.utils.json_to_sheet(rows);
    const book = window.XLSX.utils.book_new();
    window.XLSX.utils.book_append_sheet(book, sheet, 'Ventas');
    const buffer = window.XLSX.write(book, { bookType: 'xlsx', type: 'array' });
    return new Blob([buffer], { type: 'application/vnd.openxmlformats-officedocument.spreadsheetml.sheet' });
  }

  const jsonResponse = (value) => new Response(JSON.stringify(value), { status: 200, headers: { 'Content-Type': 'application/json' } });

  window.fetch = async function (input, init) {
    const url = String(input);
    if (url.startsWith('/api/refresh')) return jsonResponse({ ok: true });
    if (url.startsWith('/api/dashboard')) {
      const params = new URLSearchParams(url.includes('?') ? url.split('?').slice(1).join('?') : '');
      return jsonResponse(buildDashboard(await getSource(), params));
    }
    if (url.startsWith('/api/chat')) {
      const body = JSON.parse((init && init.body) || '{}');
      const params = new URLSearchParams();
      for (const key of ['fecha_desde', 'fecha_hasta', 'producto', 'categoria', 'canal']) {
        if (body[key]) params.set(key, body[key]);
      }
      const payload = buildDashboard(await getSource(), params);
      return jsonResponse(chatAnswer(payload, body.mensaje || ''));
    }
    if (url.startsWith('/api/export')) {
      const params = new URLSearchParams(url.includes('?') ? url.split('?').slice(1).join('?') : '');
      const rows = filteredRows(await getSource(), params);
      const format = params.get('formato') || 'csv';
      return new Response(exportBlob(format, rows), { status: 200 });
    }
    return realFetch(input, init);
  };
})();
