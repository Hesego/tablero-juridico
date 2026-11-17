/* =====================================================================
   MOTOR DE CÁLCULO – Tablero jurídico HESEGO INGENIERÍA S.A.S.
   Puro JavaScript, sin dependencias. Lo usa index.html en el navegador y
   pruebas/motor.test.mjs en Node. Recibe las hojas de Excel ya convertidas
   a matrices (filas × columnas, valores crudos) y devuelve el modelo.
   ===================================================================== */
(function (global) {
  'use strict';

  // ---------------------------------------------------------------- utilidades
  const EPOCA_EXCEL = 25569; // serial de Excel del 01/01/1970
  const norm = (s) => s == null ? '' : String(s).normalize('NFD').replace(/[̀-ͯ]/g, '')
    .replace(/\s+/g, ' ').trim().toUpperCase();
  const esNum = (v) => typeof v === 'number' && isFinite(v);
  const vacio = (v) => v == null || (typeof v === 'string' && v.trim() === '');

  /** Número desde celda (acepta "30 %", "1.234.567", "24,5"). Para parámetros, no para FO-GH-47. */
  function num(v) {
    if (esNum(v)) return v;
    if (typeof v !== 'string') return null;
    let t = v.trim().replace(/[$\s]/g, '');
    if (!t) return null;
    const pct = t.endsWith('%');
    if (pct) t = t.slice(0, -1);
    if (/^-?\d{1,3}(\.\d{3})+(,\d+)?$/.test(t)) t = t.replace(/\./g, '').replace(',', '.');
    else if (/^-?\d+,\d+$/.test(t)) t = t.replace(',', '.');
    else if (!/^-?\d+(\.\d+)?$/.test(t)) return null;
    const n = parseFloat(t);
    return pct ? n / 100 : n;
  }

  // Fechas como número entero de días desde 01/01/1970 (UTC): aritmética simple y sin husos.
  const dia = (y, m, d) => Math.floor(Date.UTC(y, m - 1, d) / 864e5);
  const partes = (d) => { const t = new Date(d * 864e5); return { y: t.getUTCFullYear(), m: t.getUTCMonth() + 1, d: t.getUTCDate() }; };
  const dow = (d) => (((d + 4) % 7) + 7) % 7; // 0 = domingo
  const hoyDia = () => { const t = new Date(); return dia(t.getFullYear(), t.getMonth() + 1, t.getDate()); };
  const finMes = (y, m) => dia(y, m + 1, 1) - 1;

  /** Fecha desde serial de Excel, Date o texto dd/mm/aaaa · aaaa-mm-dd. */
  function fecha(v) {
    if (esNum(v)) return v > 1000 && v < 2958465 ? Math.floor(v) - EPOCA_EXCEL : null;
    if (v instanceof Date && !isNaN(v)) return dia(v.getFullYear(), v.getMonth() + 1, v.getDate());
    if (typeof v === 'string') {
      const s = v.trim();
      let m = s.match(/^(\d{1,2})[/\-.](\d{1,2})[/\-.](\d{2,4})\b/);
      if (m) { let y = +m[3]; if (y < 100) y += 2000; if (+m[2] >= 1 && +m[2] <= 12) return dia(y, +m[2], +m[1]); }
      m = s.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
      if (m) return dia(+m[1], +m[2], +m[3]);
    }
    return null;
  }
  const fmtFecha = (d) => { if (d == null) return ''; const p = partes(d); return `${String(p.d).padStart(2, '0')}/${String(p.m).padStart(2, '0')}/${p.y}`; };

  // ------------------------------------------------------- festivos de Colombia
  function pascua(y) { // algoritmo anónimo gregoriano
    const a = y % 19, b = Math.floor(y / 100), c = y % 100, d = Math.floor(b / 4), e = b % 4;
    const f = Math.floor((b + 8) / 25), g = Math.floor((b - f + 1) / 3), h = (19 * a + b - d - g + 15) % 30;
    const i = Math.floor(c / 4), k = c % 4, l = (32 + 2 * e + 2 * i - h - k) % 7, m = Math.floor((a + 11 * h + 22 * l) / 451);
    const mes = Math.floor((h + l - 7 * m + 114) / 31), dd = ((h + l - 7 * m + 114) % 31) + 1;
    return dia(y, mes, dd);
  }
  const cacheFestivos = {};
  /** Festivos: fijos, trasladables al lunes (Ley 51 de 1983) y dependientes de la Pascua. */
  function festivos(y) {
    if (cacheFestivos[y]) return cacheFestivos[y];
    const lunes = (d) => d + ((8 - dow(d)) % 7);
    const p = pascua(y);
    const lista = [
      dia(y, 1, 1), dia(y, 5, 1), dia(y, 7, 20), dia(y, 8, 7), dia(y, 12, 8), dia(y, 12, 25),
      lunes(dia(y, 1, 6)), lunes(dia(y, 3, 19)), lunes(dia(y, 6, 29)), lunes(dia(y, 8, 15)),
      lunes(dia(y, 10, 12)), lunes(dia(y, 11, 1)), lunes(dia(y, 11, 11)),
      p - 3, p - 2, p + 43, p + 64, p + 71
    ];
    return (cacheFestivos[y] = new Set(lista));
  }
  const esHabil = (d) => { const w = dow(d); return w !== 0 && w !== 6 && !festivos(partes(d).y).has(d); };
  function networkdays(a, b) { if (b < a) return -networkdays(b, a); let n = 0; for (let d = a; d <= b; d++) if (esHabil(d)) n++; return n; }
  /** días hábiles = NETWORKDAYS(recibido, respuesta) − 1 (mínimo 0 si se responde en fin de semana o festivo). */
  const diasHabiles = (a, b) => b < a ? networkdays(a, b) - 1 : Math.max(0, networkdays(a, b) - 1);
  function workday(d, n) { let k = 0; while (k < n) { d++; if (esHabil(d)) k++; } return d; }

  // ---------------------------------------------------------- configuración base
  const CONFIG_BASE = {
    ans: {
      'DERECHO DE PETICION': 5, 'TUTELAS': 2, 'GESTION JURIDICA': 10, 'REVISION DOCUMENTAL Y CONCEPTO JURIDICO': 7,
      'EMBARGOS': 5, 'SINIESTROS': 8, 'DEMANDAS': 15, 'RENUNCIA MOTIVADA': 2, 'PROCESOS DISCIPLINARIOS': 5
    },
    terminoLegalPorDefecto: { 'DERECHO DE PETICION': 15, 'TUTELAS': 2 },
    grupos: {
      A: ['TUTELAS', 'DERECHO DE PETICION', 'DEMANDAS', 'EMBARGOS', 'GESTION JURIDICA', 'SINIESTROS', 'RENUNCIA MOTIVADA'],
      B: ['REVISION DOCUMENTAL Y CONCEPTO JURIDICO', 'PROCESOS DISCIPLINARIOS']
    },
    umbralAlertaPP: 5,          // hasta 5 p.p. por debajo de la meta = en alerta
    modoParcial: true,          // estratégico repartiendo peso entre tácticos con medición (se marca "parcial")
    usarHojaANS: true,          // si la FO-GH-47 trae la hoja ANS, sus valores reemplazan los de arriba
    diasProximoVencer: 2
  };

  // -------------------------------------------------------------- lector FO-GH-47
  const COLUMNAS_GH47 = {
    id: ['NO.', 'NO', 'N°'], clas: ['CLASIFICACION'], nombre: ['NOMBRE'], asunto: ['ASUNTO'],
    sede: ['PROCESO QUE REMITE'], recibido: ['FECHA DE RECIBIDO'], tipoTermino: ['TIPO DE TERMINO'],
    limite: ['FECHA LIMITE'], respuesta: ['FECHA EN QUE DI RESPUESTA'], estado: ['ESTADO'],
    destinatario: ['DESTINATARIO'], riesgo: ['RIESGO JURIDICO'], estimado: ['VALOR ESTIMADO'],
    pendiente: ['PENDIENTE DE REVISION'], observaciones: ['OBSERVACIONES'], posicion: ['POSICION DE HESEGO'],
    pretendido: ['VALOR PRETENDIDO'], solidez: ['SOLIDEZ'], pruebas: ['PRUEBAS EN CONTRA'],
    jurisprudencia: ['JURISPRUDENCIA'], conducta: ['CONDUCTA'], final: ['VALOR FINAL']
  };

  function ubicarColumnas(encabezados, calidad) {
    const h = encabezados.map(norm);
    const col = {};
    for (const [clave, opciones] of Object.entries(COLUMNAS_GH47)) {
      let idx = -1;
      for (const o of opciones) { idx = h.findIndex((x) => x === o); if (idx >= 0) break; }
      if (idx < 0) for (const o of opciones) { idx = h.findIndex((x) => x.startsWith(o)); if (idx >= 0) break; }
      col[clave] = idx;
      if (idx < 0) calidad.push({ tipo: 'Columna no encontrada', fila: 1, id: '', detalle: `FO-GH-47: no se encontró la columna «${opciones[0]}».` });
    }
    return col;
  }

  function leerCasos(filas, config, calidad) {
    if (!filas || !filas.length) return [];
    const col = ubicarColumnas(filas[0], calidad);
    const g = (fila, c) => (col[c] >= 0 ? fila[col[c]] : null);
    const todosGrupos = new Set([...config.grupos.A, ...config.grupos.B]);
    const casos = [];
    for (let i = 1; i < filas.length; i++) {
      const f = filas[i] || [];
      const clas = norm(g(f, 'clas'));
      if (!clas) continue; // filas sin clasificación se ignoran
      const id = vacio(g(f, 'id')) ? `F${i + 1}` : String(g(f, 'id')).trim();
      const reg = (tipo, detalle) => calidad.push({ tipo, fila: i + 1, id, detalle });
      const c = {
        id, fila: i + 1, clas, clasOrig: String(g(f, 'clas')).trim(),
        nombre: String(g(f, 'nombre') ?? '').trim(), asunto: String(g(f, 'asunto') ?? '').trim(),
        sede: norm(g(f, 'sede')) || 'SIN SEDE',
        recibido: fecha(g(f, 'recibido')), limite: fecha(g(f, 'limite')), respuesta: fecha(g(f, 'respuesta')),
        estado: norm(g(f, 'estado')), riesgo: norm(g(f, 'riesgo')), destinatario: String(g(f, 'destinatario') ?? '').trim(),
        pendiente: norm(g(f, 'pendiente')) === 'SI', posicion: norm(g(f, 'posicion')) === 'A FAVOR' ? 'A FAVOR' : 'EN CONTRA',
        posicionRegistrada: !vacio(g(f, 'posicion'))
      };
      const est = g(f, 'estimado'), pre = g(f, 'pretendido'), fin = g(f, 'final');
      c.estimado = esNum(est) ? est : null; c.pretendido = esNum(pre) ? pre : null; c.valorFinal = esNum(fin) ? fin : null;
      c.factores = ['solidez', 'pruebas', 'jurisprudencia', 'conducta'].map((k) => (esNum(g(f, k)) ? g(f, k) : null));

      // --- calidad de datos
      if (!todosGrupos.has(clas)) reg('Clasificación no reconocida', `«${c.clasOrig}» no pertenece a ningún táctico.`);
      if (c.recibido == null) reg('Fecha de recibido inválida', `Valor: «${g(f, 'recibido') ?? 'vacío'}».`);
      if (!vacio(g(f, 'respuesta')) && c.respuesta == null) reg('Fecha de respuesta inválida', `Valor: «${g(f, 'respuesta')}».`);
      if (c.recibido != null && c.respuesta != null && c.respuesta < c.recibido)
        reg('Respuesta anterior al recibido', `Recibido ${fmtFecha(c.recibido)}, respuesta ${fmtFecha(c.respuesta)}.`);
      if (!vacio(est) && !esNum(est)) reg('Valor estimado en texto', `Se trata como no numérico: «${String(est).slice(0, 80)}…».`);
      if (!vacio(pre) && !esNum(pre)) reg('Valor pretendido en texto', `«${String(pre).slice(0, 80)}».`);
      if (c.estado === 'CERRADO' && c.respuesta == null) reg('Cerrado sin fecha de respuesta', 'Se considera abierto hasta registrar la fecha.');
      if (c.estado !== 'CERRADO' && c.respuesta != null) reg('Respuesta registrada sin estado CERRADO', `Estado «${c.estado || 'vacío'}» con respuesta ${fmtFecha(c.respuesta)}; se mide como respondido en esa fecha. Verifique el estado.`);
      if (!['BAJO', 'MEDIO', 'ALTO'].includes(c.riesgo)) reg('Riesgo jurídico vacío o inválido', `Valor: «${g(f, 'riesgo') ?? ''}».`);

      // --- días hábiles y ANS
      c.ans = config.ans[clas] ?? null;
      c.dias = c.recibido != null && c.respuesta != null ? diasHabiles(c.recibido, c.respuesta) : null;
      c.venceANS = c.recibido != null && c.ans != null ? workday(c.recibido, c.ans) : null;

      // --- probabilidad (metodología FO-GH-48)
      const nFact = c.factores.filter((x) => x != null).length;
      c.pct = nFact === 4 ? c.factores.reduce((a, b) => a + b, 0) / 20 : null;
      if (c.estimado == null && c.pretendido == null) c.prob = 'POR VALORAR';
      else if (Math.max(c.estimado || 0, c.pretendido || 0) === 0) c.prob = 'SIN CONTINGENCIA ECONOMICA';
      else if (nFact < 4) c.prob = 'POR VALORAR';
      else c.prob = c.pct < 0.5 ? 'REMOTO' : c.pct < 0.7 ? 'POSIBLE' : 'PROBABLE';
      c.tratamiento = { PROBABLE: 'PROVISIONAR', POSIBLE: 'NOTA ADVERTENCIA', REMOTO: 'SIN RIESGO', 'POR VALORAR': 'VALIDAR' }[c.prob] || 'NO APLICA';
      const base = (c.estimado || 0) > 0 ? c.estimado : (c.pretendido || 0);
      c.reserva = c.posicion === 'EN CONTRA' && (c.prob === 'PROBABLE' || c.prob === 'POSIBLE') ? base * c.pct : 0;
      c.base = base;
      casos.push(c);
    }
    return casos;
  }

  /** Abierto al corte: recibido hasta esa fecha y sin respuesta a esa fecha (permite reconstruir cualquier mes pasado). */
  const abiertoAl = (c, corte) => c.recibido != null && c.recibido <= corte && !(c.respuesta != null && c.respuesta <= corte);
  /** Cerrado dentro del rango [ini, fin]: lo define la FECHA EN QUE DI RESPUESTA (el ESTADO solo se valida en calidad). */
  const cerradoEn = (c, ini, fin) => c.respuesta != null && c.respuesta >= ini && c.respuesta <= fin;

  // ------------------------------------------------------------- lector FO-IG-136
  function leerCascada(filas, calidad) {
    const out = { estrategico: null, tacticos: [] };
    if (!filas) return out;
    const L = (letra) => letra.charCodeAt(0) - 65;
    const v = (f, letra) => (f ? f[L(letra)] : null);
    let actual = null, vacias = 0;
    for (let i = 7; i < filas.length; i++) { // fila 8 en adelante
      const f = filas[i] || [];
      const g = v(f, 'G'), l = v(f, 'L');
      if (vacio(g) && vacio(l)) { if (++vacias >= 3 && out.tacticos.length) break; continue; }
      if (!vacio(g) && /^EJEMPLO/i.test(String(g).trim())) break;
      vacias = 0;
      if (!out.estrategico && !vacio(v(f, 'B')))
        out.estrategico = { nombre: String(v(f, 'B')).trim(), descripcion: String(v(f, 'C') ?? '').trim(), meta: num(v(f, 'E')) ?? 1 };
      if (!vacio(g)) {
        actual = { nombre: String(g).trim(), descripcion: String(v(f, 'H') ?? '').trim(), peso: num(v(f, 'F')), meta: num(v(f, 'J')) ?? 1, operativos: [], fila: i + 1 };
        out.tacticos.push(actual);
        if (actual.peso == null) calidad.push({ tipo: 'FO-IG-136 sin peso', fila: i + 1, id: '', detalle: `El táctico «${actual.nombre}» no tiene peso.` });
      }
      if (!vacio(l) && actual) {
        const sentido = norm(v(f, 'P'));
        actual.operativos.push({
          nombre: String(l).trim(), peso: num(v(f, 'M')) ?? 0, meta: num(v(f, 'O')),
          sentido: sentido.startsWith('MEN') ? 'Menor' : sentido.startsWith('IGU') ? 'Igual' : sentido.startsWith('MAY') ? 'Mayor' : null,
          objetivo: String(v(f, 'R') ?? '').trim(), definicion: String(v(f, 'S') ?? '').trim(), formula: String(v(f, 'T') ?? '').trim(),
          frecuencia: String(v(f, 'U') ?? '').trim(), origen: String(v(f, 'V') ?? '').trim(), responsable: String(v(f, 'W') ?? '').trim(),
          fila: i + 1
        });
      }
    }
    if (!out.estrategico) out.estrategico = { nombre: 'Índice de cumplimiento ponderado de la gestión jurídica', descripcion: '', meta: 1 };
    const sumaT = out.tacticos.reduce((a, t) => a + (t.peso || 0), 0);
    if (out.tacticos.length && Math.abs(sumaT - 1) > 0.001)
      calidad.push({ tipo: 'FO-IG-136 pesos', fila: '', id: '', detalle: `Los pesos tácticos suman ${(sumaT * 100).toFixed(1)} % (deben sumar 100 %).` });
    return out;
  }

  /** Asocia cada operativo de la FO-IG-136 con su regla de cálculo. */
  function asignarCalculo(op, tactico) {
    const n = norm(op.nombre), t = norm(tactico.nombre);
    const grupoTactico = t.includes('CONSTITUCIONAL') ? 'A' : (t.includes('REVISION') || t.includes('DISCIPLIN')) ? 'B' : null;
    let clases = grupoTactico;
    // El objeto del indicador es lo que aparece primero en el nombre («Procesos disciplinarios con revisión…»).
    const iD = n.indexOf('DISCIPLIN'), iR = n.indexOf('REVISION');
    if (iD >= 0 && (iR < 0 || iD < iR)) clases = ['PROCESOS DISCIPLINARIOS'];
    else if (iR >= 0 && (iD < 0 || iR < iD)) clases = ['REVISION DOCUMENTAL Y CONCEPTO JURIDICO'];
    if (n.includes('TIEMPO PROMEDIO') && grupoTactico) clases = grupoTactico; // el promedio cubre todo el táctico
    if (n.includes('TERMINOS LEGALES')) return { tipo: 'terminos', unidad: '%' };
    if (n.includes('TIEMPO PROMEDIO')) return { tipo: 'promedio', clases, unidad: 'dias' };
    if (n.includes('ANS')) return { tipo: 'ans', clases, unidad: '%' };
    if (n.includes('COBERTURA')) return { tipo: 'cobertura', unidad: '%' };
    if (n.includes('EFECTIVIDAD')) return { tipo: 'efectividad', unidad: '%' };
    if (n.includes('EXPOSICION')) return { tipo: 'exposicion', unidad: '$' };
    if (n.includes('RECUPERA')) return { tipo: 'recuperacion', unidad: '%' };
    return { tipo: 'registro', unidad: '%' };
  }

  // ----------------------------------------------------------- registro mensual
  function leerRegistro(filas, calidad) {
    const out = [];
    if (!filas || !filas.length) return out;
    const h = filas[0].map(norm);
    const c = (k) => h.findIndex((x) => x.startsWith(k));
    const ia = c('ANO'), im = c('MES'), ii = c('INDICADOR'), inu = c('NUMERADOR'), ide = c('DENOMINADOR'), iv = c('VALOR'), ic = c('COMENTARIO');
    for (let i = 1; i < filas.length; i++) {
      const f = filas[i] || [];
      if (vacio(f[ii])) continue;
      const anio = num(f[ia]), mes = num(f[im]);
      if (!anio || !mes || mes < 1 || mes > 12) { calidad.push({ tipo: 'Registro mensual', fila: i + 1, id: '', detalle: 'Año o mes inválido.' }); continue; }
      const nu = num(f[inu]), de = num(f[ide]);
      let valor = num(f[iv]);
      if (valor == null && nu != null && de) valor = nu / de;
      out.push({ anio, mes, nombre: String(f[ii]).trim(), clave: norm(f[ii]), numerador: nu, denominador: de, valor, comentario: vacio(f[ic]) ? '' : String(f[ic]).trim(), fila: i + 1 });
    }
    return out;
  }

  function leerHojaANS(filas) {
    const out = {};
    if (!filas) return out;
    for (const f of filas.slice(1)) { const k = norm(f?.[0]); if (k && esNum(f[1])) out[k] = f[1]; }
    return out;
  }

  // ------------------------------------------------------------------- modelo
  /**
   * Construye el modelo base a partir de las hojas.
   * hojas = { gh47: filas, ans: filas|null, ig136: filas|null, registro: filas|null }
   */
  function construir(hojas, configUsuario) {
    const config = JSON.parse(JSON.stringify({ ...CONFIG_BASE, ...(configUsuario || {}) }));
    const calidad = [];
    let fuenteANS = 'configuración';
    if (config.usarHojaANS && hojas.ans) {
      const a = leerHojaANS(hojas.ans);
      if (Object.keys(a).length) { Object.assign(config.ans, a); fuenteANS = 'hoja ANS de la FO-GH-47'; }
    }
    const casos = leerCasos(hojas.gh47, config, calidad);
    const cascada = leerCascada(hojas.ig136, calidad);
    if (!hojas.ig136) calidad.push({ tipo: 'Falta FO-IG-136', fila: '', id: '', detalle: 'Sin la matriz de indicadores en cascada no hay pesos ni metas: cargue el archivo FO-IG-136.' });
    for (const t of cascada.tacticos) for (const o of t.operativos) o.calculo = asignarCalculo(o, t);
    const registro = leerRegistro(hojas.registro, calidad);
    const nombresReg = new Set(cascada.tacticos.flatMap((t) => t.operativos.filter((o) => o.calculo.tipo === 'registro').map((o) => norm(o.nombre))));
    for (const r of registro) if (cascada.tacticos.length && !nombresReg.has(r.clave))
      calidad.push({ tipo: 'Registro mensual', fila: r.fila, id: '', detalle: `«${r.nombre}» no coincide con ningún operativo de la FO-IG-136.` });
    // Calidad adicional: abiertos hoy sin valoración
    const hoy = hoyDia();
    for (const c of casos) if (abiertoAl(c, hoy) && c.prob === 'POR VALORAR')
      calidad.push({ tipo: 'Abierto sin riesgo valorado', fila: c.fila, id: c.id, detalle: `${c.clasOrig}: faltan valores o factores de probabilidad.` });
    const anios = [...new Set(casos.flatMap((c) => [c.recibido, c.respuesta]).filter((d) => d != null).map((d) => partes(d).y))].sort();
    return { config, casos, cascada, registro, calidad, fuenteANS, anios, cache: {} };
  }

  /** Rango del periodo. mes = 0 → acumulado del año (hasta hoy si es el año en curso). */
  function rango(anio, mes) {
    if (mes) return { ini: dia(anio, mes, 1), fin: finMes(anio, mes), anio, mes };
    const fin = Math.min(dia(anio, 12, 31), Math.max(dia(anio, 1, 31), hoyDia()));
    return { ini: dia(anio, 1, 1), fin, anio, mes: 0 };
  }
  const anterior = (anio, mes) => (mes ? (mes === 1 ? { anio: anio - 1, mes: 12 } : { anio, mes: mes - 1 }) : { anio: anio - 1, mes: 0 });

  // --------------------------------------------------------------- operativos
  function calcularOperativo(m, op, r) {
    const { casos, config } = m;
    const clasesDe = (cl) => (Array.isArray(cl) ? cl : config.grupos[cl] || []);
    const res = { valor: null, numerador: null, denominador: null, sinCasos: false, casos: [], nota: '' };
    switch (op.calculo.tipo) {
      case 'ans': {
        const cl = clasesDe(op.calculo.clases);
        const lista = casos.filter((c) => cl.includes(c.clas) && cerradoEn(c, r.ini, r.fin));
        lista.forEach((c) => res.casos.push({ c, ok: c.dias != null && c.ans != null && c.dias <= c.ans, detalle: `${c.dias ?? '—'} de ${c.ans ?? '—'} días hábiles` }));
        res.denominador = lista.length; res.numerador = res.casos.filter((x) => x.ok).length;
        if (lista.length) res.valor = res.numerador / res.denominador; else { res.valor = 1; res.sinCasos = true; }
        break;
      }
      case 'terminos': {
        const lista = casos.filter((c) => (c.clas === 'TUTELAS' || c.clas === 'DERECHO DE PETICION') && cerradoEn(c, r.ini, r.fin));
        for (const c of lista) {
          const lim = c.limite ?? (c.recibido != null ? workday(c.recibido, config.terminoLegalPorDefecto[c.clas]) : null);
          res.casos.push({ c, ok: lim != null && c.respuesta <= lim, detalle: `Límite ${fmtFecha(lim)}${c.limite == null ? ' (por defecto)' : ''} · respuesta ${fmtFecha(c.respuesta)}` });
        }
        res.denominador = lista.length; res.numerador = res.casos.filter((x) => x.ok).length;
        if (lista.length) res.valor = res.numerador / res.denominador; else { res.valor = 1; res.sinCasos = true; }
        break;
      }
      case 'promedio': {
        const cl = clasesDe(op.calculo.clases);
        const lista = casos.filter((c) => cl.includes(c.clas) && cerradoEn(c, r.ini, r.fin) && c.dias != null);
        lista.forEach((c) => res.casos.push({ c, ok: op.meta == null || c.dias <= op.meta, detalle: `${c.dias} días hábiles` }));
        res.denominador = lista.length; res.numerador = lista.reduce((a, c) => a + c.dias, 0);
        if (lista.length) res.valor = res.numerador / res.denominador; else { res.valor = 0; res.sinCasos = true; }
        break;
      }
      case 'cobertura': {
        const lista = casos.filter((c) => abiertoAl(c, r.fin));
        lista.forEach((c) => res.casos.push({ c, ok: c.prob !== 'POR VALORAR', detalle: c.prob }));
        res.denominador = lista.length; res.numerador = res.casos.filter((x) => x.ok).length;
        if (lista.length) res.valor = res.numerador / res.denominador; else { res.valor = 1; res.sinCasos = true; }
        break;
      }
      case 'efectividad':
      case 'recuperacion': {
        const pos = op.calculo.tipo === 'efectividad' ? 'EN CONTRA' : 'A FAVOR';
        const ini = dia(r.anio, 1, 1);
        const lista = casos.filter((c) => c.posicion === pos && (pos === 'EN CONTRA' ? c.posicionRegistrada || (c.pretendido || 0) > 0 : true) &&
          cerradoEn(c, ini, r.fin) && (c.pretendido || 0) > 0 && c.valorFinal != null);
        const sp = lista.reduce((a, c) => a + c.pretendido, 0), sf = lista.reduce((a, c) => a + c.valorFinal, 0);
        lista.forEach((c) => res.casos.push({ c, ok: true, detalle: `Pretendido $${c.pretendido.toLocaleString('es-CO')} · final $${c.valorFinal.toLocaleString('es-CO')}` }));
        res.denominador = sp; res.numerador = pos === 'EN CONTRA' ? sp - sf : sf;
        if (lista.length) res.valor = res.numerador / sp;
        else res.nota = `No hay procesos ${pos === 'EN CONTRA' ? 'en contra' : 'a favor'} cerrados con valor final registrado en lo corrido del año.`;
        break;
      }
      case 'exposicion': {
        const c = contingencia(m, r.fin);
        res.valor = c.exposicion; res.casos = c.procesos.map((p) => ({ c: p, ok: true, detalle: `${p.prob} · reserva $${Math.round(p.reserva).toLocaleString('es-CO')}` }));
        res.extra = c;
        break;
      }
      case 'registro': {
        const f = norm(op.frecuencia);
        const vig = f.includes('SEMESTR') ? 6 : f.includes('TRIMESTR') ? 3 : f.includes('ANUAL') ? 12 : 1;
        const pf = partes(r.fin), idxFin = pf.y * 12 + pf.m;
        const cands = m.registro.filter((x) => x.clave === norm(op.nombre) && x.anio * 12 + x.mes <= idxFin && x.anio * 12 + x.mes > idxFin - vig)
          .sort((a, b) => (b.anio * 12 + b.mes) - (a.anio * 12 + a.mes));
        if (cands.length && cands[0].valor != null) {
          const x = cands[0];
          res.valor = x.valor; res.numerador = x.numerador; res.denominador = x.denominador; res.comentario = x.comentario; res.origenRegistro = `${x.mes}/${x.anio}`;
        } else res.nota = `Sin dato en el registro mensual para el periodo (frecuencia ${op.frecuencia || 'mensual'}).`;
        break;
      }
    }
    return res;
  }

  function contingencia(m, corte) {
    const abiertos = m.casos.filter((c) => abiertoAl(c, corte));
    const procesos = abiertos.filter((c) => c.posicion === 'EN CONTRA' && Math.max(c.pretendido || 0, c.estimado || 0) > 0)
      .sort((a, b) => b.reserva - a.reserva || (b.pretendido || 0) - (a.pretendido || 0));
    const sum = (arr, f) => arr.reduce((a, c) => a + f(c), 0);
    return {
      abiertos: abiertos.length,
      exposicion: sum(procesos, (c) => c.pretendido || 0),
      provision: sum(procesos.filter((c) => c.prob === 'PROBABLE'), (c) => c.reserva),
      nota: sum(procesos.filter((c) => c.prob === 'POSIBLE'), (c) => c.reserva),
      procesos,
      porValorar: abiertos.filter((c) => c.prob === 'POR VALORAR'),
      matriz: ['BAJO', 'MEDIO', 'ALTO'].map((ri) => ['REMOTO', 'POSIBLE', 'PROBABLE'].map((p) => abiertos.filter((c) => c.riesgo === ri && c.prob === p)))
    };
  }

  // ----------------------------------------------------------------- cascada
  function cumplimiento(valor, meta, sentido) {
    if (valor == null || meta == null || !sentido) return null;
    if (sentido === 'Mayor') return meta === 0 ? 1 : Math.min(1, valor / meta);
    if (sentido === 'Menor') { if (valor === 0) return 1; if (meta === 0) return 0; return Math.min(1, meta / valor); }
    return valor === meta ? 1 : 0;
  }
  function estado(cumpl, config) {
    if (cumpl == null) return 'sin';
    if (cumpl >= 0.99995) return 'cumple';
    return (1 - cumpl) * 100 <= config.umbralAlertaPP + 1e-9 ? 'alerta' : 'nocumple';
  }

  /** Calcula todos los indicadores del periodo. Resultado en caché por periodo. */
  function calcular(m, anio, mes) {
    const clave = `${anio}-${mes}-${m.config.modoParcial}`;
    if (m.cache[clave]) return m.cache[clave];
    const r = rango(anio, mes);
    const tacticos = m.cascada.tacticos.map((t) => {
      const ops = t.operativos.map((op) => {
        const res = calcularOperativo(m, op, r);
        const cumpl = op.peso > 0 || op.meta != null ? cumplimiento(res.valor, op.meta, op.sentido) : null;
        return { def: op, ...res, cumpl, estado: op.peso > 0 ? estado(cumpl, m.config) : (op.meta != null ? estado(cumpl, m.config) : 'seguimiento') };
      });
      const sumaPesos = ops.reduce((a, o) => a + (o.def.peso || 0), 0);
      const faltan = ops.filter((o) => o.def.peso > 0 && o.cumpl == null);
      const problemas = [];
      if (Math.abs(sumaPesos - 1) > 0.001) problemas.push(`los pesos operativos suman ${(sumaPesos * 100).toFixed(1).replace('.', ',')} %`);
      faltan.forEach((o) => problemas.push(`falta «${o.def.nombre}»`));
      const resultado = problemas.length ? null : ops.reduce((a, o) => a + (o.def.peso > 0 ? o.def.peso * o.cumpl : 0), 0);
      const cumplT = resultado == null ? null : Math.min(1, t.meta ? resultado / t.meta : resultado);
      return { def: t, operativos: ops, resultado, cumpl: cumplT, estado: estado(cumplT, m.config), problemas, faltan };
    });
    // Estratégico
    const sumaT = tacticos.reduce((a, t) => a + (t.def.peso || 0), 0);
    const conDato = tacticos.filter((t) => t.resultado != null && t.def.peso > 0);
    const sinDato = tacticos.filter((t) => t.resultado == null && t.def.peso > 0);
    let resultado = null, parcial = false;
    const pesosOk = Math.abs(sumaT - 1) <= 0.001;
    if (tacticos.length && pesosOk && !sinDato.length) resultado = tacticos.reduce((a, t) => a + (t.def.peso || 0) * (t.resultado || 0), 0);
    else if (m.config.modoParcial && conDato.length) {
      const pc = conDato.reduce((a, t) => a + t.def.peso, 0);
      resultado = conDato.reduce((a, t) => a + t.def.peso * t.resultado, 0) / pc; parcial = true;
    }
    const metaE = m.cascada.estrategico?.meta ?? 1;
    const cumplE = resultado == null ? null : Math.min(1, metaE ? resultado / metaE : resultado);
    const estrategico = { def: m.cascada.estrategico, resultado, cumpl: cumplE, estado: estado(cumplE, m.config), parcial, sinDato, pesosOk, pesoMedido: conDato.reduce((a, t) => a + t.def.peso, 0) };

    // Contingencia, atención y volumen
    const cont = contingencia(m, r.fin);
    const ref = Math.min(r.fin, hoyDia());
    const abiertos = m.casos.filter((c) => abiertoAl(c, r.fin));
    const vencidos = abiertos.filter((c) => !c.pendiente && c.clas !== 'DEMANDAS' && c.venceANS != null && c.venceANS < ref)
      .map((c) => ({ c, dias: diasHabiles(c.venceANS, ref) })).sort((a, b) => b.dias - a.dias);
    const enEspera = abiertos.filter((c) => c.pendiente);
    const terminos = abiertos.map((c) => {
      const lim = c.limite ?? (config_lim(m, c));
      return lim == null ? null : { c, limite: lim, restan: lim >= ref ? networkdays(ref, lim) - 1 : -(networkdays(lim, ref) - 1) };
    }).filter((x) => x && x.restan <= m.config.diasProximoVencer).sort((a, b) => a.restan - b.restan);
    const atencion = { ref, vencidos, porValorar: cont.porValorar, terminos, enEspera, abiertos };

    const volumen = calcularVolumen(m, r);
    const out = { rango: r, tacticos, estrategico, contingencia: cont, atencion, volumen };
    m.cache[clave] = out;
    return out;
  }
  function config_lim(m, c) {
    const d = m.config.terminoLegalPorDefecto[c.clas];
    return d != null && c.recibido != null ? workday(c.recibido, d) : null;
  }

  function calcularVolumen(m, r) {
    const enRango = (d) => d != null && d >= r.ini && d <= r.fin;
    const recibidos = m.casos.filter((c) => enRango(c.recibido));
    const cerrados = m.casos.filter((c) => cerradoEn(c, r.ini, r.fin));
    const contar = (arr, k) => arr.reduce((o, c) => ((o[c[k]] = (o[c[k]] || 0) + 1), o), {});
    const mensual = [];
    for (let mm = 1; mm <= 12; mm++) {
      const ini = dia(r.anio, mm, 1), fin = finMes(r.anio, mm);
      mensual.push({ mes: mm, recibidos: m.casos.filter((c) => c.recibido != null && c.recibido >= ini && c.recibido <= fin).length, cerrados: m.casos.filter((c) => cerradoEn(c, ini, fin)).length });
    }
    return { recibidos: recibidos.length, cerrados: cerrados.length, porClasRec: contar(recibidos, 'clas'), porClasCer: contar(cerrados, 'clas'), porSede: contar(recibidos, 'sede'), mensual };
  }

  /** Serie mensual del año (para tendencias). */
  function serieAnual(m, anio) {
    const hoy = partes(hoyDia());
    const ultimo = anio < hoy.y ? 12 : anio === hoy.y ? hoy.m : 0;
    const meses = [];
    for (let mm = 1; mm <= ultimo; mm++) meses.push({ mes: mm, res: calcular(m, anio, mm) });
    return meses;
  }

  global.Motor = {
    construir, calcular, serieAnual, rango, anterior, festivos, diasHabiles, networkdays, workday, fecha, fmtFecha, partes, dia,
    norm, cumplimiento, estado, abiertoAl, hoyDia, CONFIG_BASE
  };
})(typeof window !== 'undefined' ? window : globalThis);
