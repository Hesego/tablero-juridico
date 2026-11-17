// Pruebas del motor de cálculo (sin dependencias). Ejecutar: node --test pruebas/
const test = require('node:test');
const assert = require('node:assert/strict');
require('../motor.js');
const M = globalThis.Motor;

test('festivos de Colombia 2026 (fijos, Ley 51 de 1983 y Pascua)', () => {
  const esperados = ['01/01/2026', '12/01/2026', '23/03/2026', '02/04/2026', '03/04/2026', '01/05/2026', '18/05/2026', '08/06/2026',
    '15/06/2026', '29/06/2026', '20/07/2026', '07/08/2026', '17/08/2026', '12/10/2026', '02/11/2026', '16/11/2026', '08/12/2026', '25/12/2026'];
  const obtenidos = [...M.festivos(2026)].sort((a, b) => a - b).map(M.fmtFecha);
  assert.deepEqual(obtenidos, esperados);
});

test('días hábiles = NETWORKDAYS − 1, sin fines de semana ni festivos', () => {
  const d = M.dia;
  assert.equal(M.diasHabiles(d(2026, 10, 1), d(2026, 10, 9)), 6);   // caso 404
  assert.equal(M.diasHabiles(d(2026, 9, 25), d(2026, 10, 5)), 6);   // caso 428
  assert.equal(M.diasHabiles(d(2026, 10, 14), d(2026, 10, 30)), 12); // caso 471
  assert.equal(M.diasHabiles(d(2026, 2, 21), d(2026, 2, 22)), 0);   // sábado → domingo
  assert.equal(M.workday(d(2026, 11, 13), 2), d(2026, 11, 18));     // tutela: 16/11 es festivo
});

test('fechas: serial de Excel y texto dd/mm/aaaa', () => {
  assert.equal(M.fmtFecha(M.fecha(46296)), '01/10/2026');
  assert.equal(M.fmtFecha(M.fecha('05/11/2026')), '05/11/2026');
  assert.equal(M.fecha('N/A'), null);
});

test('cumplimiento operativo según sentido', () => {
  assert.equal(M.cumplimiento(0.9412, 1, 'Mayor'), 0.9412);
  assert.equal(M.cumplimiento(1.2, 1, 'Mayor'), 1);
  assert.equal(M.cumplimiento(2.15, 5, 'Menor'), 1);
  assert.equal(M.cumplimiento(10, 5, 'Menor'), 0.5);
  assert.equal(M.cumplimiento(0, 0, 'Menor'), 1);
  assert.equal(M.cumplimiento(1, 1, 'Igual'), 1);
  assert.equal(M.cumplimiento(0.9, 1, 'Igual'), 0);
});

// Modelo mínimo sintético para probar la cascada de punta a punta.
function hojasDePrueba() {
  const enc = ['No.', 'CLASIFICACIÓN DEL PROCESO JURÍDICO', 'NOMBRE', 'ASUNTO DE LA PETICIÓN', 'PROCESO QUE REMITE/SEDE', 'FECHA DE RECIBIDO',
    'TIPO DE TÉRMINO (HÁBIL/CALENDARIO)', 'FECHA LÍMITE RESPUESTA', 'FECHA EN QUE DI RESPUESTA', 'ESTADO', 'RIESGO JURÍDICO', 'VALOR ESTIMADO DEL RIESGO',
    'PENDIENTE DE REVISIÓN POR PROCESO ENCARGADO', 'POSICIÓN DE HESEGO\n(EN CONTRA / A FAVOR)', 'VALOR PRETENDIDO', 'SOLIDEZ DE LA DEMANDA\n(1-5)',
    'PRUEBAS EN CONTRA\n(1-5)', 'JURISPRUDENCIA DESFAVORABLE\n(1-5)', 'CONDUCTA DE LA CONTRAPARTE\n(1-5)', 'VALOR FINAL\n(a cargo de Hesego o recuperado)'];
  const f = (id, clas, rec, resp, estado, extra = {}) => [id, clas, 'Nombre ' + id, 'Asunto', 'GENERAL', rec, 'N/A', extra.lim ?? 'N/A', resp, estado, extra.riesgo ?? 'BAJO',
    extra.est ?? 0, 'NO', extra.pos ?? null, extra.pre ?? null, ...(extra.fac ?? [null, null, null, null]), extra.fin ?? null];
  const gh47 = [enc,
    f(1, 'GESTION JURIDICA', '01/10/2026', '02/10/2026', 'CERRADO'),
    f(2, 'GESTION JURIDICA', '01/10/2026', '30/10/2026', 'CERRADO'),           // 20 días hábiles > ANS 10
    f(3, 'TUTELAS', '01/10/2026', '02/10/2026', 'CERRADO'),
    f(4, 'PROCESOS DISCIPLINARIOS', '01/10/2026', '02/10/2026', 'CERRADO'),
    f(5, 'SINIESTROS', '01/09/2026', null, 'PENDIENTE', { riesgo: 'ALTO', est: 35000000, pos: 'EN CONTRA', pre: 35000000, fac: [4, 4, 4, 2] }),
    f(6, 'DEMANDAS', '01/09/2026', null, 'EN GESTIÓN', { riesgo: 'MEDIO', est: 'texto', pos: 'EN CONTRA', pre: 30000000, fac: [2, 3, 3, 2] })
  ];
  const ig = Array.from({ length: 7 }, () => []);
  const fila = (o) => { const r = []; for (const [k, v] of Object.entries(o)) r[k.charCodeAt(0) - 65] = v; return r; };
  ig.push(fila({ B: 'Índice', E: 1, F: 0.6, G: 'Oportunidad en la atención de acciones constitucionales y legales', J: 1, L: 'Solicitudes constitucionales y legales cerradas dentro del ANS', M: 0.5, O: 1, P: 'Mayor' }));
  ig.push(fila({ L: 'Términos legales cumplidos en tutelas y derechos de petición', M: 0.5, O: 1, P: 'Mayor' }));
  ig.push(fila({ F: 0.4, G: 'Control del riesgo y la contingencia jurídica', J: 1, L: 'Cobertura de la valoración del riesgo en procesos abiertos', M: 1, O: 1, P: 'Mayor' }));
  ig.push(fila({ L: 'Exposición jurídica y reserva requerida (seguimiento)', M: 0 }));
  return { gh47, ig136: ig };
}

test('cascada, contingencia y calidad con datos sintéticos', () => {
  const m = M.construir(hojasDePrueba(), { usarHojaANS: false });
  const r = M.calcular(m, 2026, 10);
  const [t1, t2] = r.tacticos;
  assert.equal(t1.operativos[0].numerador, 2);
  assert.equal(t1.operativos[0].denominador, 3);
  assert.equal(t1.operativos[1].valor, 1);
  assert.ok(Math.abs(t1.resultado - (0.5 * 2 / 3 + 0.5)) < 1e-9);
  assert.equal(t2.operativos[0].valor, 1);                 // ambos abiertos valorados
  assert.equal(r.contingencia.exposicion, 65000000);
  assert.equal(r.contingencia.provision, 35000000 * 0.7);  // probable 70 %
  assert.equal(r.contingencia.nota, 30000000 * 0.5);       // estimado en texto → pretendido × 50 %
  assert.ok(Math.abs(r.estrategico.resultado - (0.6 * t1.resultado + 0.4)) < 1e-9);
  assert.ok(m.calidad.some((q) => q.tipo === 'Valor estimado en texto'));
});
