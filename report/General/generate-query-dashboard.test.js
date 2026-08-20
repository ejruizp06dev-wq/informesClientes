const test = require('node:test');
const assert = require('node:assert/strict');
const { buildAnalyticsRows } = require('./generate-query-dashboard');

test('buildAnalyticsRows agrega campos de analítica de consultas', () => {
  const inputRows = [
    {
      Proyecto: 'Proyecto A',
      Encargado: 'Ana',
      Ultima_Modificacion: '2026-08-01T12:00:00Z',
      Horas_Estimadas: '5',
      Horas_Ejecutadas: '3',
      Estado: 'Committed',
      Tipo: 'Task',
      Titulo: 'Consulta de ejemplo'
    }
  ];

  const rows = buildAnalyticsRows(inputRows);
  assert.equal(rows.length, 1);
  assert.equal(rows[0].Developer, 'Ana');
  assert.equal(rows[0].Project, 'Proyecto A');
  assert.equal(rows[0].Query_Date, '2026-08-01');
  assert.equal(rows[0].Execution_Time_Sec, 10800);
  assert.equal(rows[0].Query_Volume, 1);
  assert.equal(rows[0].Query_Status, 'Committed');
});
