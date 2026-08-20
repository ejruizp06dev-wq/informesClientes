const test = require('node:test');
const assert = require('node:assert/strict');
const fs = require('fs');
const path = require('path');
const os = require('os');

const { obtenerRutaDatosLocales } = require('./task-auditor');

test('obtenerRutaDatosLocales devuelve un CSV local cuando existe', () => {
  const tempDir = fs.mkdtempSync(path.join(os.tmpdir(), 'task-auditor-'));
  const csvPath = path.join(tempDir, 'dashboard.csv');
  fs.writeFileSync(csvPath, 'Proyecto,Estado\nDemo,Done\n', 'utf8');

  const originalCwd = process.cwd();
  process.chdir(tempDir);

  try {
    const resolved = obtenerRutaDatosLocales();
    assert.equal(resolved, csvPath);
  } finally {
    process.chdir(originalCwd);
    fs.rmSync(tempDir, { recursive: true, force: true });
  }
});
