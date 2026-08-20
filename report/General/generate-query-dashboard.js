const fs = require('fs');
const path = require('path');
require('dotenv').config();

function parseDate(value) {
  if (!value) return null;
  const text = String(value).trim();

  const isoMatch = text.match(/^(\d{4})-(\d{1,2})-(\d{1,2})/);
  if (isoMatch) {
    const year = Number(isoMatch[1]);
    const month = Number(isoMatch[2]);
    const day = Number(isoMatch[3]);
    const date = new Date(year, month - 1, day);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const slashMatch = text.match(/^(\d{1,2})[\/\-](\d{1,2})[\/\-](\d{4})/);
  if (slashMatch) {
    const day = Number(slashMatch[1]);
    const month = Number(slashMatch[2]);
    const year = Number(slashMatch[3]);
    const date = new Date(year, month - 1, day);
    return Number.isNaN(date.getTime()) ? null : date;
  }

  const date = new Date(text);
  return Number.isNaN(date.getTime()) ? null : date;
}

function formatDateKey(date) {
  if (!date) return 'Sin fecha';
  return `${date.getFullYear()}-${String(date.getMonth() + 1).padStart(2, '0')}-${String(date.getDate()).padStart(2, '0')}`;
}

function buildAnalyticsRows(rows) {
  const fechaInicio = process.env.FECHA_INICIO_ESPECIFICO ? new Date(`${process.env.FECHA_INICIO_ESPECIFICO}T00:00:00`) : null;
  const fechaFin = process.env.FECHA_FIN_ESPECIFICO ? new Date(`${process.env.FECHA_FIN_ESPECIFICO}T23:59:59`) : null;

  return rows.flatMap((row) => {
    const parsedDate = parseDate(row.Ultima_Modificacion || row['Ultima_Modificacion'] || row['Última Modificación']);
    const dateKey = formatDateKey(parsedDate);
    const inRange = (!fechaInicio || !fechaFin || !parsedDate)
      ? true
      : (parsedDate >= fechaInicio && parsedDate <= fechaFin);

    if (!inRange) {
      return [];
    }
    const est = Number.parseFloat(row.Horas_Estimadas || row['Horas_Estimadas'] || 0) || 0;
    const ejec = Number.parseFloat(row.Horas_Ejecutadas || row['Horas_Ejecutadas'] || 0) || 0;
    const volume = Number.parseInt(row.Cant_Tareas || row['Cant_Tareas'] || row.Query_Volume || 1, 10) || 1;
    const executionSeconds = Math.max(1, Math.round((ejec || est || 1) * 3600));

    const record = {
      Developer: row.Encargado || row.Developer || 'Sin asignar',
      Project: row.Proyecto || row.Project || 'Sin proyecto',
      Query_Date: dateKey,
      Query_Status: row.Estado || row.Query_Status || 'Sin estado',
      Query_Volume: volume,
      Execution_Time_Sec: executionSeconds,
      Avg_Execution_Time_Sec: executionSeconds,
      Metric_Type: row.Tipo || row.Metric_Type || 'Task',
      Estimated_Hours: est,
      Executed_Hours: ejec,
      Progress_Percent: Number.parseFloat(row.Porcentaje_Avance || row['Progreso'] || 0) || 0,
      Source_Row: row.Titulo || row['HU_Titulo'] || 'Sin título',
      Query_Text: row.Titulo || row['HU_Titulo'] || row.Query_Text || 'Sin descripción',
    };
    return record ? [record] : [];
  });
}

function resolveInputPath() {
  const candidates = [
    path.join(process.cwd(), 'dashboard-ui', 'public', 'data', 'dashboard.csv'),
    path.join(process.cwd(), 'report', 'tables', 'dashboard.csv'),
    path.join(process.cwd(), 'dashboard.csv'),
  ];

  for (const candidate of candidates) {
    if (fs.existsSync(candidate)) {
      return candidate;
    }
  }

  return path.join(process.cwd(), 'dashboard.csv');
}

function writeLocalQueryAnalytics(inputPath, outputPath) {
  const raw = fs.readFileSync(inputPath, 'utf8');
  const lines = raw.trim().split(/\r?\n/).filter(Boolean);
  if (lines.length < 2) return [];

  const headers = lines[0].split(',').map((h) => h.trim().replace(/^"|"$/g, ''));
  const rows = [];

  for (let i = 1; i < lines.length; i += 1) {
    const values = [];
    let current = '';
    let inQuotes = false;

    for (let j = 0; j < lines[i].length; j += 1) {
      const char = lines[i][j];
      const next = lines[i][j + 1];
      if (char === '"') {
        if (inQuotes && next === '"') {
          current += '"';
          j += 1;
        } else {
          inQuotes = !inQuotes;
        }
      } else if (char === ',' && !inQuotes) {
        values.push(current.trim().replace(/^"|"$/g, ''));
        current = '';
      } else {
        current += char;
      }
    }
    values.push(current.trim().replace(/^"|"$/g, ''));

    if (values.length === headers.length) {
      const row = {};
      headers.forEach((header, idx) => {
        row[header] = values[idx];
      });
      rows.push(row);
    }
  }

  const analyticsRows = buildAnalyticsRows(rows);
  const outputHeaders = [
    'Developer',
    'Project',
    'Query_Date',
    'Query_Status',
    'Query_Volume',
    'Execution_Time_Sec',
    'Avg_Execution_Time_Sec',
    'Metric_Type',
    'Estimated_Hours',
    'Executed_Hours',
    'Progress_Percent',
    'Source_Row',
    'Query_Text',
  ];

  const csv = [outputHeaders.join(',')].concat(
    analyticsRows.map((row) => outputHeaders.map((header) => `"${String(row[header] ?? '').replace(/"/g, '""')}"`).join(','))
  ).join('\n');

  fs.mkdirSync(path.dirname(outputPath), { recursive: true });
  fs.writeFileSync(outputPath, csv, 'utf8');
  return analyticsRows;
}

if (require.main === module) {
  const inputPath = resolveInputPath();
  const outputPath = path.join(process.cwd(), 'dashboard-query-analytics.csv');
  const rows = writeLocalQueryAnalytics(inputPath, outputPath);
  console.log(`Se generaron ${rows.length} registros de analítica desde ${inputPath} en ${outputPath}`);
}

module.exports = { buildAnalyticsRows, writeLocalQueryAnalytics };
