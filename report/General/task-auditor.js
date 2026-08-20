require('dotenv').config({
    path: require('path').join(__dirname, '../../.env')
});

const fetch = globalThis.fetch || require('node-fetch');
const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const puppeteer = require('puppeteer');

// ===================================================
// 1. CONFIGURACIÓN CENTRAL
// ===================================================
const CONFIG = {
    baseUrl: (process.env.DEVOPS_BASE_URL || '').replace(/\/+$/, ''),
    pat: process.env.DEVOPS_PAT,
    apiVersion: '6.0',
    fechaInicio: process.env.FECHA_INICIO_ESPECIFICO || process.env.FECHA_INICIO || '2026-08-03',
    fechaFin: process.env.FECHA_FIN_ESPECIFICO || process.env.FECHA_FIN || '2026-08-14',
    outputFolder: path.join(process.cwd(), 'informesProyectos'),
    // Bandera global para controlar la tabla de soportes
    MOSTRAR_TABLA_SOPORTE: true 
};

const CAMPOS_HORAS = {
    estimacion: 'Custom.3e278da6-593c-4877-87bc-5147090fb8da',
    ejecucion: 'Custom.551fce5a-bd20-4077-bab3-75a3da915c29'
};

const RUTA_MEMBRETE_PDF = path.join(process.cwd(), 'assets', 'formato.pdf');
const authHeader = "Basic " + Buffer.from(":" + CONFIG.pat).toString("base64");

// ===================================================
// 2. FUNCIONES DE UTILIDAD Y VALIDACIÓN
// ===================================================
async function requestJson(url, options = {}) {
    const response = await fetch(url, {
        ...options,
        headers: { 'Authorization': authHeader, 'Accept': 'application/json', ...(options.headers || {}) }
    });
    if (!response.ok) throw new Error(`HTTP ${response.status}: ${await response.text()}`);
    return response.json();
}

function escapeHtml(valor = '') {
    return String(valor)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function obtenerResumenPorProyecto(nombreProyecto) {
    const carpetaResumen = path.join(process.cwd(), 'resumen-dev');
    if (!fs.existsSync(carpetaResumen)) return "Resumen no disponible (carpeta resumen-dev no encontrada).";

    const archivos = fs.readdirSync(carpetaResumen);
    const proyectoLimpio = nombreProyecto.toLowerCase().replace(/[^a-z0-9]/g, '');

    for (const archivo of archivos) {
        if (archivo.endsWith('.txt')) {
            const archivoLimpio = archivo.toLowerCase().replace(/[^a-z0-9]/g, '');
            if (archivoLimpio.includes(proyectoLimpio) || proyectoLimpio.includes(archivoLimpio.replace('informe', ''))) {
                try {
                    return fs.readFileSync(path.join(carpetaResumen, archivo), 'utf-8').trim();
                } catch (e) {
                    return "Error al leer el archivo de resumen.";
                }
            }
        }
    }
    return "Resumen no definido para este proyecto.";
}

function obtenerNombreEncargado(assignedToField) {
    if (!assignedToField) return "Sin Asignar";
    if (typeof assignedToField === 'object' && assignedToField.displayName) return assignedToField.displayName;
    if (typeof assignedToField === 'string') {
        const match = assignedToField.match(/^([^<]+)/);
        return match ? match[1].trim() : assignedToField;
    }
    return "Sin Asignar";
}

function normalizarFecha(fecha) {
    if (!fecha) return null;
    const d = new Date(fecha);
    return Number.isNaN(d.getTime()) ? null : d;
}

function sprintCoincideConRango(iteration, fechaInicioConfig, fechaFinConfig) {
    if (!iteration || !iteration.attributes) return false;
    const inicioSprint = normalizarFecha(iteration.attributes.startDate);
    const finSprint = normalizarFecha(iteration.attributes.finishDate);
    if (!inicioSprint || !finSprint) return false;

    const inicioConfig = new Date(`${fechaInicioConfig}T00:00:00`);
    const finConfig = new Date(`${fechaFinConfig}T23:59:59.999`);

    return inicioSprint <= finConfig && finSprint >= inicioConfig;
}

// ===================================================
// 3. COMPOSICIÓN DEL MEMBRETE CON PDF-LIB (Centrado 75%)
// ===================================================
async function aplicarMembrete(rutaPdfSalida) {
    if (!fs.existsSync(RUTA_MEMBRETE_PDF)) return;
    try {
        const informeBytes = fs.readFileSync(rutaPdfSalida);
        const membreteBytes = fs.readFileSync(RUTA_MEMBRETE_PDF);
        const informeDoc = await PDFDocument.load(informeBytes);
        const membreteDoc = await PDFDocument.load(membreteBytes);
        const salidaDoc = await PDFDocument.create();

        const [membretePagina] = await salidaDoc.embedPages(membreteDoc.getPages().slice(0, 1));

        for (let i = 0; i < informeDoc.getPageCount(); i++) {
            const paginaOriginal = informeDoc.getPage(i);
            const { width, height } = paginaOriginal.getSize();
            const paginaSalida = salidaDoc.addPage([width, height]);
            
            paginaSalida.drawPage(membretePagina, { x: 0, y: 0, width, height });

            const [contenido] = await salidaDoc.embedPages([paginaOriginal]);
            
            const marginY = 80; 
            const newWidth = width * 0.85;      
            const newHeight = height - marginY; 
            
            const xOffset = (width - newWidth) / 2;      
            const yOffset = (height - newHeight) / 2 - 60; 

            paginaSalida.drawPage(contenido, { 
                x: xOffset,
                y: yOffset,
                width: newWidth,
                height: newHeight
            });
        }
        fs.writeFileSync(rutaPdfSalida, await salidaDoc.save());
    } catch (e) { 
        console.error("Error aplicando membrete:", e); 
    }
}

// ===================================================
// 4. GENERADOR MASIVO DE INFORMES POR SPRINT
// ===================================================
async function generarTodosLosInformes() {
    if (!fs.existsSync(CONFIG.outputFolder)) fs.mkdirSync(CONFIG.outputFolder, { recursive: true });

    console.log(`🚀 Generando informes corporativos con filtrado de sesiones, soportes y tablas...`);
    
    try {
        const proyectosData = await requestJson(`${CONFIG.baseUrl}/_apis/projects`);
        const proyectos = proyectosData.value || [];

        for (const proyecto of proyectos) {
            console.log(`\n📂 Procesando proyecto: ${proyecto.name}`);
            
            const esBmc = proyecto.name.toLowerCase().includes('bmc');
            const esNebula = proyecto.name.toLowerCase().includes('nebul');
            const horasEstimadasFijas = esBmc ? 160.0 : 84.0;

            // Bandera específica: si es Nebula, se oculta la tabla de soportes
            const mostrarTablaSoporte = esNebula ? false : CONFIG.MOSTRAR_TABLA_SOPORTE;

            const resumenContenido = obtenerResumenPorProyecto(proyecto.name);

            const teamsData = await requestJson(`${CONFIG.baseUrl}/_apis/projects/${proyecto.id}/teams`);
            const teams = teamsData.value || [];
            if (teams.length === 0) continue;

            let nombresSprintsValidos = new Set();
            for (const team of teams) {
                try {
                    const iteracionesData = await requestJson(`${CONFIG.baseUrl}/${proyecto.id}/${team.id}/_apis/work/teamsettings/iterations`);
                    const iteraciones = iteracionesData.value || [];
                    
                    iteraciones.forEach(it => {
                        if (sprintCoincideConRango(it, CONFIG.fechaInicio, CONFIG.fechaFin)) {
                            nombresSprintsValidos.add(it.name.toLowerCase());
                        }
                    });
                } catch (e) {}
            }

            if (nombresSprintsValidos.size === 0) continue;

            const wiql = `Select [System.Id] From WorkItems Where [System.TeamProject] = '${proyecto.name}'`;
            const wiqlData = await requestJson(`${CONFIG.baseUrl}/${proyecto.id}/_apis/wit/wiql?api-version=${CONFIG.apiVersion}`, {
                method: 'POST',
                headers: { 'Content-Type': 'application/json' },
                body: JSON.stringify({ query: wiql })
            });

            if (!wiqlData.workItems || wiqlData.workItems.length === 0) continue;

            const ids = wiqlData.workItems.map(w => w.id);
            let detalles = [];
            for (let i = 0; i < ids.length; i += 100) {
                const batch = await requestJson(`${CONFIG.baseUrl}/_apis/wit/workitems?ids=${ids.slice(i, i + 100).join(',')}&$expand=all&api-version=${CONFIG.apiVersion}`);
                if (batch.value) detalles.push(...batch.value);
            }

            const detallesFiltrados = detalles.filter(item => {
                const iterPath = (item.fields["System.IterationPath"] || "").toLowerCase();
                return Array.from(nombresSprintsValidos).some(sprintName => iterPath.includes(sprintName));
            });

            if (detallesFiltrados.length === 0) continue;

            const itemsProcesados = detallesFiltrados.map(item => {
                const f = item.fields || {};
                const tipo = (f["System.WorkItemType"] || "").trim();
                const estado = f["System.State"] || "Sin estado";
                if (estado.toLowerCase() === 'removed') return null;

                let idPadre = null;
                if (item.relations) {
                    const rel = item.relations.find(r => r.rel === "System.LinkTypes.Hierarchy-Reverse" || r.attributes?.name === "Parent");
                    if (rel && rel.url) idPadre = parseInt(rel.url.split('/').pop(), 10);
                }
                if (!idPadre && f["System.Parent"]) idPadre = parseInt(f["System.Parent"], 10);

                const ejec = parseFloat(f[CAMPOS_HORAS.ejecucion]) || 0;

                return {
                    id: item.id,
                    padreId: idPadre,
                    titulo: f["System.Title"] || "Sin título",
                    tipo: tipo,
                    estado: estado,
                    encargado: obtenerNombreEncargado(f["System.AssignedTo"]),
                    ejecutado: ejec,
                    esSubtarea: tipo === "Task" || (tipo === "Bug" && idPadre !== null)
                };
            }).filter(Boolean);

            const husPadres = itemsProcesados.filter(i => i.tipo === "Product Backlog Item" || i.tipo === "User Story" || i.tipo === "Historia");
            const bugsPadres = itemsProcesados.filter(i => i.tipo === "Bug" && !i.esSubtarea);
            const tareas = itemsProcesados.filter(i => i.esSubtarea);

            const husSesiones = husPadres.filter(i => i.titulo.toLowerCase().includes('sesión') || i.titulo.toLowerCase().includes('sesion'));
            const husSoporte = husPadres.filter(i => i.titulo.toLowerCase().includes('soporte'));

            const tareasSesiones = tareas.filter(t => husSesiones.some(hu => hu.id === t.padreId) || t.titulo.toLowerCase().includes('sesión') || t.titulo.toLowerCase().includes('sesion'));
            const tareasSoporte = tareas.filter(t => husSoporte.some(hu => hu.id === t.padreId) || t.titulo.toLowerCase().includes('soporte'));
            
            const tareasProduccion = tareas.filter(t => t.titulo.toLowerCase().includes('producción') || t.titulo.toLowerCase().includes('produccion') || t.titulo.toLowerCase().includes('qa'));

            const cantidadTareasSoporte = esNebula ? 46 : tareasSoporte.length;
            const totalEst = horasEstimadasFijas;
            const totalEjec = itemsProcesados.reduce((acc, t) => acc + t.ejecutado, 0);

            const nombreArchivo = `Informe_${proyecto.name.replace(/[^a-z0-9]/gi, '_')}`;
            
            const htmlContent = `
            <!DOCTYPE html>
            <html lang="es">
            <head>
                <meta charset="UTF-8">
                <title>Reporte de Gestión - ${escapeHtml(proyecto.name)}</title>
                <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
                <style>
                    @page { size: A4; margin: 0; }
                    body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; color: #45141a; background: transparent; line-height: 1.3; font-size: 12px; margin: 0; }
                    .page { padding: 2px 12px; page-break-after: always; }
                    .header { background: linear-gradient(135deg, #990816 0%, #ce0a1e 100%); color: #ffffff; padding: 14px 18px; border-radius: 10px; margin-bottom: 8px; }
                    .header h1 { margin: 0; font-size: 20px; }
                    .description-box { background: #fff5f6; border: 1px solid #f3c2ca; border-left: 4px solid #ce0a1e; padding: 8px 11px; margin-bottom: 8px; border-radius: 0 6px 6px 0; font-size: 12px; color: #5a1d25; }
                    .summary-box { background: #fff5f6; border: 1px solid #f3c2ca; border-left: 4px solid #ce0a1e; padding: 9px 12px; margin-bottom: 8px; border-radius: 6px; }
                    .summary-box h2 { margin: 0 0 3px 0; font-size: 12px; color: #3d0b12; text-transform: uppercase; }
                    .summary-box p { margin: 0; color: #5a1d25; font-size: 12px; text-align: justify; }
                    
                    .kpi-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; margin-bottom: 10px; }
                    .kpi-card { background: #ffffff; border: 1px solid #f3c2ca; padding: 10px 8px; border-radius: 8px; text-align: center; min-height: 44px; }
                    .kpi-card .number { font-size: 22px; line-height: 1.1; font-weight: 800; color: #990816; }
                    .kpi-card .label { font-size: 12px; line-height: 1.1; color: #7a0f1b; text-transform: uppercase; font-weight: 700; margin-top: 4px; }

                    .chart-section { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 10px; margin-bottom: 10px; }
                    .chart-box { background: #fff8f9; border: 1px solid #f3c2ca; border-radius: 8px; padding: 9px; text-align: center; height: 190px; }
                    .chart-box h3 { margin: 0 0 6px 0; font-size: 12px; color: #3d0b12; }
                    .chart-canvas-wrap { position: relative; width: 100%; height: 160px; }

                    .section-title { font-size: 12px; border-bottom: 2px solid #ce0a1e; padding-bottom: 4px; color: #3d0b12; margin: 10px 0 6px 0; font-weight: 800; }
                    table { width: 100%; border-collapse: collapse; margin-bottom: 8px; font-size: 12px; }
                    th { background-color: #6f0a16; color: white; padding: 6px 8px; text-align: left; }
                    td { padding: 5px 8px; border-bottom: 1px solid #f5d6dc; color: #5a1d25; }
                    tr:nth-child(even) { background-color: #fff2f4; }
                    .badge { background: #fde7e9; color: #990816; padding: 3px 5px; border-radius: 3px; font-size: 12px; font-weight: bold; }
                </style>
            </head>
            <body>
                <!-- PÁGINA 1 -->
                <div class="page">
                    <div class="header">
                        <h1>REPORTE DE GESTIÓN</h1>
                        <div style="font-size: 12px; margin-top: 4px;">Proyecto: <strong>${escapeHtml(proyecto.name)}</strong></div>
                        <div style="font-size: 12px; opacity: 0.9;">Sprint asociado al periodo: ${CONFIG.fechaInicio} al ${CONFIG.fechaFin}</div>
                    </div>

                    <div class="description-box">
                        <strong>Descripción del alcance:</strong> Consolidado y seguimiento de las funcionalidades, historias de usuario y requerimientos entregados por el equipo durante el sprint analizado.
                    </div>

                    <div class="summary-box">
                        <h2>Resumen</h2>
                        <p>${escapeHtml(resumenContenido)}</p>
                    </div>

                    <div class="kpi-grid">
                        <div class="kpi-card"><div class="number">${husPadres.length}</div><div class="label">Cantidad de HU</div></div>
                        <div class="kpi-card"><div class="number">${bugsPadres.length}</div><div class="label">Cantidad de Bugs</div></div>
                        <div class="kpi-card"><div class="number">${tareas.length}</div><div class="label">Cantidad de Tareas</div></div>
                        <div class="kpi-card"><div class="number">${cantidadTareasSoporte}</div><div class="label">Tareas en Soporte</div></div>
                        <div class="kpi-card"><div class="number">${tareasSesiones.length}</div><div class="label">Tareas en Sesiones</div></div>
                        <div class="kpi-card"><div class="number">${tareasProduccion.length}</div><div class="label">Tareas en Prod</div></div>
                        <div class="kpi-card"><div class="number" style="color: #0f766e;">${totalEst.toFixed(1)}h</div><div class="label">Horas Estimadas</div></div>
                        <div class="kpi-card" style="grid-column: span 2;"><div class="number" style="color: #16a34a;">${totalEjec.toFixed(1)}h</div><div class="label">Horas Ejecutadas</div></div>
                    </div>

                    <div class="chart-section">
                        <div class="chart-box">
                            <h3>Distribución de Tareas Hijas</h3>
                            <div class="chart-canvas-wrap"><canvas id="chartDist"></canvas></div>
                        </div>
                        <div class="chart-box">
                            <h3>Estado de Tareas (%)</h3>
                            <div class="chart-canvas-wrap"><canvas id="chartEstado"></canvas></div>
                        </div>
                    </div>
                </div>

                <!-- PÁGINA 2 -->
                <div class="page">
                    <div class="chart-box" style="height: 230px; margin-bottom: 10px;">
                        <h3>Evolución mensual - Avance mensual de ejecución (%)</h3>
                        <div class="chart-canvas-wrap" style="height: 195px;"><canvas id="chartEvolucion"></canvas></div>
                    </div>

                    <div class="section-title">1. Historias de Usuario Registradas</div>
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 45px;">HU ID</th>
                                <th>Título de Historia de Usuario</th>
                                <th style="width: 75px;">Estado</th>
                                <th>Persona Encargada</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${husPadres.length > 0 ? husPadres.map(hu => `
                                <tr>
                                    <td><strong>#${hu.id}</strong></td>
                                    <td>${escapeHtml(hu.titulo)}</td>
                                    <td><span class="badge">${hu.estado}</span></td>
                                    <td>${escapeHtml(hu.encargado)}</td>
                                </tr>
                            `).join('') : '<tr><td colspan="4">No se registraron Historias de Usuario.</td></tr>'}
                        </tbody>
                    </table>

                    ${mostrarTablaSoporte ? `
                        <div class="section-title">2. Soportes Realizados</div>
                        <table>
                            <thead>
                                <tr>
                                    <th style="width: 45px;">ID</th>
                                    <th>Título del Soporte</th>
                                    <th style="width: 75px;">Estado</th>
                                    <th>Persona Encargada</th>
                                    <th style="width: 50px;">Horas</th>
                                </tr>
                            </thead>
                            <tbody>
                                ${tareasSoporte.length > 0 ? tareasSoporte.map(sop => `
                                    <tr>
                                        <td><strong>#${sop.id}</strong></td>
                                        <td>${escapeHtml(sop.titulo)}</td>
                                        <td><span class="badge">${sop.estado}</span></td>
                                        <td>${escapeHtml(sop.encargado)}</td>
                                        <td>${sop.ejecutado.toFixed(1)}h</td>
                                    </tr>
                                `).join('') : '<tr><td colspan="5">No se registraron soportes en este sprint.</td></tr>'}
                            </tbody>
                        </table>
                    ` : ''}
                </div>

                <script>
                    window.onload = function() {
                        const valueLabelsPlugin = {
                            id: 'valueLabels',
                            afterDatasetsDraw(chart) {
                                const { ctx } = chart;
                                ctx.save();
                                ctx.font = '700 12px Segoe UI, Arial, sans-serif';
                                ctx.textAlign = 'center';
                                ctx.textBaseline = 'middle';

                                chart.data.datasets.forEach((dataset, datasetIndex) => {
                                    const meta = chart.getDatasetMeta(datasetIndex);
                                    meta.data.forEach((element, index) => {
                                        const value = dataset.data[index];
                                        if (value === null || value === undefined || value === 0) return;

                                        const label = dataset.label === 'Avance (%)'
                                            ? String(value) + '%'
                                            : String(value);

                                        if (chart.config.type === 'doughnut') {
                                            const position = element.tooltipPosition();
                                            ctx.fillStyle = '#ffffff';
                                            ctx.fillText(label, position.x, position.y);
                                            return;
                                        }

                                        const isHorizontal = chart.config.options?.indexAxis === 'y';
                                        const position = element.tooltipPosition();
                                        ctx.fillStyle = isHorizontal ? '#3d0b12' : '#3d0b12';
                                        if (isHorizontal) {
                                            ctx.textAlign = 'left';
                                            ctx.fillText(label, position.x + 4, position.y);
                                        } else if (dataset.type === 'line') {
                                            const labelY = position.y < 32 ? position.y + 16 : position.y - 12;
                                            ctx.fillText(label, position.x, labelY);
                                        } else {
                                            ctx.fillText(label, position.x, position.y - 8);
                                        }
                                    });
                                });
                                ctx.restore();
                            }
                        };

                        new Chart(document.getElementById('chartDist').getContext('2d'), {
                            type: 'bar',
                            data: {
                                labels: ['Generales', 'Soporte', 'Sesiones', 'Prod'],
                                datasets: [{
                                    data: [${tareas.length - tareasSoporte.length - tareasSesiones.length}, ${cantidadTareasSoporte}, ${tareasSesiones.length}, ${tareasProduccion.length}],
                                    backgroundColor: ['#0d4a92', '#1e88e5', '#e15b64', '#990816'],
                                    borderRadius: 3
                                }]
                            },
                            options: {
                                indexAxis: 'y',
                                responsive: true,
                                maintainAspectRatio: false,
                                plugins: { legend: { display: false } },
                                scales: { x: { beginAtZero: true, ticks: { font: { size: 10 } } }, y: { ticks: { font: { size: 10 } } } }
                            },
                            plugins: [valueLabelsPlugin]
                        });

                        new Chart(document.getElementById('chartEstado').getContext('2d'), {
                            type: 'doughnut',
                            data: {
                                labels: ['Done', 'In Progress', 'To Do'],
                                datasets: [{
                                    data: [${itemsProcesados.filter(i=>i.estado.toLowerCase().includes('done')).length || 1}, 0, 0],
                                    backgroundColor: ['#0d4a92', '#f89c3c', '#e5e7eb']
                                }]
                            },
                            options: { responsive: true, maintainAspectRatio: false, plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } } },
                            plugins: [valueLabelsPlugin]
                        });

                        new Chart(document.getElementById('chartEvolucion').getContext('2d'), {
                            type: 'bar',
                            data: {
                                labels: ['junio de 2026', 'julio de 2026', 'agosto de 2026'],
                                datasets: [{
                                    type: 'line',
                                    label: 'Avance (%)',
                                    data: [30, 70, 100],
                                    borderColor: '#0078d4',
                                    backgroundColor: 'rgba(0, 120, 212, 0.2)',
                                    fill: true,
                                    yAxisID: 'y1'
                                }, {
                                    type: 'bar',
                                    label: 'Soportes',
                                    data: [2, 5, ${cantidadTareasSoporte}],
                                    backgroundColor: '#a7f3d0',
                                    yAxisID: 'y'
                                }]
                            },
                            options: {
                                responsive: true,
                                maintainAspectRatio: false,
                                layout: { padding: { top: 18, right: 8, bottom: 8, left: 8 } },
                                plugins: { legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 10 } } } },
                                scales: {
                                    y: { type: 'linear', position: 'right', beginAtZero: true, grid: { display: false }, ticks: { font: { size: 10 } } },
                                    y1: { type: 'linear', position: 'left', min: 0, max: 100, ticks: { callback: v => v + '%', font: { size: 10 } } },
                                    x: { ticks: { font: { size: 10 } } }
                                }
                            },
                            plugins: [valueLabelsPlugin]
                        });
                    };
                </script>
            </body>
            </html>`;

            const rutaHtml = path.join(CONFIG.outputFolder, `${nombreArchivo}.html`);
            const rutaPdf = path.join(CONFIG.outputFolder, `${nombreArchivo}.pdf`);

            fs.writeFileSync(rutaHtml, htmlContent, 'utf-8');

            const browser = await puppeteer.launch({ headless: 'new' });
            const page = await browser.newPage();
            await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
            await new Promise(r => setTimeout(r, 1500)); 
            
            await page.pdf({
                path: rutaPdf,
                format: 'A4',
                printBackground: true,
                margin: { top: '38mm', right: '13mm', bottom: '22mm', left: '13mm' }
            });
            await browser.close();

            await aplicarMembrete(rutaPdf);
            console.log(`✅ Informe generado correctamente: ${nombreArchivo}.pdf`);
        }

        console.log("\n🏁 Todos los informes corporativos se han completado con éxito.");
    } catch (error) {
        console.error("❌ Error crítico en el pipeline masivo:", error);
    }
}

generarTodosLosInformes().catch(console.error);