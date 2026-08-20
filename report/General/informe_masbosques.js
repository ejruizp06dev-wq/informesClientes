require('dotenv').config();
const fetch = globalThis.fetch || require('node-fetch');
const fs = require('fs');
const path = require('path');
const { PDFDocument } = require('pdf-lib');
const puppeteer = require('puppeteer');
const {
    obtenerRangoMasBosquesDesdeEnv,
    resolverSprintSemanalMasBosques
} = require('./masbosques.logic');

// ===================================================
// LÓGICA ESPECÍFICA DE MAS BOSQUES (TAREAS EN HU Y BUGS)
// ===================================================
function normalizarTexto(valor) {
    return String(valor || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
}

function esProyectoMasBosques(nombreProyecto) {
    const nombre = normalizarTexto(nombreProyecto);
    return nombre.includes('masbosques') || nombre.includes('mas bosques') || nombre === 'mas_bosques';
}

function normalizarSprint(valor) {
    return normalizarTexto(valor).replace(/\s+/g, ' ');
}

function extraerAreaSprintMasBosques(sprintNombre) {
    const match = String(sprintNombre || '').match(/\(([^)]+)\)/);
    return match ? match[1].trim() : 'Fase IV';
}

function extraerEtiquetaArea(tagsField) {
    if (!tagsField) return 'N/A';
    const tags = String(tagsField).split(';').map(t => t.trim()).filter(Boolean);
    if (tags.length === 0) {
        const tagsComa = String(tagsField).split(',').map(t => t.trim()).filter(Boolean);
        if (tagsComa.length > 0) tags.push(...tagsComa);
    }
    const etiquetaEncontrada = tags.find(t => /^aj-?/i.test(t) || /^[a-z]{1,3}-?[a-z0-9]+$/i.test(t));
    return etiquetaEncontrada ? etiquetaEncontrada.toUpperCase() : (tags[0] ? tags[0].toUpperCase() : 'N/A');
}

const CONFIG = {
    baseUrl: process.env.DEVOPS_BASE_URL,
    pat: process.env.DEVOPS_PAT,
    apiVersion: '6.0',
    // RUTAS CENTRALIZADAS (Nueva Arquitectura)
    outputFolder: path.join(process.cwd(), 'report', 'tables'),
    outputFolderDashboardUI: path.join(process.cwd(), 'dashboard-ui', 'public', 'data'),
    outputFolderDashboardSrc: path.join(process.cwd(), 'dashboard-ui', 'src', 'data')
};

const CAMPOS_HORAS = {
    estimacionCustom: 'Custom.3e278da6-593c-4877-87bc-5147090fb8da',
    ejecucionCustom: 'Custom.551fce5a-bd20-4077-bab3-75a3da915c29',
    esfuerzo: 'Microsoft.VSTS.Scheduling.Effort',
    remainingWork: 'Microsoft.VSTS.Scheduling.RemainingWork'
};

const authHeader = "Basic " + Buffer.from(":" + CONFIG.pat).toString("base64");
const RUTA_MEMBRETE_PDF = path.join(process.cwd(), 'assets', 'formato.pdf');
const HORAS_ESTIMADAS_QUEMADAS_MASBOSQUES = 88;

function escapeHtml(valor = '') {
    return String(valor)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function obtenerResumenMasBosques() {
    const rutaResumen = path.join(process.cwd(), 'Resumen_dev', 'Informe_MasBosques_Lappiz.txt');
    if (!fs.existsSync(rutaResumen)) return '';
    return fs.readFileSync(rutaResumen, 'utf-8').trim();
}

function renderResumenMasBosquesHtml(resumenTexto = '') {
    const texto = String(resumenTexto || '').trim();
    if (!texto) return '';

    return `
        <div class="summary-box">
            <h2>Resumen</h2>
            <p>${escapeHtml(texto)}</p>
        </div>
    `;
}

async function aplicarMembreteSobrePdf(rutaPdfSalida) {
    if (!fs.existsSync(RUTA_MEMBRETE_PDF)) return false;
    try {
        const informeBytes = fs.readFileSync(rutaPdfSalida);
        const membreteBytes = fs.readFileSync(RUTA_MEMBRETE_PDF);
        const informeDoc = await PDFDocument.load(informeBytes);
        const membreteDoc = await PDFDocument.load(membreteBytes);
        const salidaDoc = await PDFDocument.create();

        if (membreteDoc.getPageCount() === 0) return false;

        const membretePagina = membreteDoc.getPage(0);
        const [membreteEmbebido] = await salidaDoc.embedPages([membretePagina]);

        for (let index = 0; index < informeDoc.getPageCount(); index += 1) {
            const paginaOriginal = informeDoc.getPage(index);
            const { width, height } = paginaOriginal.getSize();
            const paginaSalida = salidaDoc.addPage([width, height]);
            paginaSalida.drawPage(membreteEmbebido, { x: 0, y: 0, width, height });

            const [contenidoEmbebido] = await salidaDoc.embedPages([paginaOriginal]);
            paginaSalida.drawPage(contenidoEmbebido, {
                x: 22,
                y: 58,
                width: Math.max(1, width - 44),
                height: Math.max(1, height - 130)
            });
        }

        fs.writeFileSync(rutaPdfSalida, await salidaDoc.save());
        return true;
    } catch {
        return false;
    }
}

function obtenerSprintLimpioReporte(sprint) {
    const valor = String(sprint || '').trim();
    if (!valor) return 'Sin Iteracion';
    return valor.includes('\\') ? valor.split('\\').pop() : valor;
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

function contarComentarios(item) {
    if (!item.relations) return 0;
    return item.relations.filter(r => r.rel === "System.LinkTypes.Comment-Forward").length;
}

function contarArchivosAdjuntos(item) {
    if (!item.relations) return 0;
    return item.relations.filter(r => r.rel === "AttachedFile").length;
}

function tieneEtiquetas(tagsField) {
    if (!tagsField) return false;
    const tags = String(tagsField).split(';').map(t => t.trim()).filter(Boolean);
    if (tags.length === 0) {
        const tagsComa = String(tagsField).split(',').map(t => t.trim()).filter(Boolean);
        return tagsComa.length > 0;
    }
    return tags.length > 0;
}

function obtenerEtiquetasLista(tagsField) {
    if (!tagsField) return [];
    let tags = String(tagsField).split(';').map(t => t.trim()).filter(Boolean);
    if (tags.length === 0) {
        tags = String(tagsField).split(',').map(t => t.trim()).filter(Boolean);
    }
    return tags;
}

function parsearFechaLocal(valor) {
    if (!valor) return null;
    const fecha = new Date(valor);
    return Number.isNaN(fecha.getTime()) ? null : fecha;
}

function formatearFechaCorta(fecha) {
    if (!(fecha instanceof Date) || Number.isNaN(fecha.getTime())) return 'Sin fecha';
    return fecha.toLocaleDateString('es-CO', { day: '2-digit', month: '2-digit', year: 'numeric' });
}

// ===================================================
// GENERADOR EXCLUSIVO PARA MAS BOSQUES
// ===================================================
async function generarInformeMasBosques() {
    if (!fs.existsSync(CONFIG.outputFolder)) {
        fs.mkdirSync(CONFIG.outputFolder, { recursive: true });
    }

    const fechaInicioBase = parsearFechaLocal(process.env.FECHA_INICIO_ESPECIFICO || process.env.FECHA_INICIO) || new Date(new Date().getFullYear(), new Date().getMonth(), 1, 0, 0, 0, 0);
    const fechaFinBase = parsearFechaLocal(process.env.FECHA_FIN_ESPECIFICO || process.env.FECHA_FIN) || new Date();
    const rangoReporte = obtenerRangoMasBosquesDesdeEnv(fechaInicioBase, fechaFinBase);

    console.log("==================================================");
    console.log(`[MAS BOSQUES] Extrayendo elementos del Sprint 5 (Fase IV)...`);
    console.log("==================================================");

    try {
        const proyectosUrl = `${CONFIG.baseUrl}/_apis/projects?api-version=${CONFIG.apiVersion}`;
        const proyectosResponse = await fetch(proyectosUrl, { headers: { 'Authorization': authHeader } });
        if (!proyectosResponse.ok) throw new Error(`Error al listar proyectos en DevOps: ${proyectosResponse.statusText}`);

        const proyectosData = await proyectosResponse.json();
        const proyectoMasBosques = (proyectosData.value || []).find(p => esProyectoMasBosques(p.name));

        if (!proyectoMasBosques) {
            console.error("❌ No se encontró el proyecto 'Mas_Bosques' en Azure DevOps.");
            return;
        }

        const projectName = proyectoMasBosques.name;
        const projectId = proyectoMasBosques.id;

        const wiqlUrl = `${CONFIG.baseUrl}/${projectId}/_apis/wit/wiql?api-version=${CONFIG.apiVersion}`;
        const wiqlBody = { query: `Select [System.Id] From WorkItems Where [System.TeamProject] = '${projectName}'` };

        const wiqlResponse = await fetch(wiqlUrl, {
            method: 'POST',
            headers: { 'Content-Type': 'application/json', 'Authorization': authHeader },
            body: JSON.stringify(wiqlBody)
        });

        if (!wiqlResponse.ok) throw new Error(`Error en consulta WIQL para ${projectName}`);

        const wiqlData = await wiqlResponse.json();
        if (!wiqlData.workItems || wiqlData.workItems.length === 0) {
            console.warn("⚠️ No se encontraron elementos de trabajo para Mas Bosques.");
            return;
        }

        const ids = wiqlData.workItems.map(item => item.id);
        const itemsDetalle = await obtenerDetallesMasivos(ids);

        const itemsProcesados = itemsDetalle.map(item => {
            const tipo = (item.fields["System.WorkItemType"] || "").trim();
            const estado = item.fields["System.State"] || "Sin estado";
            if (normalizarTexto(estado) === 'removed') return null;

            let idPadre = null;
            if (item.relations) {
                const relPadre = item.relations.find(r => r.rel === "System.LinkTypes.Hierarchy-Reverse" || r.attributes?.name === "Parent");
                if (relPadre && relPadre.url) {
                    const partes = relPadre.url.split('/');
                    idPadre = parseInt(partes[partes.length - 1], 10);
                }
            }
            if (!idPadre && item.fields["System.Parent"]) {
                idPadre = parseInt(item.fields["System.Parent"], 10);
            }

            const estCustom = parseFloat(item.fields[CAMPOS_HORAS.estimacionCustom]) || 0;
            const esfuerzo = parseFloat(item.fields[CAMPOS_HORAS.esfuerzo]) || 0;
            const remaining = parseFloat(item.fields[CAMPOS_HORAS.remainingWork]) || 0;
            const horasEst = estCustom > 0 ? estCustom : (esfuerzo > 0 ? esfuerzo : remaining);
            const horasEjec = parseFloat(item.fields[CAMPOS_HORAS.ejecucionCustom]) || 0;

            const sprintPath = item.fields["System.IterationPath"] || "Sin Iteración";
            const cantComentarios = contarComentarios(item);
            const cantArchivos = contarArchivosAdjuntos(item);
            const tagsField = item.fields["System.Tags"];
            const tieneEtiquetasFlag = tieneEtiquetas(tagsField);
            const etiquetasList = obtenerEtiquetasLista(tagsField);

            return {
                id: item.id,
                huId: idPadre || item.id,
                padreId: idPadre,
                titulo: (item.fields["System.Title"] || "Sin título").replace(/[,;"'\n\r]/g, " ").trim(),
                tipo: tipo,
                estado: estado,
                sprint: sprintPath,
                sprintLimpio: obtenerSprintLimpioReporte(sprintPath),
                encargado: obtenerNombreEncargado(item.fields["System.AssignedTo"]),
                areaEtiqueta: extraerEtiquetaArea(item.fields["System.Tags"]),
                horasEstimadas: horasEst,
                horasEjecutadas: horasEjec,
                esSubtarea: tipo === "Task" || (tipo === "Bug" && idPadre !== null),
                fechaCreacionRaw: item.fields["System.CreatedDate"] || null,
                fechaCambioRaw: item.fields["System.ChangedDate"] || null,
                cantComentarios: cantComentarios,
                cantArchivos: cantArchivos,
                tieneEtiquetas: tieneEtiquetasFlag,
                etiquetas: etiquetasList
            };
        }).filter(Boolean);

        const sprintActivoNombre = resolverSprintSemanalMasBosques(
            itemsProcesados,
            rangoReporte.fechaInicio,
            rangoReporte.fechaFin,
            parsearFechaLocal,
            obtenerSprintLimpioReporte
        ) || 'Sin Iteracion';

        const itemsSprintActivo = itemsProcesados.filter(i => {
            const fechaCreacion = parsearFechaLocal(i.fechaCreacionRaw);
            const fechaCambio = parsearFechaLocal(i.fechaCambioRaw);
            const enRangoPorFecha = (fechaCreacion && fechaCreacion >= rangoReporte.fechaInicio && fechaCreacion <= rangoReporte.fechaFin)
                || (fechaCambio && fechaCambio >= rangoReporte.fechaInicio && fechaCambio <= rangoReporte.fechaFin);
            const enSprintObjetivo = normalizarSprint(i.sprintLimpio) === normalizarSprint(sprintActivoNombre);
            return enRangoPorFecha || enSprintObjetivo;
        });

        const itemsBacklogSprint = itemsSprintActivo.filter(i => !i.esSubtarea && i.tipo !== "Feature");
        // Consideramos tareas hijas tanto las de tipo 'Task' como los Bugs que actúan como subtareas de un elemento padre
        const tareasHijasSprint = itemsSprintActivo.filter(i => i.esSubtarea);

        const husPadres = itemsBacklogSprint.filter(i => i.tipo === "Product Backlog Item").map(hu => {
            const hijasAsociadas = tareasHijasSprint.filter(t => String(t.padreId) === String(hu.id));
            const cantTareasAsociadas = hijasAsociadas.length;
            const estAsociadas = hu.horasEstimadas > 0 ? hu.horasEstimadas : hijasAsociadas.reduce((acc, t) => acc + t.horasEstimadas, 0);
            const ejecAsociadas = hijasAsociadas.reduce((acc, t) => acc + t.horasEjecutadas, 0);
            const areaFinal = hu.areaEtiqueta !== 'N/A' ? hu.areaEtiqueta : 'Fase IV';

            return {
                ...hu,
                area: areaFinal,
                cantTareasAsociadas,
                horasEstimadasAsociadas: estAsociadas,
                horasEjecutadasAsociadas: ejecAsociadas
            };
        });

        const bugsPadres = itemsBacklogSprint.filter(i => i.tipo === "Bug").map(bug => {
            const hijasAsociadas = tareasHijasSprint.filter(t => String(t.padreId) === String(bug.id));
            const cantTareasAsociadas = hijasAsociadas.length;
            const estAsociadas = bug.horasEstimadas > 0 ? bug.horasEstimadas : hijasAsociadas.reduce((acc, t) => acc + t.horasEstimadas, 0);
            const ejecAsociadas = hijasAsociadas.reduce((acc, t) => acc + t.horasEjecutadas, 0);
            const areaFinal = bug.areaEtiqueta !== 'N/A' ? bug.areaEtiqueta : 'Fase IV';

            return {
                ...bug,
                area: areaFinal,
                cantTareasAsociadas,
                horasEstimadasAsociadas: estAsociadas,
                horasEjecutadasAsociadas: ejecAsociadas
            };
        });

        const totalEstCalculado = tareasHijasSprint.reduce((acc, t) => acc + t.horasEstimadas, 0);
        const totalEst = HORAS_ESTIMADAS_QUEMADAS_MASBOSQUES;
        const totalEjec = tareasHijasSprint.reduce((acc, t) => acc + t.horasEjecutadas, 0);
        const horasSoporteEjec = tareasHijasSprint.filter(t => normalizarTexto(t.titulo).includes('soporte')).reduce((acc, t) => acc + t.horasEjecutadas, 0);
        const horasSinSoporteEjec = Math.max(totalEjec - horasSoporteEjec, 0);
        const tareasConComentarios = [...husPadres, ...bugsPadres].filter(t => t.cantComentarios > 0).length;
        const tareasConArchivos = [...husPadres, ...bugsPadres].filter(t => t.cantArchivos > 0).length;
        const tareasConEtiquetas = [...husPadres, ...bugsPadres].filter(t => t.tieneEtiquetas).length;

        const areaSprintGeneral = extraerAreaSprintMasBosques(sprintActivoNombre);
        const descripcionAlcance = "Proyecto enfocado en la estabilización de código local, arquitectura de datos y soporte operativo LPZ.";

        // ===================================================
        // 📊 AUDITORÍA EN CONSOLA
        // ===================================================
        console.log("\n==================================================");
        console.log(`📋 [AUDITORÍA DEVOPS] SPRINT: ${sprintActivoNombre}`);
        console.log("==================================================");
        console.log(`🔹 Rango aplicado: ${formatearFechaCorta(rangoReporte.fechaInicio)} al ${formatearFechaCorta(rangoReporte.fechaFin)}`);
        console.log(`🔹 Cantidad de Historias de Usuario (HU): ${husPadres.length}`);
        console.log(`🔹 Cantidad de Bugs: ${bugsPadres.length}`);
        console.log(`🔹 Cantidad total de Tareas hijas: ${tareasHijasSprint.length}`);
        console.log(`🔹 Horas Estimadas Calculadas: ${totalEstCalculado.toFixed(1)}h`);
        console.log(`🔹 Horas Estimadas Quemadas: ${totalEst.toFixed(1)}h`);
        console.log(`🔹 Horas Estimadas Totales: ${totalEst.toFixed(1)}h`);
        console.log(`🔹 Horas Ejecutadas Totales: ${totalEjec.toFixed(1)}h`);
        console.log(`🔹 Tareas con Comentarios/Discusión: ${tareasConComentarios}`);
        console.log(`🔹 Tareas con Archivos Adjuntos: ${tareasConArchivos}`);
        console.log(`🔹 Tareas con Etiquetas: ${tareasConEtiquetas}\n`);

        console.log("--- DETALLE DE HISTORIAS DE USUARIO ---");
        console.table(husPadres.map(h => ({
            "ID": `#${h.id}`,
            "Título": h.titulo.substring(0, 30) + '...',
            "Área": h.area,
            "Estado": h.estado,
            "Cant. Tareas": h.cantTareasAsociadas,
            "Hrs Est.": `${h.horasEstimadasAsociadas.toFixed(1)}h`,
            "Hrs Ejec.": `${h.horasEjecutadasAsociadas.toFixed(1)}h`,
            "Comentarios": h.cantComentarios,
            "Archivos": h.cantArchivos,
            "Etiquetas": h.tieneEtiquetas ? 'Sí' : 'No'
        })));

        console.log("--- DETALLE DE BUGS ---");
        console.table(bugsPadres.map(b => ({
            "ID": `#${b.id}`,
            "Título": b.titulo.substring(0, 30) + '...',
            "Área": b.area,
            "Estado": b.estado,
            "Cant. Tareas": b.cantTareasAsociadas,
            "Hrs Est.": `${b.horasEstimadasAsociadas.toFixed(1)}h`,
            "Hrs Ejec.": `${b.horasEjecutadasAsociadas.toFixed(1)}h`,
            "Comentarios": b.cantComentarios,
            "Archivos": b.cantArchivos,
            "Etiquetas": b.tieneEtiquetas ? 'Sí' : 'No'
        })));
        console.log("==================================================\n");

        const scriptGraficos = `
            <script>
                Chart.register(ChartDataLabels);

                new Chart(document.getElementById('chartHoras').getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: ['Horas Tareas', 'Horas Soporte'],
                        datasets: [{
                            data: [${horasSinSoporteEjec.toFixed(1)}, ${horasSoporteEjec.toFixed(1)}],
                            backgroundColor: ['#1e88e5', '#f89c3c'],
                            borderRadius: 6
                        }]
                    },
                    options: { animation: false, responsive: true, maintainAspectRatio: false, plugins: { legend: { display: false } } }
                });

                new Chart(document.getElementById('chartDistribucion').getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: ['Tareas generales', 'Tareas soporte'],
                        datasets: [{
                            data: [${tareasHijasSprint.length}, ${tareasHijasSprint.filter(t => normalizarTexto(t.titulo).includes('soporte')).length}],
                            backgroundColor: ['#0d4a92', '#1e88e5']
                        }]
                    },
                    options: { animation: false, responsive: true, maintainAspectRatio: false, indexAxis: 'y', plugins: { legend: { display: false } } }
                });
            </script>
        `;

        const resumenHtml = renderResumenMasBosquesHtml(obtenerResumenMasBosques());

        const htmlContent = `
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>Reporte de Gestion - Mas_Bosques</title>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2"></script>
            <style>
                @page { size: A4; margin: 12mm; }
                body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; color: #45141a; background: #ffffff; line-height: 1.35; font-size: 9.5pt; margin: 0; }
                .page { padding: 8px; }
                .header { background: linear-gradient(135deg, #990816 0%, #7a0f1b 45%, #ce0a1e 100%); color: #ffffff; padding: 18px; border-radius: 16px; margin-bottom: 10px; }
                .header h1 { margin: 0; font-size: 18pt; }
                .summary-box { background: #fff5f6; border: 1px solid #f3c2ca; border-left: 4px solid #ce0a1e; padding: 10px 13px; margin-bottom: 10px; border-radius: 10px; }
                .summary-box h2 { margin: 0 0 6px 0; font-size: 10pt; color: #3d0b12; text-transform: uppercase; }
                .summary-box p { margin: 0; color: #5a1d25; font-size: 8.8pt; text-align: justify; }
                .description-box { background: rgba(248,252,255,0.95); border-left: 4px solid #ce0a1e; padding: 9px 13px; margin-bottom: 10px; border-radius: 0 12px 12px 0; }
                .kpi-row { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; margin-bottom: 10px; }
                .kpi-card { background: #ffffff; border: 1px solid #f3c2ca; padding: 10px; border-radius: 12px; text-align: center; }
                .kpi-card .number { font-size: 17pt; font-weight: 800; color: #3d0b12; }
                .kpi-card .label { font-size: 7.5pt; color: #315875; text-transform: uppercase; font-weight: 700; margin-top: 4px; }
                .state-kpi-row { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; margin-bottom: 10px; }
                .state-kpi-card { background: #ffffff; border: 1px solid #f3c2ca; border-radius: 12px; padding: 10px; text-align: center; }
                .state-kpi-card .number { font-size: 17pt; font-weight: 800; color: #3d0b12; }
                .state-kpi-card .label { font-size: 8pt; color: #244863; text-transform: uppercase; font-weight: 700; margin-top: 4px; }
                .chart-section { margin-bottom: 10px; }
                .charts-row { display: grid; gap: 10px; margin-bottom: 10px; grid-template-columns: repeat(2, minmax(0, 1fr)); }
                .chart-box { background: #fff8f9; border: 1px solid #f3c2ca; border-radius: 12px; padding: 10px; text-align: center; }
                .chart-box h3 { margin: 0 0 8px 0; font-size: 9pt; color: #3d0b12; }
                .chart-canvas-wrap { position: relative; width: 100%; height: 180px; }
                .chart-canvas-wrap canvas { width: 100% !important; height: 100% !important; }
                .section-title { font-size: 11pt; border-bottom: 2px solid #ce0a1e; padding-bottom: 4px; color: #3d0b12; margin: 14px 0 8px 0; font-weight: 800; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 8.5pt; }
                th { background-color: #6f0a16; color: white; padding: 7px 10px; text-align: left; }
                td { padding: 5px 9px; border-bottom: 1px solid #f5d6dc; color: #5a1d25; }
                tr:nth-child(even) { background-color: #fff2f4; }
                .badge-alert { background-color: #fef2f2; color: #991b1b; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 8pt; }
            </style>
        </head>
        <body>
            <div class="page">
                <div class="header">
                    <h1>Proyecto: Mas_Bosques</h1>
                    <div style="margin-top: 6px; font-size: 8.5pt; color: #ffe0e6; display: flex; justify-content: space-between;">
                        <span>Sprint: <strong>${sprintActivoNombre}</strong></span>
                        <span>Área Asignada: <strong>${areaSprintGeneral}</strong></span>
                    </div>
                </div>

                ${resumenHtml}

                <div class="description-box">
                    <strong>Descripción del alcance:</strong> ${descripcionAlcance}
                </div>

                <div class="kpi-row">
                    <div class="kpi-card">
                        <div class="number">${husPadres.length}</div>
                        <div class="label">Cantidad de HU</div>
                    </div>
                    <div class="kpi-card">
                        <div class="number" style="color: #b91c1c;">${bugsPadres.length}</div>
                        <div class="label">Cantidad de Bugs</div>
                    </div>
                    <div class="kpi-card">
                        <div class="number" style="color: #4a5568;">${tareasHijasSprint.length}</div>
                        <div class="label">Cantidad de Tareas</div>
                    </div>
                </div>

                <div class="state-kpi-row">
                    <div class="state-kpi-card">
                        <div class="number" style="color: #0f766e;">${totalEst.toFixed(1)}h</div>
                        <div class="label">Horas Estimadas</div>
                    </div>
                    <div class="state-kpi-card">
                        <div class="number" style="color: #16a34a;">${totalEjec.toFixed(1)}h</div>
                        <div class="label">Horas Ejecutadas</div>
                    </div>
                </div>

                <div class="state-kpi-row" style="grid-template-columns: repeat(3, minmax(0, 1fr));">
                    <div class="state-kpi-card">
                        <div class="number" style="color: #1e40af;">${tareasConComentarios}</div>
                        <div class="label">Con Comentarios</div>
                    </div>
                    <div class="state-kpi-card">
                        <div class="number" style="color: #7c3aed;">${tareasConArchivos}</div>
                        <div class="label">Con Archivos</div>
                    </div>
                    <div class="state-kpi-card">
                        <div class="number" style="color: #d97706;">${tareasConEtiquetas}</div>
                        <div class="label">Con Etiquetas</div>
                    </div>
                </div>

                <div class="chart-section">
                    <div class="charts-row">
                        <div class="chart-box">
                            <h3>Distribución de Tareas</h3>
                            <div class="chart-canvas-wrap"><canvas id="chartDistribucion"></canvas></div>
                        </div>
                        <div class="chart-box">
                            <h3>Horas por Tipo</h3>
                            <div class="chart-canvas-wrap"><canvas id="chartHoras"></canvas></div>
                        </div>
                    </div>
                </div>

                <div class="table-section">
                    <div class="section-title">1. Historias de Usuario Registradas (${sprintActivoNombre})</div>
                    ${husPadres.length > 0 ? `
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 60px;">HU ID</th>
                                <th>Título de Historia de Usuario</th>
                                <th style="width: 80px;">Área</th>
                                <th style="width: 80px;">Estado</th>
                                <th style="width: 60px; text-align: center;">Tareas</th>
                                <th style="width: 65px; text-align: right;">Est.</th>
                                <th style="width: 65px; text-align: right;">Ejec.</th>
                                <th style="width: 80px; text-align: center;">Gestión</th>
                                <th>Persona Encargada</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${husPadres.map(hu => `
                                <tr>
                                    <td><strong>#${hu.id}</strong></td>
                                    <td>${hu.titulo}</td>
                                    <td><span style="font-weight: 600; color: #0f766e;">${hu.area}</span></td>
                                    <td><span class="badge-alert">${hu.estado}</span></td>
                                    <td style="text-align: center; font-weight: bold;">${hu.cantTareasAsociadas}</td>
                                    <td style="text-align: right;">${hu.horasEstimadasAsociadas.toFixed(1)}h</td>
                                    <td style="text-align: right;">${hu.horasEjecutadasAsociadas.toFixed(1)}h</td>
                                    <td style="text-align: center; font-size: 8pt;">
                                        ${hu.cantComentarios > 0 ? `<span style="background: #e3f2fd; padding: 2px 4px; border-radius: 3px; margin-right: 2px;">💬 ${hu.cantComentarios}</span>` : ''}
                                        ${hu.cantArchivos > 0 ? `<span style="background: #f3e5f5; padding: 2px 4px; border-radius: 3px; margin-right: 2px;">📎 ${hu.cantArchivos}</span>` : ''}
                                        ${hu.tieneEtiquetas ? `<span style="background: #fff3e0; padding: 2px 4px; border-radius: 3px;">🏷️ Sí</span>` : '<span style="color: #999; font-size: 7.5pt;">—</span>'}
                                    </td>
                                    <td>${hu.encargado}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    ` : '<p style="color: #64748b; font-size: 8.5pt;">No se registraron Historias de Usuario en el sprint.</p>'}
                </div>

                <div class="table-section">
                    <div class="section-title">2. Bugs Registrados (${sprintActivoNombre})</div>
                    ${bugsPadres.length > 0 ? `
                    <table>
                        <thead>
                            <tr>
                                <th style="width: 60px;">Bug ID</th>
                                <th>Título del Bug</th>
                                <th style="width: 80px;">Área</th>
                                <th style="width: 80px;">Estado</th>
                                <th style="width: 60px; text-align: center;">Tareas</th>
                                <th style="width: 65px; text-align: right;">Est.</th>
                                <th style="width: 65px; text-align: right;">Ejec.</th>
                                <th style="width: 80px; text-align: center;">Gestión</th>
                                <th>Persona Encargada</th>
                            </tr>
                        </thead>
                        <tbody>
                            ${bugsPadres.map(bug => `
                                <tr>
                                    <td><strong>#${bug.id}</strong></td>
                                    <td>${bug.titulo}</td>
                                    <td><span style="font-weight: 600; color: #0f766e;">${bug.area}</span></td>
                                    <td><span class="badge-alert">${bug.estado}</span></td>
                                    <td style="text-align: center; font-weight: bold;">${bug.cantTareasAsociadas}</td>
                                    <td style="text-align: right;">${bug.horasEstimadasAsociadas.toFixed(1)}h</td>
                                    <td style="text-align: right;">${bug.horasEjecutadasAsociadas.toFixed(1)}h</td>
                                    <td style="text-align: center; font-size: 8pt;">
                                        ${bug.cantComentarios > 0 ? `<span style="background: #e3f2fd; padding: 2px 4px; border-radius: 3px; margin-right: 2px;">💬 ${bug.cantComentarios}</span>` : ''}
                                        ${bug.cantArchivos > 0 ? `<span style="background: #f3e5f5; padding: 2px 4px; border-radius: 3px; margin-right: 2px;">📎 ${bug.cantArchivos}</span>` : ''}
                                        ${bug.tieneEtiquetas ? `<span style="background: #fff3e0; padding: 2px 4px; border-radius: 3px;">🏷️ Sí</span>` : '<span style="color: #999; font-size: 7.5pt;">—</span>'}
                                    </td>
                                    <td>${bug.encargado}</td>
                                </tr>
                            `).join('')}
                        </tbody>
                    </table>
                    ` : '<p style="color: #64748b; font-size: 8.5pt;">No se registraron bugs en el sprint.</p>'}
                </div>

                ${scriptGraficos}
            </div>
        </body>
        </html>
        `;

        // Función auxiliar para guardar en múltiples ubicaciones
        const guardarEnMultiplesCarpetas = async (htmlContent, nombreBase) => {
            const carpetas = [
                CONFIG.outputFolder,
                CONFIG.outputFolderDashboardUI,
                CONFIG.outputFolderDashboardSrc
            ];

            for (const carpeta of carpetas) {
                if (!fs.existsSync(carpeta)) {
                    fs.mkdirSync(carpeta, { recursive: true });
                }

                const rutaHtml = path.join(carpeta, `${nombreBase}.html`);
                const rutaPdf = path.join(carpeta, `${nombreBase}.pdf`);

                fs.writeFileSync(rutaHtml, htmlContent, 'utf-8');

                const browser = await puppeteer.launch({ headless: 'new' });
                const page = await browser.newPage();
                await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
                await page.pdf({
                    path: rutaPdf,
                    format: 'A4',
                    printBackground: true,
                    margin: { top: '12mm', right: '12mm', bottom: '12mm', left: '12mm' }
                });
                await browser.close();

                await aplicarMembreteSobrePdf(rutaPdf);
                console.log(` ✅ Informe guardado en: ${rutaPdf}`);
            }
        };

        await guardarEnMultiplesCarpetas(htmlContent, 'Informe_Mas_Bosques');
        console.log(` ✅ Informe de Mas Bosques actualizado en todas las ubicaciones\n`);

    } catch (error) {
        console.error("❌ Error crítico generando informe de Mas Bosques:", error);
    }
}

async function obtenerDetallesMasivos(ids) {
    let result = [];
    const tamañoChunk = 150;
    for (let i = 0; i < ids.length; i += tamañoChunk) {
        const chunk = ids.slice(i, i + tamañoChunk).join(',');
        const detailUrl = `${CONFIG.baseUrl}/_apis/wit/workitems?ids=${chunk}&$expand=all&api-version=${CONFIG.apiVersion}`;
        const response = await fetch(detailUrl, { headers: { 'Authorization': authHeader, 'Accept': 'application/json' } });
        if (response.ok) {
            const data = await response.json();
            if (data.value) result = result.concat(data.value);
        }
    }
    return result;
}

// Ejecución directa
generarInformeMasBosques();