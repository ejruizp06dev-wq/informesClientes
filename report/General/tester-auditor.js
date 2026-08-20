require('dotenv').config(); // 1. CARGAR VARIABLES DE ENTORNO AL INICIO
const fetch = globalThis.fetch || require('node-fetch');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { performance } = require('perf_hooks'); // Para medir el tiempo exacto de ejecución del script

// ===================================================
// CONFIGURACIÓN DINÁMICA DE FECHAS (DE HOY AL PASADO)
// ===================================================
const diasHaciaAtras = 7; // <- Cambia este número para definir cuántos días hacia atrás auditar
const hoy = new Date();

const fechaInicioDinamica = new Date();
fechaInicioDinamica.setDate(hoy.getDate() - diasHaciaAtras);
fechaInicioDinamica.setHours(0, 0, 0, 0); // Desde las 12:00 AM del día pasado

const fechaFinDinamica = new Date();
fechaFinDinamica.setHours(23, 59, 59, 999); // Hasta las 11:59 PM de hoy

// ===================================================
// ARQUITECTURA DE CONFIGURACIÓN DINÁMICA Y RUTAS
// ===================================================
const homeDir = os.homedir(); // Resuelve dinámicamente C:\Users\Usuario

const CONFIG = {
    baseUrl: process.env.DEVOPS_BASE_URL, // Leído desde tu .env local de forma segura
    pat: process.env.DEVOPS_PAT,         // Leído desde tu .env local de forma segura
    apiVersion: '6.0',
    
    outputHtml: 'tester.html',
    outputCsv: 'tester.csv',
   
    // RUTAS CENTRALIZADAS (Nueva Arquitectura)
    outputFolderReports: path.join(process.cwd(), 'report', 'tables'),
    outputFolderDashboardUI: path.join(process.cwd(), 'dashboard-ui', 'public', 'data'),
    outputFolderDashboardSrc: path.join(process.cwd(), 'dashboard-ui', 'src', 'data'),
    
    // RUTA DE SHAREPOINT REAL (Mapeada según tu explorador de Windows para Kiai)
    outputFolderSharePoint: path.join(
        homeDir, 
        'Kiai', 
        'Gestion - Documentos', 
        'Reportes', 
        'Todos los documentos'
    ),

    // Rango de fechas auto-calculado dinámicamente
    fechaInicio: fechaInicioDinamica,
    fechaFin: fechaFinDinamica,

    // Permitido ejecutar todos los días (Lunes a Domingo = 0-6)
    diasPermitidos: [1, 2, 3, 4, 5, 6, 0]
};

// IDs de campos de horas personalizados en Lappiz
const CAMPOS_HORAS = {
    estimacion: 'Custom.3e278da6-593c-4877-87bc-5147090fb8da',
    ejecucion: 'Custom.551fce5a-bd20-4077-bab3-75a3da915c29'
};

const authHeader = "Basic " + Buffer.from(":" + CONFIG.pat).toString("base64");

function validarCampo(valor) {
    const poblado = valor !== undefined && valor !== null && valor.toString().trim() !== "";
    return {
        texto: poblado ? valor.toString() : "0",
        clase: poblado ? "complete" : "incomplete"
    };
}

function limpiarParaCsv(texto) {
    if (!texto) return "";
    let cadena = texto.toString().replace(/"/g, '""');
    return `"${cadena}"`;
}

// Validación de restricción de días de la semana
const diaSemanaHoy = hoy.getDay();
const distribucionDiasPermitidos = CONFIG.diasPermitidos.includes(diaSemanaHoy);

if (!distribucionDiasPermitidos) {
    console.error(`\n\x1b[31m[EJECUCIÓN BLOQUEADA]\x1b[0m La ejecución está fuera del horario permitido.\n`);
    process.exit(0);
}

async function generarInformeGlobalTesters() {
    const tiempoInicioScript = performance.now();

    console.clear();
    console.log("==================================================");
    console.log(`[QA AUDIT MULTI-PROYECTO] Inicializando Pipeline`);
    console.log(`Rango dinámico: ${CONFIG.fechaInicio.toLocaleDateString()} al ${CONFIG.fechaFin.toLocaleDateString()} (Últimos ${diasHaciaAtras} días)`);
    console.log("==================================================");

    try {
        // 1. OBTENER TODOS LOS PROYECTOS DE LA COLECCIÓN
        console.log(`\n[1/4] Consultando la lista de proyectos en la colección...`);
        const projectsUrl = `${CONFIG.baseUrl}/_apis/projects?api-version=${CONFIG.apiVersion}`;
        const projectsResponse = await fetch(projectsUrl, {
            method: 'GET',
            headers: { 'Authorization': authHeader, 'Accept': 'application/json' }
        });

        if (!projectsResponse.ok) {
            throw new Error(`Error consultando proyectos de la colección: Código ${projectsResponse.status}`);
        }

        const projectsData = await projectsResponse.json();
        const listaProyectos = projectsData.value || [];
        console.log(`✔ Se detectaron ${listaProyectos.length} proyectos en la colección.`);

        const filasTestsHtml = [];
        const filasTestsCsv = [];
       
        filasTestsCsv.push([
            "Proyecto ID", "Proyecto Nombre", "ID Artifact", "Tipo Elemento",
            "Titulo del Escenario / Plan de Prueba", "Tester Asignado", "Estado Actual",
            "Sprint Origen", "Area Path", "Prioridad", "Automation Status",
            "Estructura de Pasos", "Hrs Estimadas", "Hrs Ejecutadas", "Duracion Test (ms)", "Ultima Modificacion"
        ].map(limpiarParaCsv).join(","));

        let totalProyectosProcesados = 0;
        let totalItemsExtraidos = 0;

        // 2. ITERAR EN CADA PROYECTO DETECTADO
        for (const proyecto of listaProyectos) {
            console.log(`\n➔ Procesando Proyecto: [${proyecto.name}]...`);
            try {
                const wiqlUrl = `${CONFIG.baseUrl}/${proyecto.id}/_apis/wit/wiql?api-version=${CONFIG.apiVersion}`;
                const wiqlBody = {
                    query: `Select [System.Id] From WorkItems Where [System.TeamProject] = '${proyecto.name}' AND ([System.WorkItemType] = 'Test Case' OR [System.WorkItemType] = 'Caso de Prueba' OR [System.WorkItemType] = 'Test Plan' OR [System.WorkItemType] = 'Test Suite')`
                };

                const wiqlResponse = await fetch(wiqlUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Authorization': authHeader },
                    body: JSON.stringify(wiqlBody)
                });

                const wiqlData = await wiqlResponse.json();
                if (!wiqlData.workItems || wiqlData.workItems.length === 0) {
                    totalProyectosProcesados++;
                    continue;
                }

                const todosLosIds = wiqlData.workItems.map(item => item.id);
                const listaCompletaValue = [];
                const tamañoLote = 150;

                for (let i = 0; i < todosLosIds.length; i += tamañoLote) {
                    const loteIds = todosLosIds.slice(i, i + tamañoLote).join(',');
                    const detailUrl = `${CONFIG.baseUrl}/_apis/wit/workitems?ids=${loteIds}&$expand=all&api-version=${CONFIG.apiVersion}`;
                   
                    const detailResponse = await fetch(detailUrl, {
                        method: 'GET',
                        headers: { 'Authorization': authHeader, 'Accept': 'application/json' }
                    });

                    const textResponse = await detailResponse.text();
                    if (!detailResponse.ok || textResponse.trim().startsWith("<!DOCTYPE")) continue;

                    const chunkData = JSON.parse(textResponse);
                    if (chunkData.value) listaCompletaValue.push(...chunkData.value);
                }

                let itemsFiltradosEnProyecto = 0;

                // 3. FILTRADO Y EXTRACCIÓN MÉTRICA
                listaCompletaValue.forEach(item => {
                    const fields = item.fields;
                    const fechaItemRaw = fields["System.ChangedDate"] || fields["System.CreatedDate"];
                    if (!fechaItemRaw) return;

                    const fechaItem = new Date(fechaItemRaw);
                    if (fechaItem < CONFIG.fechaInicio || fechaItem > CONFIG.fechaFin) return;

                    itemsFiltradosEnProyecto++;
                    totalItemsExtraidos++;

                    const tipo = fields["System.WorkItemType"];
                    const ultimaEdicion = fechaItem.toLocaleString('es-CO', { timeZone: 'America/Bogota' });
                    const testerAsignado = fields["System.AssignedTo"] ? fields["System.AssignedTo"].displayName : "SIN ASIGNAR";
                    const estado = fields["System.State"] || "Sin Estado";
                    const sprintAsociado = fields["System.IterationPath"] ? fields["System.IterationPath"].split('\\').pop() : "Sin Sprint";
                   
                    const prioridad = validarCampo(fields["Microsoft.VSTS.Common.Priority"]);
                    const automationStatus = validarCampo(fields["Microsoft.VSTS.TCM.AutomationStatus"]);
                    const areaPath = fields["System.AreaPath"] || "Raíz";
                   
                    const pasos = fields["Microsoft.VSTS.TCM.Steps"];
                    const tienePasos = pasos !== undefined && pasos !== null && pasos.toString().trim() !== "";
                    const textoPasos = tienePasos ? '✔ ESTRUCTURADOS' : '❌ SIN PASOS';

                    const hrsEstimadas = parseFloat(fields[CAMPOS_HORAS.estimacion]) || 0;
                    const hrsEjecutadas = parseFloat(fields[CAMPOS_HORAS.ejecucion]) || 0;

                    const activacionPrueba = fields["Microsoft.VSTS.Common.StateChangeDate"] ? new Date(fields["Microsoft.VSTS.Common.StateChangeDate"]) : null;
                    const duracionTestMs = activacionPrueba ? Math.max(0, fechaItem.getTime() - activacionPrueba.getTime()) : 0;
                    const duracionTexto = duracionTestMs > 0 ? `${(duracionTestMs / 1000).toFixed(1)}s` : "N/A";

                    filasTestsHtml.push(`
                        <tr>
                            <td style="background-color: #fcfbfe; font-weight: bold; color: #5c2d91;">${proyecto.name}</td>
                            <td><strong>#${item.id}</strong></td>
                            <td><span class="badge badge-type">${tipo}</span></td>
                            <td style="text-align: left; font-weight: 500;">${fields["System.Title"] || "Sin título"}</td>
                            <td><span class="badge badge-tester">${testerAsignado}</span></td>
                            <td><span class="badge badge-state">${estado}</span></td>
                            <td><span class="badge badge-sprint">${sprintAsociado}</span></td>
                            <td class="${prioridad.clase}">${prioridad.texto}</td>
                            <td class="${automationStatus.clase}">${automationStatus.texto}</td>
                            <td class="${tienePasos ? 'complete' : 'incomplete'}">${textoPasos}</td>
                            <td style="font-weight: bold; color: #106ebe;">${hrsEstimadas.toFixed(1)}h</td>
                            <td style="font-weight: bold; color: #107c41;">${hrsEjecutadas.toFixed(1)}h</td>
                            <td style="color: #666; font-family: monospace;">${duracionTexto}</td>
                            <td style="font-family: monospace; font-size: 11px; color: #444;">${ultimaEdicion}</td>
                        </tr>
                    `);

                    filasTestsCsv.push([
                        proyecto.id, proyecto.name, item.id, tipo, fields["System.Title"] || "Sin título",
                        testerAsignado, estado, sprintAsociado, areaPath, prioridad.texto,
                        automationStatus.texto, textoPasos, hrsEstimadas, hrsEjecutadas, duracionTestMs, ultimaEdicion
                    ].map(limpiarParaCsv).join(","));
                });

                totalProyectosProcesados++;

            } catch (projectError) {
                console.error(`   ❌ Error procesando el proyecto ${proyecto.name}:`, projectError.message);
            }
        }

        const tiempoFinScript = performance.now();
        const tiempoEjecuciónScriptSegundos = ((tiempoFinScript - tiempoInicioScript) / 1000).toFixed(2);

        // 5. CONSTRUCCIÓN HTML
        const htmlContenido = `
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>Auditoría de Testing Multi-Proyecto - Dashboard Metric</title>
            <style>
                body { font-family: 'Segoe UI', system-ui, sans-serif; background-color: #f7f5fa; color: #222; margin: 25px; }
                .container { background-color: #fff; padding: 35px; border-radius: 10px; box-shadow: 0 4px 12px rgba(92, 45, 145, 0.08); }
                h1 { color: #5c2d91; font-size: 24px; margin-bottom: 5px; }
                .meta-box { background: #f1ecf7; padding: 10px 15px; border-left: 4px solid #5c2d91; border-radius: 0 6px 6px 0; font-size: 13px; margin-bottom: 20px; }
                table { width: 100%; border-collapse: collapse; margin-top: 15px; font-size: 12px; }
                th { background-color: #5c2d91; color: white; padding: 12px 10px; font-weight: 600; text-transform: uppercase; font-size: 11px; letter-spacing: 0.5px; }
                td { padding: 10px; border-bottom: 1px solid #e8e3f0; text-align: center; }
                tr:hover { background-color: #fcfbfe; }
                .badge { padding: 4px 8px; border-radius: 4px; font-weight: bold; font-size: 11px; }
                .badge-tester { background-color: #e8f4fd; color: #106ebe; }
                .badge-state { background-color: #fff4ce; color: #794500; }
                .badge-type { background-color: #f3f2f1; color: #323130; border: 1px solid #edebe9; }
                .complete { background-color: #dff6dd; color: #107c41; font-weight: 500; }
                .incomplete { background-color: #fde7e9; color: #a80000; font-weight: bold; }
            </style>
        </head>
        <body>
            <div class="container">
                <h1>📋 Reporte Consolidado de Cobertura y Tiempos de Testeo</h1>
                <div class="meta-box">
                    Rango de Auditoría Dinámico: <strong>${CONFIG.fechaInicio.toLocaleDateString()} al ${CONFIG.fechaFin.toLocaleDateString()}</strong> &nbsp;|&nbsp;
                    Tiempo total de ejecución del script: <strong style="color: #107c41;">${tiempoEjecuciónScriptSegundos} segundos</strong>
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>Proyecto DevOps</th><th>ID Artifact</th><th>Tipo Elemento</th><th>Título del Escenario</th><th>Tester Asignado</th>
                            <th>Estado Actual</th><th>Sprint</th><th>Prioridad</th><th>Automation</th><th>Pasos</th>
                            <th>Hrs Estimadas</th><th>Hrs Ejecutadas</th><th>Duración Test</th><th>Última Modificación</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${filasTestsHtml.length > 0 ? filasTestsHtml.join('') : '<tr><td colspan="14">No se encontraron artefactos de prueba ejecutados en este período.</td></tr>'}
                    </tbody>
                </table>
            </div>
        </body>
        </html>
        `;

        // ===================================================
        // ESCRITURA MÚLTIPLE DE ARCHIVOS (LOCAL, DASHBOARD UI Y SHAREPOINT)
        // ===================================================
        
        // 1. Escritura en report/tables
        try {
            if (!fs.existsSync(CONFIG.outputFolderReports)) {
                fs.mkdirSync(CONFIG.outputFolderReports, { recursive: true });
            }
            fs.writeFileSync(path.join(CONFIG.outputFolderReports, CONFIG.outputHtml), htmlContenido, 'utf-8');
            fs.writeFileSync(path.join(CONFIG.outputFolderReports, CONFIG.outputCsv), '\ufeff' + filasTestsCsv.join("\n"), 'utf-8');
            console.log(`✔ Archivos de pruebas guardados en report/tables.`);
        } catch (errLocal) {
            console.error("❌ Error al guardar en report/tables:", errLocal.message);
        }

        // 2. Escritura en dashboard-ui/public/data
        try {
            if (!fs.existsSync(CONFIG.outputFolderDashboardUI)) {
                fs.mkdirSync(CONFIG.outputFolderDashboardUI, { recursive: true });
            }
            fs.writeFileSync(path.join(CONFIG.outputFolderDashboardUI, CONFIG.outputHtml), htmlContenido, 'utf-8');
            fs.writeFileSync(path.join(CONFIG.outputFolderDashboardUI, CONFIG.outputCsv), '\ufeff' + filasTestsCsv.join("\n"), 'utf-8');
            console.log(`✔ Archivos de pruebas guardados en dashboard-ui/public/data.`);
        } catch (errDashboard) {
            console.error("❌ Error al guardar en dashboard-ui/public/data:", errDashboard.message);
        }

        // 3. Escritura en dashboard-ui/src/data
        try {
            if (!fs.existsSync(CONFIG.outputFolderDashboardSrc)) {
                fs.mkdirSync(CONFIG.outputFolderDashboardSrc, { recursive: true });
            }
            fs.writeFileSync(path.join(CONFIG.outputFolderDashboardSrc, CONFIG.outputHtml), htmlContenido, 'utf-8');
            fs.writeFileSync(path.join(CONFIG.outputFolderDashboardSrc, CONFIG.outputCsv), '\ufeff' + filasTestsCsv.join("\n"), 'utf-8');
            console.log(`✔ Archivos de pruebas guardados en dashboard-ui/src/data.`);
        } catch (errSrc) {
            console.error("❌ Error al guardar en dashboard-ui/src/data:", errSrc.message);
        }

        // 4. Escritura Directa en SharePoint (Kiai)
        try {
            if (!fs.existsSync(CONFIG.outputFolderSharePoint)) {
                fs.mkdirSync(CONFIG.outputFolderSharePoint, { recursive: true });
            }
            fs.writeFileSync(path.join(CONFIG.outputFolderSharePoint, CONFIG.outputHtml), htmlContenido, 'utf-8');
            fs.writeFileSync(path.join(CONFIG.outputFolderSharePoint, CONFIG.outputCsv), '\ufeff' + filasTestsCsv.join("\n"), 'utf-8');
            console.log(`✔ Archivos de pruebas guardados y sincronizados en SharePoint con éxito.\n`);
        } catch (errNube) {
            console.log("⚠ Error de guardado directo en SharePoint:", errNube.message);
        }

    } catch (error) {
        console.error("\n\x1b[31m[ERROR CRÍTICO INESPERADO]\x1b[0m", error.message);
    }
}

generarInformeGlobalTesters();