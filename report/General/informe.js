require('dotenv').config();
const fetch = globalThis.fetch || require('node-fetch');
const fs = require('fs');
const path = require('path');
const os = require('os');
const { PDFDocument } = require('pdf-lib');
const puppeteer = require('puppeteer'); // Módulo agregado para renderizar PDFs
const {
    esProyectoMasBosques,
    obtenerRangoMasBosquesDesdeEnv,
    obtenerFechaInicioHistorialMasBosques,
    normalizarSprint,
    resolverSprintSemanalMasBosques
} = require('./masbosques.logic');

// ===================================================
// DICCIONARIO DE DESCRIPCIONES POR PROYECTO
// ===================================================
const DESCRIPCIONES_PROYECTOS = {
    // "Nombre_Proyecto_DevOps": "Descripción del proyecto...",
};
const DESCRIPCION_ALCANCE_DEFAULT = "Consolidado y seguimiento de las funcionalidades, historias de usuario y requerimientos entregados por el cliente durante el periodo analizado.";

// ===================================================
// ARQUITECTURA DE CONFIGURACIÓN DINÁMICA / .ENV
// ===================================================
const homeDir = os.homedir(); 
const hoy = new Date();

let fechaInicio;
let fechaFin;
let esRangoPersonalizado = false;
let fechaInicioEspecifico;
let fechaFinEspecifico;
let esRangoEspecificoPersonalizado = false;

// Intentar leer las fechas desde el archivo .env
if (process.env.FECHA_INICIO && process.env.FECHA_FIN) {
    esRangoPersonalizado = true;
    
    // Parseo seguro evitando desfases de huso horario (UTC vs Local)
    const [anoI, mesI, diaI] = process.env.FECHA_INICIO.split('-').map(Number);
    fechaInicio = new Date(anoI, mesI - 1, diaI, 0, 0, 0, 0);

    const [anoF, mesF, diaF] = process.env.FECHA_FIN.split('-').map(Number);
    fechaFin = new Date(anoF, mesF - 1, diaF, 23, 59, 59, 999);
} else {
    // Fallback: rango mensual fijo solicitado (1 de junio de 2026 al 31 de julio de 2026)
    fechaInicio = new Date(2026, 5, 1, 0, 0, 0, 0);
    fechaFin = new Date(2026, 6, 31, 23, 59, 59, 999);
}

// Rango específico para secciones puntuales del PDF (controlado desde .env)
if (process.env.FECHA_INICIO_ESPECIFICO && process.env.FECHA_FIN_ESPECIFICO) {
    esRangoEspecificoPersonalizado = true;

    const [anoEI, mesEI, diaEI] = process.env.FECHA_INICIO_ESPECIFICO.split('-').map(Number);
    fechaInicioEspecifico = new Date(anoEI, mesEI - 1, diaEI, 0, 0, 0, 0);

    const [anoEF, mesEF, diaEF] = process.env.FECHA_FIN_ESPECIFICO.split('-').map(Number);
    fechaFinEspecifico = new Date(anoEF, mesEF - 1, diaEF, 23, 59, 59, 999);
} else {
    // Fallback: inicio de mes a hoy
    fechaInicioEspecifico = new Date(hoy.getFullYear(), hoy.getMonth(), 1, 0, 0, 0, 0);
    fechaFinEspecifico = new Date(hoy);
    fechaFinEspecifico.setHours(23, 59, 59, 999);
}

const CONFIG = {
    baseUrl: process.env.DEVOPS_BASE_URL, 
    pat: process.env.DEVOPS_PAT,
    apiVersion: '6.0',
    
    outputFile: 'dashboard.csv',
    outputHtmlFile: 'dashboard.html',
    
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

    ejecucion: {
        diasPermitidos: [1, 2, 3, 4, 5, 6, 0], // Permitidos todos los días (Lunes a Domingo = 0-6)
        fechaInicio: fechaInicio,
        fechaFin: fechaFin
    },
    ejecucionEspecifica: {
        fechaInicio: fechaInicioEspecifico,
        fechaFin: fechaFinEspecifico,
        esPersonalizado: esRangoEspecificoPersonalizado
    }
};

const CAMPOS_HORAS = {
    estimacion: 'Custom.3e278da6-593c-4877-87bc-5147090fb8da',
    ejecucion: 'Custom.551fce5a-bd20-4077-bab3-75a3da915c29'
};

const authHeader = "Basic " + Buffer.from(":" + CONFIG.pat).toString("base64");

const RUTA_ICONO_REPORTE = path.join(process.cwd(), 'assets', 'icon.jpg');
const ICONO_REPORTE_DATA_URI = fs.existsSync(RUTA_ICONO_REPORTE)
    ? `data:image/jpeg;base64,${fs.readFileSync(RUTA_ICONO_REPORTE).toString('base64')}`
    : '';
const RUTA_MEMBRETE_PDF = path.join(process.cwd(), 'assets', 'formato.pdf');

async function aplicarMembreteSobrePdf(rutaPdfSalida) {
    if (!fs.existsSync(RUTA_MEMBRETE_PDF)) {
        return false;
    }

    try {
        const informeBytes = fs.readFileSync(rutaPdfSalida);
        const membreteBytes = fs.readFileSync(RUTA_MEMBRETE_PDF);
        const informeDoc = await PDFDocument.load(informeBytes);
        const membreteDoc = await PDFDocument.load(membreteBytes);
        const salidaDoc = await PDFDocument.create();

        if (membreteDoc.getPageCount() === 0) {
            return false;
        }

        const membretePagina = membreteDoc.getPage(0);
        const [membreteEmbebido] = await salidaDoc.embedPages([membretePagina]);

        for (let index = 0; index < informeDoc.getPageCount(); index += 1) {
            const paginaOriginal = informeDoc.getPage(index);

            const { width, height } = paginaOriginal.getSize();
            const paginaSalida = salidaDoc.addPage([width, height]);
            paginaSalida.drawPage(membreteEmbebido, {
                x: 0,
                y: 0,
                width,
                height
            });

            const [contenidoEmbebido] = await salidaDoc.embedPages([paginaOriginal]);

            // Area segura para no tapar el logo/encabezado y pie del membrete.
            const margenIzquierdo = 22;
            const margenDerecho = 22;
            const margenSuperior = 72;
            const margenInferior = 58;
            const anchoContenido = Math.max(1, width - margenIzquierdo - margenDerecho);
            const altoContenido = Math.max(1, height - margenSuperior - margenInferior);

            paginaSalida.drawPage(contenidoEmbebido, {
                x: margenIzquierdo,
                y: margenInferior,
                width: anchoContenido,
                height: altoContenido
            });
        }

        const salidaBytes = await salidaDoc.save();
        fs.writeFileSync(rutaPdfSalida, salidaBytes);
        return true;
    } catch {
        return false;
    }
}

function parsearFechaReporte(valor) {
    if (!valor) return null;
    const fecha = new Date(valor);
    return Number.isNaN(fecha.getTime()) ? null : fecha;
}

function formatearFechaReporte(fecha) {
    if (!fecha) return 'Sin fecha';
    return fecha.toLocaleDateString('es-CO', { day: '2-digit', month: 'short', year: 'numeric' });
}

function obtenerFechaFilaReporte(fila) {
    return parsearFechaReporte(fila?.fechaRaw || fila?.ultimaModificacion);
}

function obtenerSprintLimpioReporte(sprint) {
    const valor = String(sprint || '').trim();
    if (!valor) return 'Sin Iteracion';
    return valor.includes('\\') ? valor.split('\\').pop() : valor;
}

function normalizarTextoReporte(valor) {
    return String(valor || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
}

function esBacklogSoporteReporte(titulo) {
    return normalizarTextoReporte(titulo).includes('soporte');
}

function esBacklogSesionesReporte(titulo) {
    const tituloNormalizado = normalizarTextoReporte(titulo)
        .replace(/[^a-z0-9\s-]/g, ' ')
        .replace(/\s+/g, ' ')
        .trim();

    return /^sesiones?(\b|\s|-|:)/.test(tituloNormalizado);
}

function esBacklogPasoProduccionReporte(titulo) {
    return normalizarTextoReporte(titulo).includes('paso a produccion');
}

function esBacklogEstimacionesReporte(titulo) {
    return normalizarTextoReporte(titulo).includes('estimacion');
}

function esEstadoRemovidoReporte(estado) {
    const estadoNormalizado = normalizarTextoReporte(estado);
    return estadoNormalizado === 'removed' || estadoNormalizado === 'remove';
}

function dividirEtiquetaEjeReporte(texto, maxLen = 16) {
    const limpio = String(texto || '').trim();
    if (!limpio) return ['Sin dato'];

    const palabras = limpio.split(/\s+/);
    const lineas = [];
    let actual = '';

    palabras.forEach(palabra => {
        const candidata = actual ? `${actual} ${palabra}` : palabra;
        if (candidata.length <= maxLen) {
            actual = candidata;
        } else {
            if (actual) lineas.push(actual);
            actual = palabra;
        }
    });

    if (actual) lineas.push(actual);
    return lineas.length ? lineas : [limpio];
}

function abreviarEtiquetaSprintReporte(texto, maxLen = 18) {
    const limpio = String(texto || '').trim();
    if (!limpio) return 'Sin dato';

    let abreviado = limpio
        .replace(/\(Ajustes Jurídica\)/gi, '(Aj. Jur.)')
        .replace(/\(Área jurídica\)/gi, '(Jur.)')
        .replace(/\(Área técnica\)/gi, '(Téc.)')
        .replace(/\(Fase III\)/gi, '(FIII)')
        .replace(/\(Fase IV\)/gi, '(FIV)');

    if (abreviado.length > maxLen) {
        abreviado = abreviado.slice(0, maxLen - 1).trimEnd() + '…';
    }

    return abreviado;
}

function esTipoTareaOperativaReporte(tipo) {
    const tipoNormalizado = normalizarTextoReporte(tipo);
    return tipoNormalizado === 'task' || tipoNormalizado === 'tarea' || tipoNormalizado.includes('task');
}

function contarItemsUnicosReporte(items = []) {
    const ids = new Set(
        items.map(item => String(item.tareaId || `${item.huId || ''}-${item.titulo || ''}`))
    );
    return ids.size;
}

function obtenerHorasEstimadasObjetivoPorMes(fechaReferencia) {
    const fecha = fechaReferencia instanceof Date ? fechaReferencia : new Date(fechaReferencia);
    const mes = fecha.getMonth() + 1;

    if (mes === 6) return 170;
    if (mes >= 7) return 160;
    return 160;
}

function obtenerNotaVisualProyecto(nombreProyecto) {
    const nombreLimpio = String(nombreProyecto || '').replace(/[^a-zA-Z0-9_-]/g, '_');
    const rutaNota = path.join(process.cwd(), 'notas_dev', `Informe_${nombreLimpio}.txt`);
    if (!fs.existsSync(rutaNota)) return '';

    return fs.readFileSync(rutaNota, 'utf-8').trim();
}

function escapeHtmlReporte(valor = '') {
    return String(valor)
        .replace(/&/g, '&amp;')
        .replace(/</g, '&lt;')
        .replace(/>/g, '&gt;')
        .replace(/"/g, '&quot;')
        .replace(/'/g, '&#39;');
}

function obtenerResumenProyecto(nombreProyecto) {
    const nombreLimpio = String(nombreProyecto || '').replace(/[^a-zA-Z0-9_-]/g, '_');
    const nombreCompacto = nombreLimpio.replace(/[_-]+/g, '');
    const candidatos = [
        `Resumen_${nombreLimpio}.txt`,
        `Informe_${nombreLimpio}.txt`,
        `Resumen_${nombreLimpio}_Lappiz.txt`,
        `Informe_${nombreLimpio}_Lappiz.txt`,
        `Resumen_${nombreCompacto}.txt`,
        `Informe_${nombreCompacto}.txt`,
        `Resumen_${nombreCompacto}_Lappiz.txt`,
        `Informe_${nombreCompacto}_Lappiz.txt`
    ];

    for (const archivo of candidatos) {
        const rutaResumen = path.join(process.cwd(), 'Resumen_dev', archivo);
        if (fs.existsSync(rutaResumen)) {
            return fs.readFileSync(rutaResumen, 'utf-8').trim();
        }
    }

    return '';
}

function renderResumenProyectoHtml(resumenTexto = '') {
    const bloques = String(resumenTexto)
        .split(/\r?\n\s*\r?\n/)
        .map(bloque => bloque.trim())
        .filter(Boolean);

    if (!bloques.length) return '';

    const parrafos = bloques
        .map(bloque => {
            const lineas = bloque
                .split(/\r?\n/)
                .map(linea => linea.trim())
                .filter(Boolean);

            const items = lineas
                .filter(linea => linea.startsWith('*'))
                .map(linea => linea.replace(/^\*+\s*/, '').trim())
                .filter(Boolean);

            if (items.length === lineas.length) {
                return `<ul class="summary-list">${items.map(item => `<li>${escapeHtmlReporte(item)}</li>`).join('')}</ul>`;
            }

            const texto = lineas
                .filter(linea => !linea.startsWith('*'))
                .join(' ')
                .trim();

            const partes = [];
            if (texto) {
                partes.push(`<p class="summary-paragraph">${escapeHtmlReporte(texto)}</p>`);
            }
            if (items.length) {
                partes.push(`<ul class="summary-list">${items.map(item => `<li>${escapeHtmlReporte(item)}</li>`).join('')}</ul>`);
            }

            return partes.join('');
        })
        .join('');

    return `
        <div class="summary-box">
            <strong>Resumen</strong>
            ${parrafos}
        </div>
    `;
}

function formatearNotaVisualProyectoHtml(notaTexto = '') {
    const lineas = String(notaTexto)
        .split(/\r?\n/)
        .map(linea => linea.trim())
        .filter(Boolean);

    if (!lineas.length) return '';

    return lineas.map(linea => `<div style="margin-bottom: 6px; line-height: 1.4;">${linea}</div>`).join('');
}

function obtenerClavesPeriodoReporte(fechaInicio, fechaFin) {
    const mesesEs = ['enero', 'febrero', 'marzo', 'abril', 'mayo', 'junio', 'julio', 'agosto', 'septiembre', 'octubre', 'noviembre', 'diciembre'];
    const mesesEn = ['january', 'february', 'march', 'april', 'may', 'june', 'july', 'august', 'september', 'october', 'november', 'december'];
    const cursor = new Date(fechaInicio.getFullYear(), fechaInicio.getMonth(), 1);
    const limite = new Date(fechaFin.getFullYear(), fechaFin.getMonth(), 1);
    const claves = new Set();

    while (cursor <= limite) {
        const year = cursor.getFullYear();
        const month = cursor.getMonth();
        claves.add(`${mesesEs[month]}${year}`);
        claves.add(`${mesesEn[month]}${year}`);
        claves.add(`${year}-${String(month + 1).padStart(2, '0')}`);
        cursor.setMonth(cursor.getMonth() + 1);
    }

    return [...claves];
}

function coincideSprintConPeriodoReporte(sprint, fechaInicio, fechaFin) {
    const sprintNormalizado = normalizarTextoReporte(sprint);
    if (!sprintNormalizado) return false;

    return obtenerClavesPeriodoReporte(fechaInicio, fechaFin)
        .some(clave => sprintNormalizado.includes(clave));
}

function cumpleRangoFechasOSprintEspecifico(item) {
    if (cumpleRangoFechas(item)) return true;

    const sprint = item?.fields?.["System.IterationPath"] || '';
    return coincideSprintConPeriodoReporte(
        sprint,
        CONFIG.ejecucionEspecifica.fechaInicio,
        CONFIG.ejecucionEspecifica.fechaFin
    );
}

function calcularComparativoPeriodo(items, fechaFinBase = new Date()) {
    const finActual = new Date(fechaFinBase);
    finActual.setHours(23, 59, 59, 999);

    const inicioActual = new Date(finActual);
    inicioActual.setDate(inicioActual.getDate() - 14);
    inicioActual.setHours(0, 0, 0, 0);

    const finPrevio = new Date(inicioActual);
    finPrevio.setMilliseconds(-1);

    const inicioPrevio = new Date(inicioActual);
    inicioPrevio.setDate(inicioPrevio.getDate() - 15);
    inicioPrevio.setHours(0, 0, 0, 0);

    const calcularTotales = filas => filas.reduce((acumulado, fila) => {
        const fechaFila = obtenerFechaFilaReporte(fila);
        if (!fechaFila) return acumulado;

        if (fila.esSubtarea) acumulado.totalSubtareas += 1;
        else acumulado.totalHus += 1;

        if (fila.tipo === 'Bug') acumulado.totalBugs += 1;
        if (!fila.esSubtarea && fila.tipo !== 'Bug') acumulado.totalTasks += 1;

        acumulado.horasEstimadas += parseFloat(fila.horasEstimadas) || 0;
        acumulado.horasEjecutadas += parseFloat(fila.horasEjecutadas) || 0;
        acumulado.items += 1;
        return acumulado;
    }, {
        items: 0,
        totalHus: 0,
        totalTasks: 0,
        totalBugs: 0,
        totalSubtareas: 0,
        horasEstimadas: 0,
        horasEjecutadas: 0
    });

    const periodoActual = items.filter(fila => {
        const fechaFila = obtenerFechaFilaReporte(fila);
        return fechaFila && fechaFila >= inicioActual && fechaFila <= finActual;
    });

    const periodoPrevio = items.filter(fila => {
        const fechaFila = obtenerFechaFilaReporte(fila);
        return fechaFila && fechaFila >= inicioPrevio && fechaFila <= finPrevio;
    });

    const actual = calcularTotales(periodoActual);
    const previo = calcularTotales(periodoPrevio);

    return {
        inicioActual,
        finActual,
        inicioPrevio,
        finPrevio,
        actual,
        previo,
        variacionItems: actual.items - previo.items,
        variacionHoras: actual.horasEjecutadas - previo.horasEjecutadas
    };
}

function resumirEvolutivoMensual(conteoMensual = {}) {
    const ordenMeses = {
        'Enero 2026': 1, 'Febrero 2026': 2, 'Marzo 2026': 3, 'Abril 2026': 4,
        'Mayo 2026': 5, 'Junio 2026': 6, 'Julio 2026': 7, 'Agosto 2026': 8,
        'Septiembre 2026': 9, 'Octubre 2026': 10, 'Noviembre 2026': 11, 'Diciembre 2026': 12
    };

    return Object.keys(conteoMensual)
        .sort((a, b) => (ordenMeses[a] || 99) - (ordenMeses[b] || 99))
        .map(mes => {
            const total = Object.values(conteoMensual[mes]).reduce((acumulado, datos) => {
                return acumulado + (datos.hus || 0) + (datos.tasks || 0) + (datos.bugs || 0);
            }, 0);

            return { mes, total };
        });
}

function construirPieReporte() {
    return `
        <footer class="report-footer">
            <div class="report-footer__line"></div>
            <div class="report-footer__content">
                ${ICONO_REPORTE_DATA_URI ? `<img src="${ICONO_REPORTE_DATA_URI}" alt="Icono del reporte" class="report-footer__icon">` : ''}
                <div>
                    <div class="report-footer__title">Reporte de seguimiento y evolutivo</div>
                    <div class="report-footer__subtitle">Generado automáticamente desde la base operativa de Azure DevOps</div>
                </div>
            </div>
        </footer>
    `;
}

function validarDiaEjecucion() {
    const diaSemana = hoy.getDay();
    return CONFIG.ejecucion.diasPermitidos.includes(diaSemana);
}

function cumpleRangoFechas(item) {
    const campos = item.fields || {};
    const fechaCreacion = campos["System.CreatedDate"] ? new Date(campos["System.CreatedDate"]) : null;
    const fechaCambio = campos["System.ChangedDate"] ? new Date(campos["System.ChangedDate"]) : null;
   
    const inicio = CONFIG.ejecucion.fechaInicio;
    const fin = CONFIG.ejecucion.fechaFin;

    return (fechaCreacion >= inicio && fechaCreacion <= fin) || (fechaCambio >= inicio && fechaCambio <= fin);
}

function obtenerNombreEncargado(assignedToField) {
    if (!assignedToField) return "Sin Asignar";
    if (typeof assignedToField === 'object' && assignedToField.displayName) {
        return assignedToField.displayName;
    }
    if (typeof assignedToField === 'string') {
        const match = assignedToField.match(/^([^<]+)/);
        return match ? match[1].trim() : assignedToField;
    }
    return "Sin Asignar";
}

function generarHTML(rows, metrizGlobales, contexto = {}) {
    const comparativo = calcularComparativoPeriodo(rows, contexto.fechaFin || CONFIG.ejecucion.fechaFin || hoy);
    const evolutivoMensual = resumirEvolutivoMensual(contexto.conteoMensual || {});
    const porcentajeGlobal = metrizGlobales.totalEst > 0 ? (metrizGlobales.totalEjec / metrizGlobales.totalEst) * 100 : 0;
    const etiquetaRango = esRangoPersonalizado ? 'Rango Estático .env' : 'Rango Evolutivo Base';

    const rowsOrdenadas = [...rows].sort((a, b) => {
        const fechaA = obtenerFechaFilaReporte(a)?.getTime() || 0;
        const fechaB = obtenerFechaFilaReporte(b)?.getTime() || 0;
        return fechaB - fechaA;
    });

    let tablaRows = '';
    rowsOrdenadas.forEach(r => {
        const esHijo = r.esSubtarea;
        const claseFila = esHijo ? 'fila-tarea' : 'fila-padre';
        const estiloTitulo = esHijo ? 'padding-left: 25px; color: #555;' : 'font-weight: 600;';
        const displayTasks = esHijo ? '-' : r.cantTareas;
        const displayBugs = esHijo ? '-' : r.cantBugs;

        tablaRows += `
            <tr class="${claseFila}">
                <td>${r.proyectoNombre}</td>
                <td>${r.sprint}</td>
                <td><strong>${r.huId}</strong></td>
                <td style="${estiloTitulo}">${esHijo ? '↳ ' : ''}${r.titulo}</td>
                <td><span class="badge-tipo ${r.tipo.toLowerCase().replace(/\s+/g, '-')}">${r.tipo}</span></td>
                <td><span class="badge-estado ${r.estado.toLowerCase()}">${r.estado}</span></td>
                <td>${r.encargado}</td>
                <td style="text-align: center; font-weight: bold; color: #4a5568;">${displayTasks}</td>
                <td style="text-align: center; font-weight: bold; color: #a80000;">${displayBugs}</td>
                <td style="text-align: right;">${r.horasEstimadas}</td>
                <td style="text-align: right;">${r.horasEjecutadas}</td>
                <td>
                    <div class="progress-bar-container">
                        <div class="progress-bar" style="width: ${r.porcentajeAvance}%"></div>
                        <span class="progress-text">${r.porcentajeAvance}%</span>
                    </div>
                </td>
                <td style="font-family: monospace; font-size: 11px; color: #555;">${r.ultimaModificacion}</td>
            </tr>
        `;
    });

    const mesesEvolutivo = evolutivoMensual.map(item => item.mes);
    const valoresEvolutivo = evolutivoMensual.map(item => item.total);
    const itemsPuntuales = rowsOrdenadas
        .filter(r => {
            const fechaFila = obtenerFechaFilaReporte(r);
            return fechaFila && fechaFila >= comparativo.inicioActual;
        })
        .slice(0, 30);

    const porcentajeVariacion = comparativo.previo.items > 0
        ? ((comparativo.actual.items - comparativo.previo.items) / comparativo.previo.items) * 100
        : 0;
    const porcentajeHoras = comparativo.previo.horasEjecutadas > 0
        ? ((comparativo.actual.horasEjecutadas - comparativo.previo.horasEjecutadas) / comparativo.previo.horasEjecutadas) * 100
        : 0;

    const tituloCobertura = `${formatearFechaReporte(contexto.fechaInicio || CONFIG.ejecucion.fechaInicio)} al ${formatearFechaReporte(contexto.fechaFin || CONFIG.ejecucion.fechaFin)}`;

    return `
    <!DOCTYPE html>
    <html lang="es">
    <head>
        <meta charset="UTF-8">
        <meta name="viewport" content="width=device-width, initial-scale=1.0">
        <title>Informe Evolutivo - Azure DevOps Lappiz</title>
        <style>
            :root {
                --bg: #f2efe8;
                --paper: #ffffff;
                --ink: #16212c;
                --muted: #64748b;
                --gold: #c58b2f;
                --teal: #0f766e;
                --slate: #e2e8f0;
            }
            * { box-sizing: border-box; }
            body {
                margin: 0;
                font-family: 'Segoe UI', Tahoma, Geneva, Verdana, sans-serif;
                color: var(--ink);
                background:
                    radial-gradient(circle at top left, rgba(197,139,47,0.14), transparent 30%),
                    radial-gradient(circle at top right, rgba(15,118,110,0.10), transparent 24%),
                    linear-gradient(180deg, #fbfaf7 0%, var(--bg) 100%);
            }
            .page { max-width: 1440px; margin: 0 auto; padding: 24px; }
            .hero {
                display: grid;
                grid-template-columns: 1.25fr 0.75fr;
                gap: 18px;
                padding: 28px;
                border-radius: 28px;
                color: #f8fafc;
                background: linear-gradient(135deg, #0f172a 0%, #1e293b 48%, #0f766e 100%);
                box-shadow: 0 18px 40px rgba(15,23,42,0.18);
                overflow: hidden;
                position: relative;
            }
            .hero::after {
                content: '';
                position: absolute;
                inset: auto -12% -40% auto;
                width: 280px;
                height: 280px;
                border-radius: 50%;
                background: rgba(255,255,255,0.08);
                filter: blur(4px);
            }
            .eyebrow {
                display: inline-block;
                padding: 7px 12px;
                border-radius: 999px;
                background: rgba(255,255,255,0.12);
                letter-spacing: 0.12em;
                font-size: 10px;
                text-transform: uppercase;
                font-weight: 700;
                margin-bottom: 12px;
            }
            h1 { margin: 0 0 12px 0; font-size: clamp(28px, 4vw, 46px); line-height: 1.02; }
            .hero p { margin: 0; color: rgba(248,250,252,0.86); max-width: 780px; font-size: 14px; }
            .hero-meta { display: flex; gap: 10px; flex-wrap: wrap; margin-top: 18px; }
            .pill { padding: 10px 14px; border-radius: 999px; background: rgba(255,255,255,0.1); font-size: 12px; font-weight: 700; }
            .hero-grid { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 12px; align-content: start; z-index: 1; }
            .hero-card {
                background: rgba(255,255,255,0.12);
                border: 1px solid rgba(255,255,255,0.14);
                border-radius: 18px;
                padding: 16px;
                backdrop-filter: blur(8px);
            }
            .hero-card .label { font-size: 11px; text-transform: uppercase; letter-spacing: 0.08em; color: rgba(255,255,255,0.72); }
            .hero-card .value { margin-top: 8px; font-size: 24px; font-weight: 800; }
            .hero-card .note { margin-top: 4px; font-size: 12px; color: rgba(255,255,255,0.72); }
            .section {
                margin-top: 18px;
                background: rgba(255,255,255,0.72);
                border: 1px solid rgba(226,232,240,0.9);
                border-radius: 24px;
                padding: 20px;
                box-shadow: 0 14px 32px rgba(15,23,42,0.06);
                backdrop-filter: blur(8px);
            }
            .section-header { display: flex; justify-content: space-between; gap: 12px; align-items: end; margin-bottom: 16px; }
            .section-header h2 { margin: 0; font-size: 20px; color: var(--ink); }
            .section-header .meta { color: var(--muted); font-size: 12px; }
            .kpi-row { display: grid; grid-template-columns: repeat(4, minmax(0, 1fr)); gap: 12px; }
            .kpi-card {
                background: var(--paper);
                border-radius: 18px;
                border: 1px solid rgba(226,232,240,0.95);
                padding: 16px;
            }
            .kpi-card .label { font-size: 11px; color: var(--muted); text-transform: uppercase; letter-spacing: 0.08em; }
            .kpi-card .number { margin-top: 8px; font-size: 28px; font-weight: 800; color: var(--ink); }
            .kpi-card .sub { margin-top: 4px; font-size: 12px; color: var(--muted); }
            .comparison-grid { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 12px; margin-top: 14px; }
            .comparison-card { background: var(--paper); border: 1px solid var(--slate); border-radius: 18px; padding: 16px; }
            .comparison-card strong { display: block; margin-bottom: 8px; font-size: 13px; color: var(--ink); }
            .comparison-card .line { display: flex; justify-content: space-between; gap: 10px; font-size: 12px; color: var(--muted); padding: 4px 0; }
            .comparison-card .delta { font-weight: 800; color: var(--gold); }
            .charts-row { display: grid; grid-template-columns: 1.1fr 0.9fr; gap: 12px; margin-top: 14px; }
            .chart-box { background: var(--paper); border: 1px solid var(--slate); border-radius: 18px; padding: 14px; }
            .chart-box h3 { margin: 0 0 10px 0; font-size: 13px; color: var(--ink); }
            .chart-canvas-wrap { position: relative; width: 100%; height: 230px; }
            .chart-canvas-wrap canvas { width: 100% !important; height: 100% !important; }
            table { width: 100%; border-collapse: collapse; background: var(--paper); border-radius: 18px; overflow: hidden; margin-top: 14px; border: 1px solid var(--slate); }
            th { background: linear-gradient(90deg, #0f172a, #1e293b); color: white; padding: 12px 14px; text-align: left; font-size: 13px; }
            td { padding: 11px 14px; border-bottom: 1px solid #edf2f7; font-size: 12px; }
            .fila-padre { background-color: #ffffff; font-weight: 500; }
            .fila-tarea { background-color: #fcfcfc; }
            tr:hover { background-color: #f8fafc !important; }
            .badge-tipo { padding: 4px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; text-transform: uppercase; }
            .badge-tipo.product-backlog-item { background: #e2e8f0; color: #334155; }
            .badge-tipo.bug { background: #fee2e2; color: #991b1b; }
            .badge-tipo.task, .badge-tipo.tarea { background: #dbeafe; color: #1d4ed8; }
            .badge-estado { padding: 4px 8px; border-radius: 999px; font-size: 10px; font-weight: 700; display: inline-block; text-transform: uppercase; }
            .badge-estado.done, .badge-estado.completed, .badge-estado.cerrado, .badge-estado.resolved { background: #dcfce7; color: #166534; }
            .badge-estado.active, .badge-estado.asignado, .badge-estado.en-progreso { background: #fef3c7; color: #92400e; }
            .progress-bar-container { background: #e2e8f0; border-radius: 999px; position: relative; height: 16px; width: 100%; overflow: hidden; }
            .progress-bar { background: linear-gradient(90deg, #0f766e, #22c55e); height: 100%; border-radius: 999px; transition: width 0.3s ease; }
            .progress-text { position: absolute; width: 100%; text-align: center; top: 0; left: 0; font-size: 10px; font-weight: 800; color: #0f172a; line-height: 16px; }
            .detail-panel { display: grid; grid-template-columns: 1fr; gap: 12px; }
            .detail-note { font-size: 12px; color: var(--muted); }
            .report-footer { margin-top: 18px; padding: 18px 4px 6px; }
            .report-footer__line { height: 1px; background: linear-gradient(90deg, rgba(15,23,42,0), rgba(15,23,42,0.2), rgba(15,23,42,0)); margin-bottom: 14px; }
            .report-footer__content { display: flex; align-items: center; gap: 14px; justify-content: center; color: var(--muted); }
            .report-footer__icon { width: 34px; height: 34px; object-fit: cover; border-radius: 8px; box-shadow: 0 6px 14px rgba(15,23,42,0.12); }
            .report-footer__title { font-size: 12px; font-weight: 800; color: var(--ink); text-align: center; }
            .report-footer__subtitle { font-size: 11px; text-align: center; }
            @media print { body { background: white; } .page { padding: 0; } .section, .hero { box-shadow: none; } }
            @media (max-width: 980px) { .hero, .charts-row, .kpi-row, .comparison-grid { grid-template-columns: 1fr; } .hero-grid { grid-template-columns: 1fr 1fr; } }
        </style>
        <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
    </head>
    <body>
        <div class="page">
            <section class="hero">
                <div>
                    <div class="eyebrow">Informe evolutivo y comparativo</div>
                    <h1>Reporte de avance consolidado</h1>
                    <p>La lectura ejecutiva parte del rango global ${tituloCobertura}. El bloque puntual compara el comportamiento de los últimos 15 días frente a los 15 días previos, para que el cliente vea evolución, carga y variación reciente.</p>
                    <div class="hero-meta">
                        <span class="pill">Cobertura desde ${formatearFechaReporte(contexto.fechaInicio || CONFIG.ejecucion.fechaInicio)}</span>
                        <span class="pill">Corte puntual ${formatearFechaReporte(comparativo.inicioActual)} al ${formatearFechaReporte(comparativo.finActual)}</span>
                        <span class="pill">Comparativo previo ${formatearFechaReporte(comparativo.inicioPrevio)} al ${formatearFechaReporte(comparativo.finPrevio)}</span>
                    </div>
                </div>
                <div class="hero-grid">
                    <div class="hero-card">
                        <div class="label">Historias / Bugs</div>
                        <div class="value">${metrizGlobales.totalHUs}</div>
                        <div class="note">Nivel superior del backlog trabajado</div>
                    </div>
                    <div class="hero-card">
                        <div class="label">Sub-elementos</div>
                        <div class="value">${metrizGlobales.totalTasks} / ${metrizGlobales.totalBugsHijos}</div>
                        <div class="note">Tasks y bugs vinculados</div>
                    </div>
                    <div class="hero-card">
                        <div class="label">Horas ejecutadas</div>
                        <div class="value">${metrizGlobales.totalEjec.toFixed(1)}h</div>
                        <div class="note">De ${metrizGlobales.totalEst.toFixed(1)}h estimadas</div>
                    </div>
                    <div class="hero-card">
                        <div class="label">Ejecución global</div>
                        <div class="value">${porcentajeGlobal.toFixed(1)}%</div>
                        <div class="note">Balance del periodo evaluado</div>
                    </div>
                </div>
            </section>

            <section class="section">
                <div class="section-header">
                    <div>
                        <h2>Lectura ejecutiva</h2>
                        <div class="meta">Comparativo de 15 días y evolución mensual consolidada</div>
                    </div>
                    <div class="meta">Generado automáticamente el ${new Date().toLocaleString()}</div>
                </div>
                <div class="kpi-row">
                    <div class="kpi-card">
                        <div class="label">Items últimos 15 días</div>
                        <div class="number">${comparativo.actual.items}</div>
                        <div class="sub">Variación: ${comparativo.variacionItems >= 0 ? '+' : ''}${comparativo.variacionItems} (${porcentajeVariacion.toFixed(1)}%)</div>
                    </div>
                    <div class="kpi-card">
                        <div class="label">Horas ejecutadas</div>
                        <div class="number">${comparativo.actual.horasEjecutadas.toFixed(1)}h</div>
                        <div class="sub">Variación: ${comparativo.variacionHoras >= 0 ? '+' : ''}${comparativo.variacionHoras.toFixed(1)}h (${porcentajeHoras.toFixed(1)}%)</div>
                    </div>
                    <div class="kpi-card">
                        <div class="label">HU / Tasks / Bugs</div>
                        <div class="number">${comparativo.actual.totalHus} / ${comparativo.actual.totalTasks} / ${comparativo.actual.totalBugs}</div>
                        <div class="sub">Corte puntual más reciente</div>
                    </div>
                    <div class="kpi-card">
                        <div class="label">Periodo previo</div>
                        <div class="number">${comparativo.previo.items}</div>
                        <div class="sub">Últimos 15 días anteriores</div>
                    </div>
                </div>

                <div class="comparison-grid">
                    <div class="comparison-card">
                        <strong>Rango global</strong>
                        <div class="line"><span>Inicio</span><span>${formatearFechaReporte(contexto.fechaInicio || CONFIG.ejecucion.fechaInicio)}</span></div>
                        <div class="line"><span>Fin</span><span>${formatearFechaReporte(contexto.fechaFin || CONFIG.ejecucion.fechaFin)}</span></div>
                        <div class="line"><span>Lectura</span><span class="delta">Evolutivo completo</span></div>
                    </div>
                    <div class="comparison-card">
                        <strong>Últimos 15 días</strong>
                        <div class="line"><span>Items</span><span>${comparativo.actual.items}</span></div>
                        <div class="line"><span>Horas</span><span>${comparativo.actual.horasEjecutadas.toFixed(1)}h</span></div>
                        <div class="line"><span>Foco</span><span class="delta">Actividad puntual</span></div>
                    </div>
                    <div class="comparison-card">
                        <strong>15 días previos</strong>
                        <div class="line"><span>Items</span><span>${comparativo.previo.items}</span></div>
                        <div class="line"><span>Horas</span><span>${comparativo.previo.horasEjecutadas.toFixed(1)}h</span></div>
                        <div class="line"><span>Lectura</span><span class="delta">Base comparativa</span></div>
                    </div>
                </div>

                <div class="charts-row">
                    <div class="chart-box">
                        <h3>Evolución mensual consolidada</h3>
                        <div class="chart-canvas-wrap"><canvas id="chartEvolutivo"></canvas></div>
                    </div>
                    <div class="chart-box">
                        <h3>Balance de horas</h3>
                        <div class="chart-canvas-wrap"><canvas id="chartHoras"></canvas></div>
                    </div>
                </div>
            </section>

            <section class="section detail-panel">
                <div class="section-header">
                    <div>
                        <h2>Detalle puntual de actividad</h2>
                        <div class="meta">Tareas y elementos actualizados dentro del corte de los últimos 15 días</div>
                    </div>
                    <div class="detail-note">Se muestran hasta 30 registros recientes para lectura del cliente.</div>
                </div>
                <table>
                    <thead>
                        <tr>
                            <th>Proyecto</th>
                            <th>Sprint</th>
                            <th>HU ID</th>
                            <th>Título</th>
                            <th>Tipo</th>
                            <th>Estado</th>
                            <th>Encargado</th>
                            <th style="text-align: center;">Cant Tareas</th>
                            <th style="text-align: center;">Cant Bugs</th>
                            <th style="text-align: right;">Hrs Estimadas</th>
                            <th style="text-align: right;">Hrs Ejecutadas</th>
                            <th style="width: 120px;">Progreso</th>
                            <th>Última Modificación</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${itemsPuntuales.map(r => `
                            <tr class="${r.esSubtarea ? 'fila-tarea' : 'fila-padre'}">
                                <td>${r.proyectoNombre}</td>
                                <td>${r.sprint}</td>
                                <td><strong>${r.huId}</strong></td>
                                <td>${r.esSubtarea ? '↳ ' : ''}${r.titulo}</td>
                                <td><span class="badge-tipo ${r.tipo.toLowerCase().replace(/\s+/g, '-')}">${r.tipo}</span></td>
                                <td><span class="badge-estado ${r.estado.toLowerCase()}">${r.estado}</span></td>
                                <td>${r.encargado}</td>
                                <td style="text-align: center; font-weight: bold; color: #4a5568;">${r.esSubtarea ? '-' : r.cantTareas}</td>
                                <td style="text-align: center; font-weight: bold; color: #a80000;">${r.esSubtarea ? '-' : r.cantBugs}</td>
                                <td style="text-align: right;">${r.horasEstimadas}</td>
                                <td style="text-align: right;">${r.horasEjecutadas}</td>
                                <td>
                                    <div class="progress-bar-container">
                                        <div class="progress-bar" style="width: ${r.porcentajeAvance}%"></div>
                                        <span class="progress-text">${Number(r.porcentajeAvance).toFixed(1)}%</span>
                                    </div>
                                </td>
                                <td>${r.ultimaModificacion}</td>
                            </tr>`).join('')}
                    </tbody>
                </table>

                <table>
                    <thead>
                        <tr>
                            <th>Proyecto</th>
                            <th>HU ID</th>
                            <th>Título</th>
                            <th>Tipo</th>
                            <th>Estado</th>
                            <th>Encargado</th>
                            <th style="text-align: right;">Horas</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${tablaRows}
                    </tbody>
                </table>
            </section>

            ${construirPieReporte()}
        </div>

        <script>
            new Chart(document.getElementById('chartHoras').getContext('2d'), {
                type: 'bar',
                data: {
                    labels: ['Horas Estimadas', 'Horas Ejecutadas'],
                    datasets: [{
                        data: [${metrizGlobales.totalEst.toFixed(1)}, ${metrizGlobales.totalEjec.toFixed(1)}],
                        backgroundColor: ['#94a3b8', '#0f766e'],
                        borderRadius: 8
                    }]
                },
                options: {
                    animation: false,
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: { y: { beginAtZero: true } }
                }
            });

            new Chart(document.getElementById('chartEvolutivo').getContext('2d'), {
                type: 'line',
                data: {
                    labels: ${JSON.stringify(mesesEvolutivo)},
                    datasets: [{
                        label: 'Items por mes',
                        data: ${JSON.stringify(valoresEvolutivo)},
                        tension: 0.35,
                        fill: true,
                        borderColor: '#c58b2f',
                        backgroundColor: 'rgba(197,139,47,0.14)',
                        pointRadius: 4,
                        pointBackgroundColor: '#0f172a'
                    }]
                },
                options: {
                    animation: false,
                    responsive: true,
                    maintainAspectRatio: false,
                    plugins: { legend: { display: false } },
                    scales: { y: { beginAtZero: true, precision: 0 } }
                }
            });
        </script>
    </body>
    </html>
    `;
}

// ===================================================
// GENERADOR DE INFORMES INDIVIDUALES (HTML + PDF) POR PROYECTO
// ===================================================
async function generarInformesGerencialesPorProyecto(rowsConsolidadas) {
    const dirInformes = path.join(process.cwd(), 'informesProyectos');

    if (!fs.existsSync(dirInformes)) {
        fs.mkdirSync(dirInformes, { recursive: true });
    }

    const proyectosMap = {};
    rowsConsolidadas.forEach(r => {
        if (!proyectosMap[r.proyectoNombre]) {
            proyectosMap[r.proyectoNombre] = [];
        }
        proyectosMap[r.proyectoNombre].push(r);
    });

    const nombresProyectos = Object.keys(proyectosMap);
    console.log(`\n🚀 Generando informes en PDF/HTML para ${nombresProyectos.length} proyecto(s)...`);

    const browser = await puppeteer.launch({ headless: 'new' });

    for (const nombreProyecto of nombresProyectos) {
        // Exclusión del proyecto Mas_Bosques en la generación de informes individuales
        if (nombreProyecto.toLowerCase().includes('masbosques') || nombreProyecto.toLowerCase().includes('mas_bosques')) {
            continue;
        }

        const items = (proyectosMap[nombreProyecto] || []).filter(i => !esEstadoRemovidoReporte(i.estado));
        const usaLogicaMasBosques = esProyectoMasBosques(nombreProyecto);
        const rangoMasBosques = usaLogicaMasBosques
            ? obtenerRangoMasBosquesDesdeEnv(CONFIG.ejecucionEspecifica.fechaInicio, CONFIG.ejecucionEspecifica.fechaFin)
            : null;

        const fechaInicioAnalisis = usaLogicaMasBosques
            ? new Date(rangoMasBosques.fechaInicio)
            : new Date(CONFIG.ejecucionEspecifica.fechaInicio);
        const fechaFinAnalisis = usaLogicaMasBosques
            ? new Date(rangoMasBosques.fechaFin)
            : new Date(CONFIG.ejecucionEspecifica.fechaFin);

        const itemsPeriodoPorFecha = items.filter(i => {
            const fechaCreacionItem = parsearFechaReporte(i.fechaCreacionRaw);
            const fechaCambioItem = parsearFechaReporte(i.fechaCambioRaw);
            const enCreacion = fechaCreacionItem && fechaCreacionItem >= fechaInicioAnalisis && fechaCreacionItem <= fechaFinAnalisis;
            const enCambio = fechaCambioItem && fechaCambioItem >= fechaInicioAnalisis && fechaCambioItem <= fechaFinAnalisis;

            if (enCreacion || enCambio) return true;

            const fechaItem = obtenerFechaFilaReporte(i);
            return fechaItem && fechaItem >= fechaInicioAnalisis && fechaItem <= fechaFinAnalisis;
        });
        const idsItemsPeriodoPorFecha = new Set(
            itemsPeriodoPorFecha.map(i => String(i.tareaId || i.huId || ''))
        );

        const backlogPadresProyecto = items.filter(i => {
            if (i.esSubtarea) return false;
            if (i.tipo === "Product Backlog Item") return true;
            if (i.tipo === "Bug") return !esEstadoRemovidoReporte(i.estado);
            return false;
        });
        const subtareasProyecto = items.filter(i => i.esSubtarea);
        const tieneSprintProyecto = !usaLogicaMasBosques && items.some(i => normalizarTextoReporte(i.sprint));
        const sprintSemanalMasBosques = usaLogicaMasBosques
            ? resolverSprintSemanalMasBosques(items, fechaInicioAnalisis, fechaFinAnalisis, parsearFechaReporte, obtenerSprintLimpioReporte)
            : '';
        const idsPadresVisiblesPorSubtarea = new Set(
            subtareasProyecto
                .filter(i => coincideSprintConPeriodoReporte(i.sprint, fechaInicioAnalisis, fechaFinAnalisis))
                .map(i => String(i.huId))
        );

        const idsPadresPorSprintMasBosques = new Set(
            subtareasProyecto
                .filter(i => normalizarSprint(obtenerSprintLimpioReporte(i.sprint)) === normalizarSprint(sprintSemanalMasBosques))
                .map(i => String(i.huId))
        );

        const itemsBacklogPadresPeriodo = tieneSprintProyecto
            ? backlogPadresProyecto.filter(i => {
                const idHu = String(i.huId || i.tareaId || '');
                return coincideSprintConPeriodoReporte(i.sprint, fechaInicioAnalisis, fechaFinAnalisis)
                    || idsPadresVisiblesPorSubtarea.has(idHu);
            })
            : itemsPeriodoPorFecha.filter(i => !i.esSubtarea && (i.tipo === "Product Backlog Item" || i.tipo === "Bug"));

        const itemsBacklogPadresPeriodoAjustado = usaLogicaMasBosques
            ? backlogPadresProyecto.filter(i => {
                const idHu = String(i.huId || i.tareaId || '');
                const sprintPadre = normalizarSprint(obtenerSprintLimpioReporte(i.sprint));
                const sprintObjetivo = normalizarSprint(sprintSemanalMasBosques);
                if (sprintObjetivo) {
                    return sprintPadre === sprintObjetivo || idsPadresPorSprintMasBosques.has(idHu);
                }
                return idsItemsPeriodoPorFecha.has(idHu);
            })
            : tieneSprintProyecto
            ? backlogPadresProyecto.filter(i => {
                const idHu = String(i.huId || i.tareaId || '');
                return coincideSprintConPeriodoReporte(i.sprint, fechaInicioAnalisis, fechaFinAnalisis)
                    || idsPadresVisiblesPorSubtarea.has(idHu);
            })
            : itemsBacklogPadresPeriodo;

        const idsPadresPeriodo = new Set(itemsBacklogPadresPeriodoAjustado.map(i => String(i.huId || i.tareaId || '')));
        const subtareasPeriodo = usaLogicaMasBosques
            ? subtareasProyecto.filter(i => {
                const idTarea = String(i.tareaId || i.huId || '');
                const sprintHijo = normalizarSprint(obtenerSprintLimpioReporte(i.sprint));
                const sprintObjetivo = normalizarSprint(sprintSemanalMasBosques);
                if (!idsPadresPeriodo.has(String(i.huId))) return false;
                if (sprintObjetivo) return sprintHijo === sprintObjetivo;
                return idsItemsPeriodoPorFecha.has(idTarea);
            })
            : tieneSprintProyecto
            ? subtareasProyecto.filter(i => {
                const idTarea = String(i.tareaId || i.huId || '');
                return idsPadresPeriodo.has(String(i.huId))
                    && (
                        coincideSprintConPeriodoReporte(i.sprint, fechaInicioAnalisis, fechaFinAnalisis)
                        || idsItemsPeriodoPorFecha.has(idTarea)
                    );
            })
            : itemsPeriodoPorFecha.filter(i => i.esSubtarea);

        const itemsPeriodo = [...itemsBacklogPadresPeriodoAjustado, ...subtareasPeriodo];
        const etiquetaSprintBacklogKpi = usaLogicaMasBosques
            ? (sprintSemanalMasBosques || 'Sin sprint')
            : tieneSprintProyecto
            ? obtenerSprintLimpioReporte(subtareasPeriodo[0]?.sprint || itemsBacklogPadresPeriodoAjustado[0]?.sprint)
            : 'Sin sprint';

        const husPadres = itemsBacklogPadresPeriodoAjustado.filter(i => i.tipo === "Product Backlog Item");
        const bugsPadres = itemsBacklogPadresPeriodoAjustado.filter(i => i.tipo === "Bug");
        const tareasHijas = subtareasPeriodo.filter(i => i.tipo !== "Bug");
        const idsHuSoporte = new Set(itemsBacklogPadresPeriodoAjustado.filter(i => esBacklogSoporteReporte(i.titulo)).map(i => String(i.huId || i.tareaId || '')));
        const idsHuSesiones = new Set(itemsBacklogPadresPeriodoAjustado.filter(i => esBacklogSesionesReporte(i.titulo)).map(i => String(i.huId || i.tareaId || '')));
        const idsHuPasoProduccion = new Set(itemsBacklogPadresPeriodoAjustado.filter(i => esBacklogPasoProduccionReporte(i.titulo)).map(i => String(i.huId || i.tareaId || '')));
        const idsHuEstimaciones = new Set(itemsBacklogPadresPeriodoAjustado.filter(i => esBacklogEstimacionesReporte(i.titulo)).map(i => String(i.huId || i.tareaId || '')));

        const tareasSoporte = tareasHijas.filter(i => idsHuSoporte.has(String(i.huId)));
        const tareasSesiones = tareasHijas.filter(i => idsHuSesiones.has(String(i.huId)));
        const tareasPasoProduccion = tareasHijas.filter(i => idsHuPasoProduccion.has(String(i.huId)));
        const tareasEspecialesIds = new Set([
            ...tareasSoporte.map(i => String(i.tareaId || i.huId || '')),
            ...tareasSesiones.map(i => String(i.tareaId || i.huId || '')),
            ...tareasPasoProduccion.map(i => String(i.tareaId || i.huId || ''))
        ]);
        const tareasGenerales = tareasHijas.filter(i => !tareasEspecialesIds.has(String(i.tareaId || i.huId || '')));

        const itemsSoporte = subtareasPeriodo.filter(i => idsHuSoporte.has(String(i.huId)));
        const itemsPasoProduccion = subtareasPeriodo.filter(i => idsHuPasoProduccion.has(String(i.huId)));

        const resumenBacklog = itemsBacklogPadresPeriodoAjustado
            .map(item => {
                const tareasAsociadas = tareasHijas.filter(tarea => String(tarea.huId) === String(item.huId));
                const horasEstimadasAsociadas = tareasAsociadas.reduce((acumulado, tarea) => acumulado + (parseFloat(tarea.horasEstimadas) || 0), 0);
                const horasEjecutadasAsociadas = tareasAsociadas.reduce((acumulado, tarea) => acumulado + (parseFloat(tarea.horasEjecutadas) || 0), 0);

                return {
                    huId: item.huId,
                    titulo: item.titulo,
                    tipo: item.tipo,
                    estado: item.estado,
                    cantTareasAsociadas: tareasAsociadas.length,
                    horasEstimadasAsociadas,
                    horasEjecutadasAsociadas
                };
            })
            .sort((a, b) => a.tipo.localeCompare(b.tipo, 'es', { sensitivity: 'base' }) || String(a.huId).localeCompare(String(b.huId), 'es', { numeric: true }));

        const conteoEstadosTareas = tareasHijas.reduce((acumulado, tarea) => {
            const estado = tarea.estado || 'Sin estado';
            acumulado[estado] = (acumulado[estado] || 0) + 1;
            return acumulado;
        }, {});

        const estadosTareasOrdenados = Object.entries(conteoEstadosTareas).sort((a, b) => b[1] - a[1]);
        const totalEstadosTareas = estadosTareasOrdenados.reduce((acumulado, [, valor]) => acumulado + valor, 0);
        const labelsEstadosTareas = estadosTareasOrdenados.map(([estado]) => estado);
        const valoresEstadosTareas = estadosTareasOrdenados.map(([, valor]) => valor);

        const horasTareasHijasEst = tareasHijas.reduce((acc, t) => acc + (parseFloat(t.horasEstimadas) || 0), 0);
        const horasTareasHijasEjec = tareasHijas.reduce((acc, t) => acc + (parseFloat(t.horasEjecutadas) || 0), 0);
        const horasSoporteEst = tareasSoporte.reduce((acc, t) => acc + (parseFloat(t.horasEstimadas) || 0), 0);
        const horasSoporteEjec = tareasSoporte.reduce((acc, t) => acc + (parseFloat(t.horasEjecutadas) || 0), 0);

        let labelsHistorialMensual = [];
        let valoresHistorialMensual = [];
        let valoresSoportesMensuales = [];

        if (usaLogicaMasBosques) {
            const inicioHistorialSprints = obtenerFechaInicioHistorialMasBosques(process.env);
            const historialSprintMap = {};
            let ordenSprint = 0;

            items.forEach(item => {
                const fechaItem = obtenerFechaFilaReporte(item);
                const enRangoFecha = !fechaItem || (fechaItem >= inicioHistorialSprints && fechaItem <= fechaFinAnalisis);
                if (!enRangoFecha) return;

                const sprint = obtenerSprintLimpioReporte(item.sprint);
                const claveSprint = normalizarSprint(sprint);
                if (!claveSprint || claveSprint === 'sin iteracion') return;

                if (!historialSprintMap[claveSprint]) {
                    historialSprintMap[claveSprint] = {
                        fechaOrden: fechaItem ? new Date(fechaItem) : null,
                        orden: ordenSprint++,
                        label: sprint,
                        est: 0,
                        ejec: 0,
                        soportesAtendidos: 0
                    };
                }

                if (fechaItem && (!historialSprintMap[claveSprint].fechaOrden || fechaItem < historialSprintMap[claveSprint].fechaOrden)) {
                    historialSprintMap[claveSprint].fechaOrden = new Date(fechaItem);
                }
            });

            subtareasProyecto
                .filter(item => item.tipo !== 'Bug')
                .forEach(item => {
                    const fechaItem = obtenerFechaFilaReporte(item);
                    const enRangoFecha = !fechaItem || (fechaItem >= inicioHistorialSprints && fechaItem <= fechaFinAnalisis);
                    if (!enRangoFecha) return;

                    const sprint = obtenerSprintLimpioReporte(item.sprint);
                    const claveSprint = normalizarSprint(sprint);
                    if (!claveSprint || claveSprint === 'sin iteracion') return;

                    if (!historialSprintMap[claveSprint]) {
                        historialSprintMap[claveSprint] = {
                            fechaOrden: fechaItem ? new Date(fechaItem) : null,
                            orden: ordenSprint++,
                            label: sprint,
                            est: 0,
                            ejec: 0,
                            soportesAtendidos: 0
                        };
                    }

                    if (fechaItem && (!historialSprintMap[claveSprint].fechaOrden || fechaItem < historialSprintMap[claveSprint].fechaOrden)) {
                        historialSprintMap[claveSprint].fechaOrden = new Date(fechaItem);
                    }

                    historialSprintMap[claveSprint].est += parseFloat(item.horasEstimadas) || 0;
                    historialSprintMap[claveSprint].ejec += parseFloat(item.horasEjecutadas) || 0;

                    const esSoporte = esBacklogSoporteReporte(item.titulo) || esBacklogSoporteReporte(item.tituloPadre);
                    const estado = (item.estado || '').toLowerCase();
                    const soporteAtendido = esSoporte && ['done', 'completed', 'cerrado', 'resolved'].includes(estado);
                    if (soporteAtendido) {
                        historialSprintMap[claveSprint].soportesAtendidos += 1;
                    }
                });

            const historialSprints = Object.values(historialSprintMap)
                .sort((a, b) => {
                    if (a.fechaOrden && b.fechaOrden) return a.fechaOrden - b.fechaOrden;
                    if (a.fechaOrden && !b.fechaOrden) return -1;
                    if (!a.fechaOrden && b.fechaOrden) return 1;
                    return a.orden - b.orden;
                })
                .map(sprint => ({
                    label: sprint.label,
                    avance: sprint.est > 0 ? Number(((sprint.ejec / sprint.est) * 100).toFixed(1)) : 0,
                    soportesAtendidos: sprint.soportesAtendidos
                }));

            labelsHistorialMensual = historialSprints.map(item => item.label);
            valoresHistorialMensual = historialSprints.map(item => item.avance);
            valoresSoportesMensuales = historialSprints.map(item => item.soportesAtendidos);
        } else {
            const historialMensualMap = {};
            const inicioHistorialMensual = new Date(
                fechaFinAnalisis.getFullYear(),
                fechaFinAnalisis.getMonth() - 2,
                1,
                0,
                0,
                0,
                0
            );
            const finHistorialMensual = new Date(
                fechaFinAnalisis.getFullYear(),
                fechaFinAnalisis.getMonth() + 1,
                0,
                23,
                59,
                59,
                999
            );

            const cursorMes = new Date(inicioHistorialMensual);
            while (cursorMes <= finHistorialMensual) {
                const keyMes = `${cursorMes.getFullYear()}-${String(cursorMes.getMonth() + 1).padStart(2, '0')}`;
                historialMensualMap[keyMes] = {
                    fechaOrden: new Date(cursorMes.getFullYear(), cursorMes.getMonth(), 1),
                    label: cursorMes.toLocaleDateString('es-CO', { month: 'long', year: 'numeric' }),
                    est: 0,
                    ejec: 0,
                    soportesAtendidos: 0
                };
                cursorMes.setMonth(cursorMes.getMonth() + 1);
            }

            subtareasProyecto
                .filter(item => item.tipo !== 'Bug')
                .forEach(item => {
                    const fechaItem = obtenerFechaFilaReporte(item);
                    if (!fechaItem) return;

                    if (fechaItem < inicioHistorialMensual || fechaItem > finHistorialMensual) return;

                    const keyMes = `${fechaItem.getFullYear()}-${String(fechaItem.getMonth() + 1).padStart(2, '0')}`;
                    if (!historialMensualMap[keyMes]) {
                        historialMensualMap[keyMes] = {
                            fechaOrden: new Date(fechaItem.getFullYear(), fechaItem.getMonth(), 1),
                            label: fechaItem.toLocaleDateString('es-CO', { month: 'long', year: 'numeric' }),
                            est: 0,
                            ejec: 0,
                            soportesAtendidos: 0
                        };
                    }

                    historialMensualMap[keyMes].est += parseFloat(item.horasEstimadas) || 0;
                    historialMensualMap[keyMes].ejec += parseFloat(item.horasEjecutadas) || 0;

                    const esSoporte = esBacklogSoporteReporte(item.titulo) || esBacklogSoporteReporte(item.tituloPadre);
                    const estado = (item.estado || '').toLowerCase();
                    const soporteAtendido = esSoporte && ['done', 'completed', 'cerrado', 'resolved'].includes(estado);
                    if (soporteAtendido) {
                        historialMensualMap[keyMes].soportesAtendidos += 1;
                    }
                });

            const historialMensual = Object.values(historialMensualMap)
                .sort((a, b) => a.fechaOrden - b.fechaOrden)
                .map(mes => ({
                    label: mes.label,
                    avance: mes.est > 0 ? Number(((mes.ejec / mes.est) * 100).toFixed(1)) : 0,
                    soportesAtendidos: mes.soportesAtendidos
                }));

            labelsHistorialMensual = historialMensual.map(item => item.label);
            valoresHistorialMensual = historialMensual.map(item => item.avance);
            valoresSoportesMensuales = historialMensual.map(item => item.soportesAtendidos);
        }

        const totalEst = obtenerHorasEstimadasObjetivoPorMes(fechaFinAnalisis);
        const totalEjec = tareasHijas.reduce((acc, i) => acc + (parseFloat(i.horasEjecutadas) || 0), 0);
        const horasSinSoporteEst = Math.max(totalEst - horasSoporteEst, 0);
        const horasSinSoporteEjec = Math.max(totalEjec - horasSoporteEjec, 0);
        const maxAvanceHistorial = valoresHistorialMensual.length > 0
            ? Math.max(...valoresHistorialMensual)
            : 100;
        const maxEscalaAvance = Math.max(100, Math.ceil((maxAvanceHistorial + 10) / 10) * 10);

        const itemsHuRegistradas = itemsBacklogPadresPeriodoAjustado.filter(i => {
            const idHu = String(i.huId || i.tareaId || '');
            const esHu = i.tipo === "Product Backlog Item";
            const esExcluida = idsHuSoporte.has(idHu)
                || idsHuSesiones.has(idHu)
                || idsHuPasoProduccion.has(idHu)
                || idsHuEstimaciones.has(idHu);
            return esHu && !esExcluida;
        });
        const bugsRegistrados = bugsPadres;
        const ESTADOS_PROGRESS_REPORTE = ['active', 'in progress', 'en progreso', 'asignado', 'doing', 'desarrollo', 'active/desarrollo'];
        const ESTADOS_DONE_REPORTE = ['done', 'completed', 'cerrado', 'resolved', 'closed', 'finalizado'];
        const husRegistradasOrdenadas = itemsHuRegistradas
            .map(hu => {
                const tareasAsociadas = subtareasPeriodo.filter(tarea => String(tarea.huId) === String(hu.huId));
                const horasEjecutadasHu = parseFloat(hu.horasEjecutadas) || 0;
                const horasEjecutadasTareas = tareasAsociadas.reduce((acc, tarea) => acc + (parseFloat(tarea.horasEjecutadas) || 0), 0);
                const horasRegistradasTotales = horasEjecutadasHu + horasEjecutadasTareas;
                const estadosTareas = tareasAsociadas.map(tarea => String(tarea.estado || '').trim().toLowerCase());
                const tieneTareaInProgress = estadosTareas.some(estado => ESTADOS_PROGRESS_REPORTE.includes(estado));
                const tieneTareaDone = estadosTareas.some(estado => ESTADOS_DONE_REPORTE.includes(estado));

                let prioridad = 3;
                if (horasRegistradasTotales === 0) {
                    prioridad = 0;
                } else if (tieneTareaInProgress) {
                    prioridad = 1;
                } else if (tieneTareaDone) {
                    prioridad = 2;
                }

                return {
                    ...hu,
                    prioridad,
                    horasRegistradasTotales
                };
            })
            .sort((a, b) => {
                if (a.prioridad !== b.prioridad) return a.prioridad - b.prioridad;
                if (a.horasRegistradasTotales !== b.horasRegistradasTotales) return a.horasRegistradasTotales - b.horasRegistradasTotales;
                return String(a.huId).localeCompare(String(b.huId), 'es', { numeric: true });
            });
        const descripcion = DESCRIPCIONES_PROYECTOS[nombreProyecto] || DESCRIPCION_ALCANCE_DEFAULT;
        const comparativoProyecto = calcularComparativoPeriodo(itemsPeriodo, fechaFinAnalisis);
        const periodoEtiqueta = `${fechaInicioAnalisis.toLocaleDateString('es-CO', { day: '2-digit', month: 'long', year: 'numeric' })} al ${fechaFinAnalisis.toLocaleDateString('es-CO', { day: '2-digit', month: 'long', year: 'numeric' })}`;
        const fechaInicioTitulo = usaLogicaMasBosques
            ? fechaInicioAnalisis
            : CONFIG.ejecucionEspecifica.fechaInicio;
        const fechaFinTitulo = usaLogicaMasBosques
            ? fechaFinAnalisis
            : CONFIG.ejecucionEspecifica.fechaFin;
        const periodoEtiquetaTitulo = `${fechaInicioTitulo.toLocaleDateString('es-CO', { day: '2-digit', month: 'long', year: 'numeric' })} al ${fechaFinTitulo.toLocaleDateString('es-CO', { day: '2-digit', month: 'long', year: 'numeric' })}`;
        const porcentajeGlobalProyecto = totalEst > 0 ? (totalEjec / totalEst) * 100 : 0;
        const notaVisualProyecto = obtenerNotaVisualProyecto(nombreProyecto);
        const notaVisualProyectoHtml = formatearNotaVisualProyectoHtml(notaVisualProyecto);
        const resumenProyecto = obtenerResumenProyecto(nombreProyecto);
        const resumenEjecutivoHtml = renderResumenProyectoHtml(resumenProyecto);

        const htmlContent = `
        <!DOCTYPE html>
        <html lang="es">
        <head>
            <meta charset="UTF-8">
            <title>Reporte de Gestion - ${nombreProyecto}</title>
            <script src="https://cdn.jsdelivr.net/npm/chart.js"></script>
            <script src="https://cdn.jsdelivr.net/npm/chartjs-plugin-datalabels@2"></script>
            <style>
                @page { size: A4; margin: 12mm; }
                :root {
                    --lpz-primary: #ce0a1e;
                    --lpz-primary-dark: #990816;
                    --lpz-navy: #3d0b12;
                    --lpz-ink: #45141a;
                    --lpz-soft: #ffeef1;
                    --lpz-card: #fff8f9;
                }
                body { font-family: 'Segoe UI', Tahoma, Arial, sans-serif; color: var(--lpz-ink); background: #ffffff; line-height: 1.35; font-size: 9.5pt; margin: 0; }
                .page { padding: 8px; }
                .header { background: linear-gradient(135deg, var(--lpz-primary-dark) 0%, #7a0f1b 45%, var(--lpz-primary) 100%); color: #ffffff; padding: 18px; border-radius: 16px; margin-bottom: 10px; position: relative; overflow: hidden; }
                .header::after { content: ''; position: absolute; right: -20px; bottom: -30px; width: 140px; height: 140px; border-radius: 50%; background: rgba(255,255,255,0.08); }
                .header h1 { margin: 0; font-size: 18pt; letter-spacing: -0.5px; }
                .header .meta { margin-top: 5px; color: #ffd7de; font-size: 9pt; font-weight: 700; text-transform: uppercase; }
                .header-brand { display: flex; align-items: center; gap: 10px; margin-bottom: 6px; }
                .header-brand__icon { height: 30px; width: auto; max-width: 72px; object-fit: contain; margin-left: auto; flex-shrink: 0; }
                .description-box { background: rgba(248,252,255,0.95); border-left: 4px solid var(--lpz-primary); padding: 9px 13px; margin-bottom: 10px; border-radius: 0 12px 12px 0; }
                .dev-note { background: #ffffff; border: 1px solid #f3c2ca; color: #5a1d25; border-radius: 10px; padding: 8px 11px; margin-bottom: 10px; font-size: 8.5pt; }
                .dev-note-list { margin: 6px 0 0 16px; padding: 0; }
                .dev-note-list li { margin: 0 0 6px 0; line-height: 1.45; }
                .dev-note-list li:last-child { margin-bottom: 0; }
                .summary-box { background: #ffffff; border: 1px solid #f3c2ca; color: #5a1d25; border-radius: 10px; padding: 8px 11px; margin-bottom: 10px; font-size: 8.5pt; }
                .summary-intro { margin: 6px 0 8px 0; line-height: 1.45; }
                .summary-paragraph { margin: 5px 0; line-height: 1.35; }
                .summary-activities-title { margin-top: 8px; font-weight: 700; }
                .summary-list { margin: 0 0 0 16px; padding: 0; }
                .summary-list li { margin: 0 0 4px 0; line-height: 1.35; }
                .summary-list li:last-child { margin-bottom: 0; }
                .kpi-row { display: grid; grid-template-columns: repeat(3, minmax(0, 1fr)); gap: 8px; margin-bottom: 10px; align-items: stretch; }
                .kpi-card { background: #ffffff; border: 1px solid #f3c2ca; padding: 10px; border-radius: 12px; text-align: center; box-shadow: 0 2px 8px rgba(121, 15, 27, 0.08); }
                .kpi-card .number { font-size: 17pt; font-weight: 800; color: var(--lpz-navy); }
                .kpi-card .label { font-size: 7.5pt; color: #315875; text-transform: uppercase; font-weight: 700; margin-top: 4px; }
                .state-kpi-row { display: grid; grid-template-columns: repeat(2, minmax(0, 1fr)); gap: 8px; margin-bottom: 10px; align-items: stretch; }
                .state-kpi-card { background: #ffffff; border: 1px solid #f3c2ca; border-radius: 12px; padding: 10px; text-align: center; }
                .state-kpi-card .number { font-size: 17pt; font-weight: 800; color: var(--lpz-navy); }
                .state-kpi-card .label { font-size: 8pt; color: #244863; text-transform: uppercase; font-weight: 700; margin-top: 4px; }
                .comparison-box { display: grid; grid-template-columns: 1fr; gap: 8px; margin-bottom: 10px; }
                .comparison-card { background: #ffffff; border: 1px solid #f3c2ca; border-radius: 12px; padding: 10px; }
                .comparison-card strong { display: block; margin-bottom: 8px; color: var(--lpz-navy); }
                .comparison-card .line { display: flex; justify-content: space-between; font-size: 8.5pt; color: #36556d; padding: 2px 0; }
                .chart-section { margin-bottom: 10px; }
                .chart-section-title { font-size: 10pt; color: var(--lpz-navy); font-weight: 800; margin-bottom: 6px; }
                .charts-row { display: grid; gap: 10px; margin-bottom: 10px; page-break-inside: avoid; break-inside: avoid-page; align-items: start; }
                .charts-row-two { grid-template-columns: repeat(2, minmax(0, 1fr)); }
                .charts-row-single { grid-template-columns: 1fr; }
                .chart-box { background: var(--lpz-card); border: 1px solid #f3c2ca; border-radius: 12px; padding: 10px; text-align: center; }
                .chart-wide { grid-column: 1 / -1; }
                .chart-box h3 { margin: 0 0 8px 0; font-size: 9pt; color: var(--lpz-navy); }
                .chart-help { margin: 0 0 10px 0; font-size: 8pt; color: #5a1d25; }
                .chart-canvas-wrap { position: relative; width: 100%; height: 190px; }
                .chart-canvas-wrap--compact { height: 170px; }
                .chart-canvas-wrap--evolucion { height: 320px; }
                .chart-canvas-wrap--donut {
                    height: auto;
                    aspect-ratio: 1 / 1;
                    max-width: 220px;
                    margin: 0 auto;
                }
                .chart-canvas-wrap canvas { width: 100% !important; height: 100% !important; }
                .section-title { font-size: 11pt; border-bottom: 2px solid var(--lpz-primary); padding-bottom: 4px; color: var(--lpz-navy); margin: 14px 0 8px 0; font-weight: 800; page-break-after: avoid; }
                table { width: 100%; border-collapse: collapse; margin-bottom: 10px; font-size: 8.5pt; }
                tr { page-break-inside: avoid; }
                th { background-color: #6f0a16; color: white; padding: 7px 10px; text-align: left; }
                td { padding: 5px 9px; border-bottom: 1px solid #f5d6dc; color: #5a1d25; }
                tr:nth-child(even) { background-color: #fff2f4; }
                .badge-alert { background-color: #fef2f2; color: #991b1b; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 8pt; }
                .badge-soporte { background-color: #f0fdf4; color: #166534; padding: 2px 6px; border-radius: 4px; font-weight: bold; font-size: 8pt; }
                .page-break { page-break-before: always; }
                .task-list-note { font-size: 8.5pt; color: #475569; margin-bottom: 8px; }
                .report-block { break-inside: avoid-page; page-break-inside: avoid; }
                .table-section { break-inside: avoid-page; page-break-inside: avoid; margin-bottom: 10px; }
                .table-section p { margin-top: 0; }
                .backlog-summary-table { break-inside: auto; page-break-inside: auto; }
                .backlog-summary-table thead { display: table-header-group; }
                .backlog-summary-table tbody { break-inside: auto; page-break-inside: auto; }
                @media print {
                    .description-box,
                    .summary-box,
                    .dev-note,
                    .kpi-row,
                    .state-kpi-row,
                    .comparison-box,
                    .comparison-card,
                    .chart-section,
                    .chart-box,
                    .table-section,
                    table {
                        break-inside: avoid-page;
                        page-break-inside: avoid;
                    }

                    .backlog-summary-table,
                    .backlog-summary-table tbody {
                        break-inside: auto;
                        page-break-inside: auto;
                    }
                }
            </style>
        </head>
        <body>
            <div class="page">
                <div class="header">
                    <div class="header-brand">
                        <div class="meta">Reporte de Gestion</div>
                        
                    </div>
                    <h1>Proyecto: ${nombreProyecto}</h1>
                    <div style="margin-top: 8px; font-size: 8.5pt; color: #ffe0e6;">Análisis del ${periodoEtiquetaTitulo}</div>
                </div>

                ${resumenEjecutivoHtml}

                <div class="description-box report-block">
                    <strong>Descripción del alcance:</strong> ${descripcion}
                </div>

                ${notaVisualProyectoHtml ? `<div class="dev-note"><strong>Nota:</strong>${notaVisualProyectoHtml}</div>` : ''}

                <div class="kpi-row report-block">
                    <div class="kpi-card">
                        <div class="number">${new Set(husPadres.map(hu => String(hu.huId || hu.tareaId || ''))).size}</div>
                        <div class="label">Cantidad de HU </div>
                    </div>
                    <div class="kpi-card">
                        <div class="number" style="color: #b91c1c;">${bugsPadres.length}</div>
                        <div class="label">Cantidad de Bugs</div>
                    </div>
                    <div class="kpi-card">
                        <div class="number" style="color: #4a5568;">${tareasHijas.length}</div>
                        <div class="label">Cantidad de Tareas </div>
                    </div>
                    <div class="kpi-card">
                        <div class="number" style="color: #d97706;">${tareasSoporte.length}</div>
                        <div class="label">Tareas en Soporte</div>
                    </div>
                    <div class="kpi-card">
                        <div class="number" style="color: #0f766e;">${contarItemsUnicosReporte(tareasSesiones.filter(i => esTipoTareaOperativaReporte(i.tipo)))}</div>
                        <div class="label">Tareas en Sesiones</div>
                    </div>
                    <div class="kpi-card">
                        <div class="number" style="color: #c2410c;">${tareasPasoProduccion.length}</div>
                        <div class="label">Tareas en Paso a Producción</div>
                    </div>
                </div>

                <div class="state-kpi-row report-block">
                    <div class="state-kpi-card">
                        <div class="number" style="color: #0f766e;">${totalEst.toFixed(1)}h</div>
                        <div class="label">Horas Estimadas</div>
                    </div>
                    <div class="state-kpi-card">
                        <div class="number" style="color: #16a34a;">${totalEjec.toFixed(1)}h</div>
                        <div class="label">Horas Ejecutadas</div>
                    </div>
                </div>

                <div class="comparison-box report-block">
                    <div class="comparison-card">
                        <strong>Balance del mes</strong>
                        <div class="line"><span>Sprint backlog</span><span>${etiquetaSprintBacklogKpi}</span></div>
                        <div class="line"><span>Tareas en soporte</span><span>${contarItemsUnicosReporte(tareasSoporte.filter(i => esTipoTareaOperativaReporte(i.tipo)))}</span></div>
                        <div class="line"><span>Paso a producción</span><span>${tareasPasoProduccion.length}</span></div>
                        <div class="line"><span>Horas ejecutadas</span><span>${totalEjec.toFixed(1)}h</span></div>
                        <div class="line"><span>Avance</span><span style="font-weight:700; color:#ce0a1e;">${porcentajeGlobalProyecto.toFixed(1)}%</span></div>
                    </div>
                </div>

                <div class="chart-section report-block">
                    <div class="charts-row charts-row-two">
                        <div class="chart-box">
                            <h3>Distribución de Tareas Hijas</h3>
                            <div class="chart-canvas-wrap"><canvas id="chartDistribucion"></canvas></div>
                        </div>
                        <div class="chart-box">
                            <h3>Estado de Tareas (%)</h3>
                            <div class="chart-canvas-wrap chart-canvas-wrap--donut"><canvas id="chartEstadosTareas"></canvas></div>
                        </div>
                    </div>
                </div>

                <div class="chart-section report-block">
                    <div class="chart-section-title">Balance de carga horaria</div>
                    <div class="charts-row charts-row-single">
                        <div class="chart-box chart-wide">
                            <h3>Horas ejecutadas por tipo (periodo específico)</h3>
                            <div class="chart-canvas-wrap"><canvas id="chartHoras"></canvas></div>
                        </div>
                    </div>
                </div>

                <div class="chart-section report-block">
                    <div class="chart-section-title">Evolución mensual</div>
                    <div class="charts-row charts-row-single">
                        <div class="chart-box chart-wide">
                            <h3>Avance mensual de ejecución (%)</h3>
                            <div class="chart-canvas-wrap chart-canvas-wrap--evolucion"><canvas id="chartEvolucionMensual"></canvas></div>
                        </div>
                    </div>
                </div>

            <div class="table-section">
            <div class="section-title">1. Historias de Usuario Registradas</div>
            ${husRegistradasOrdenadas.length > 0 ? `
            <table>
                <thead>
                    <tr>
                        <th style="width: 80px;">HU ID</th>
                        <th>Título de Historia de Usuario</th>
                        <th>Estado</th>
                        <th>Persona Encargada</th>
                    </tr>
                </thead>
                <tbody>
                    ${husRegistradasOrdenadas.map(hu => `
                        <tr>
                            <td><strong>#${hu.huId}</strong></td>
                            <td>${hu.titulo}</td>
                            <td><span class="badge-alert">${hu.estado}</span></td>
                            <td>${hu.encargado}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            ` : '<p style="color: #64748b; font-size: 8.5pt;">No se registraron Historias de Usuario en el período auditado.</p>'}
            </div>

            <div class="table-section">
            <div class="section-title">2. Bugs Registrados</div>
            ${bugsRegistrados.length > 0 ? `
            <table>
                <thead>
                    <tr>
                        <th style="width: 80px;">Bug ID</th>
                        <th>Título del Bug</th>
                        <th>Estado</th>
                        <th>Persona Encargada</th>
                    </tr>
                </thead>
                <tbody>
                    ${bugsRegistrados.map(bug => `
                        <tr>
                            <td><strong>#${bug.huId || bug.tareaId || '-'}</strong></td>
                            <td>${bug.titulo}</td>
                            <td><span class="badge-alert">${bug.estado}</span></td>
                            <td>${bug.encargado || 'Sin asignar'}</td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            ` : '<p style="color: #64748b; font-size: 8.5pt;">No se registraron bugs en el período auditado.</p>'}
            </div>

            <div class="table-section">
            <div class="section-title">3. Desglose de Actividades de Soporte Contabilizadas</div>
            ${itemsSoporte.length > 0 ? `
            <table>
                <thead>
                    <tr>
                        <th style="width: 80px;">ID</th>
                        <th>Actividad / Detalle de Soporte</th>
                        <th>Tipo</th>
                        <th>Estado</th>
                        <th style="text-align: right;">Horas Ejecutadas</th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsSoporte.map(s => `
                        <tr>
                            <td>#${s.tareaId || s.huId}</td>
                            <td>${s.titulo}</td>
                            <td>${s.tipo}</td>
                            <td>${s.estado}</td>
                            <td style="text-align: right;"><span class="badge-soporte">${s.horasEjecutadas}h</span></td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            ` : '<p style="color: #64748b; font-size: 8.5pt;">No se registraron tareas o bugs etiquetados como soporte en el período auditado.</p>'}
            </div>

            <div class="table-section">
            <div class="section-title">4. Desglose de Actividades de Paso a Producción</div>
            ${itemsPasoProduccion.length > 0 ? `
            <table>
                <thead>
                    <tr>
                        <th style="width: 80px;">ID</th>
                        <th>Actividad / Detalle de Paso a Producción</th>
                        <th>Tipo</th>
                        <th>Estado</th>
                        <th style="text-align: right;">Horas Ejecutadas</th>
                    </tr>
                </thead>
                <tbody>
                    ${itemsPasoProduccion.map(s => `
                        <tr>
                            <td>#${s.tareaId || s.huId}</td>
                            <td>${s.titulo}</td>
                            <td>${s.tipo}</td>
                            <td>${s.estado}</td>
                            <td style="text-align: right;"><span class="badge-soporte">${s.horasEjecutadas}h</span></td>
                        </tr>
                    `).join('')}
                </tbody>
            </table>
            ` : '<p style="color: #64748b; font-size: 8.5pt;">No se registraron tareas asociadas a paso a producción en el período auditado.</p>'}
            </div>

            </div>

            <div class="page page-break">
                <div class="header" style="margin-top: 0;">
                    <div class="meta">Detalle operativo</div>
                    <h1>Resumen de backlog y tareas asociadas</h1>
                    <div style="margin-top: 8px; font-size: 8.5pt; color: #cbd5e1;">${nombreProyecto} | ${periodoEtiqueta}</div>
                </div>

                <div class="task-list-note">Esta página resume cada item de backlog del período específico, su estado actual y la cantidad de tareas hijas asociadas que sí entraron en el corte.</div>

                ${resumenBacklog.length > 0 ? `
                <table class="backlog-summary-table">
                    <thead>
                        <tr>
                            <th style="width: 70px;">ID</th>
                            <th>Item Backlog</th>
                            <th style="width: 90px;">Tipo</th>
                            <th style="width: 120px;">Estado</th>
                            <th style="width: 90px; text-align: right;">Tareas</th>
                            <th style="width: 85px; text-align: right;">Est.</th>
                            <th style="width: 85px; text-align: right;">Ejec.</th>
                        </tr>
                    </thead>
                    <tbody>
                        ${resumenBacklog
                            .map(t => `
                                <tr>
                                    <td>#${t.huId || '-'}</td>
                                    <td>${t.titulo || 'Sin título'}</td>
                                    <td>${t.tipo || 'Sin tipo'}</td>
                                    <td>${t.estado || 'Sin estado'}</td>
                                    <td style="text-align: right;">${t.cantTareasAsociadas}</td>
                                    <td style="text-align: right;">${t.horasEstimadasAsociadas.toFixed(1)}h</td>
                                    <td style="text-align: right;">${t.horasEjecutadasAsociadas.toFixed(1)}h</td>
                                </tr>
                            `).join('')}
                    </tbody>
                </table>
                ` : '<p style="color: #64748b; font-size: 8.5pt;">No se encontraron items de backlog dentro del período específico.</p>'}

            <script>
                Chart.register(ChartDataLabels);

                new Chart(document.getElementById('chartHoras').getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: ['Horas Tareas', 'Horas Soporte'],
                        datasets: [{
                            label: 'Horas ejecutadas',
                            data: [${horasSinSoporteEjec.toFixed(1)}, ${horasSoporteEjec.toFixed(1)}],
                            backgroundColor: ['#1e88e5', '#f89c3c'],
                            borderRadius: 6
                        }]
                    },
                    options: {
                        animation: false,
                        responsive: true,
                        maintainAspectRatio: false,
                        plugins: {
                            legend: { display: false },
                            tooltip: {
                                callbacks: {
                                    footer: (items) => {
                                        if (!items || !items.length) return '';
                                        return 'Total: ' + Number(items[0].parsed.y || 0).toFixed(1) + 'h';
                                    }
                                }
                            },
                            datalabels: {
                                color: '#0b1f3a',
                                anchor: 'end',
                                align: 'top',
                                font: { size: 8, weight: '700' },
                                formatter: (value) => Number(value || 0).toFixed(1) + 'h'
                            }
                        },
                        scales: {
                            y: { beginAtZero: true }
                        }
                    }
                });

                new Chart(document.getElementById('chartDistribucion').getContext('2d'), {
                    type: 'bar',
                    data: {
                        labels: ['Tareas generales', 'Tareas soporte', 'Tareas sesiones', 'Tareas paso a producción'],
                        datasets: [{
                            data: [${tareasGenerales.length}, ${tareasSoporte.length}, ${tareasSesiones.length}, ${tareasPasoProduccion.length}],
                            backgroundColor: ['#0d4a92', '#1e88e5', '#ec6aa8', '#be00ff']
                        }]
                    },
                    options: {
                        animation: false,
                        responsive: true,
                        maintainAspectRatio: false,
                        indexAxis: 'y',
                        plugins: {
                            legend: { display: false },
                            datalabels: {
                                color: '#0b1f3a',
                                anchor: 'end',
                                align: 'right',
                                font: { size: 8, weight: '700' },
                                formatter: (value, context) => {
                                    const data = context.dataset.data;
                                    const total = data.reduce((acc, n) => acc + n, 0);
                                    const porcentaje = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                                    return value + ' (' + porcentaje + '%)';
                                }
                            }
                        },
                        scales: {
                            x: { beginAtZero: true, ticks: { precision: 0 } },
                            y: { ticks: { font: { size: 8 } } }
                        }
                    }
                });

                new Chart(document.getElementById('chartEstadosTareas').getContext('2d'), {
                    type: 'doughnut',
                    data: {
                        labels: ${JSON.stringify(labelsEstadosTareas)},
                        datasets: [{
                            data: ${JSON.stringify(valoresEstadosTareas)},
                            backgroundColor: ['#0d4a92', '#1e88e5', '#ec6aa8', '#be00ff', '#56c58a', '#f89c3c', '#f2de59', '#d62839']
                        }]
                    },
                    options: {
                        animation: false,
                        responsive: true,
                        maintainAspectRatio: true,
                        aspectRatio: 1,
                        plugins: {
                            legend: { position: 'bottom', labels: { boxWidth: 10, font: { size: 9 } } },
                            datalabels: {
                                color: '#ffffff',
                                font: { size: 8, weight: '700' },
                                formatter: (value, context) => {
                                    const data = context.dataset.data;
                                    const total = data.reduce((acc, n) => acc + n, 0);
                                    const porcentaje = total > 0 ? ((value / total) * 100).toFixed(1) : '0.0';
                                    return porcentaje + '%';
                                }
                            }
                        }
                    }
                });

                new Chart(document.getElementById('chartEvolucionMensual').getContext('2d'), {
                    type: 'line',
                    data: {
                        labels: ${JSON.stringify(labelsHistorialMensual.map(label => abreviarEtiquetaSprintReporte(label, 18)))},
                        datasets: [
                            {
                                type: 'line',
                                label: 'Avance de ejecución (%)',
                                data: ${JSON.stringify(valoresHistorialMensual)},
                                tension: 0.3,
                                fill: true,
                                borderColor: '#1e88e5',
                                backgroundColor: 'rgba(30,136,229,0.18)',
                                pointBackgroundColor: '#0d4a92',
                                pointRadius: 4,
                                yAxisID: 'yPorcentaje'
                            },
                            {
                                type: 'bar',
                                label: 'Soportes atendidos',
                                data: ${JSON.stringify(valoresSoportesMensuales)},
                                backgroundColor: 'rgba(86,197,138,0.45)',
                                borderColor: '#56c58a',
                                borderWidth: 1,
                                yAxisID: 'ySoportes'
                            }
                        ]
                    },
                    options: {
                        animation: false,
                        responsive: true,
                        maintainAspectRatio: false,
                        layout: {
                            padding: {
                                bottom: 10,
                                left: 8,
                                right: 8
                            }
                        },
                        plugins: {
                            legend: { display: true, position: 'bottom' },
                            tooltip: {
                                callbacks: {
                                    title: (items) => {
                                        if (!items || !items.length) return '';
                                        return ${JSON.stringify(labelsHistorialMensual)}[items[0].dataIndex] || '';
                                    }
                                }
                            },
                            datalabels: {
                                color: context => {
                                    if (context.dataset.type !== 'line') return '#0b1f3a';
                                    const valor = Number(context.dataset.data[context.dataIndex] || 0);
                                    return valor > 100 ? '#b91c1c' : '#0b1f3a';
                                },
                                font: { size: 8, weight: '700' },
                                formatter: (value, context) => context.dataset.type === 'line' ? (String(value) + '%') : String(value),
                                align: context => {
                                    if (context.dataset.type !== 'line') return 'end';
                                    const valor = Number(context.dataset.data[context.dataIndex] || 0);
                                    return valor > 100 ? 'bottom' : 'top';
                                },
                                anchor: context => context.dataset.type === 'line' ? 'end' : 'end'
                            }
                        },
                        scales: {
                            yPorcentaje: {
                                type: 'linear',
                                position: 'left',
                                beginAtZero: true,
                                max: ${maxEscalaAvance},
                                ticks: {
                                    callback: value => value + '%'
                                }
                            },
                            ySoportes: {
                                type: 'linear',
                                position: 'right',
                                beginAtZero: true,
                                grid: {
                                    drawOnChartArea: false
                                }
                            },
                            x: {
                                ticks: {
                                    autoSkip: true,
                                    maxTicksLimit: 8,
                                    maxRotation: 0,
                                    minRotation: 0,
                                    font: { size: 8 }
                                },
                                offset: true
                            }
                        }
                    }
                });
            </script>
            </div>
        </body>
        </html>
        `;

        const nombreLimpio = nombreProyecto.replace(/[^a-zA-Z0-9_-]/g, '_');
        const rutaHtml = path.join(dirInformes, `Informe_${nombreLimpio}.html`);
        const rutaPdf = path.join(dirInformes, `Informe_${nombreLimpio}.pdf`);

        fs.writeFileSync(rutaHtml, htmlContent, 'utf-8');

        const page = await browser.newPage();
        await page.setContent(htmlContent, { waitUntil: 'networkidle0' });
        await page.pdf({
            path: rutaPdf,
            format: 'A4',
            printBackground: true,
            margin: { top: '12mm', right: '12mm', bottom: '12mm', left: '12mm' }
        });
        await page.close();

        const membreteAplicado = await aplicarMembreteSobrePdf(rutaPdf);
        if (!membreteAplicado) {
            console.warn(` ⚠️ No se pudo aplicar membrete en ${path.basename(rutaPdf)}. Verifica assets/formato.pdf.`);
        }

        console.log(` ✅ Generado: Informe_${nombreLimpio}.pdf y .html`);
    }

    await browser.close();
    console.log(`📂 Todos los informes gerenciales quedaron guardados en: ${dirInformes}\n`);
}

async function exportarDatosDashboard() {
    if (!validarDiaEjecucion()) return;

    let globales = { totalHUs: 0, totalTasks: 0, totalBugsHijos: 0, totalEst: 0, totalEjec: 0, totalSoporte: 0 };
    let rowsConsolidadas = [];
    let logsConsolaProyectos = [];
    let encargadosPorProyectoAcumulador = [];
    
    let desgloseEstadosPorHU = {};
    let conteoMensual = {}; 

    try {
        console.log("==================================================");
        console.log(`[AUTOMATIZACIÓN] Iniciando extracción Multi-Proyecto...`);
        console.log(`[RANGO] Desde: ${CONFIG.ejecucion.fechaInicio.toLocaleDateString()} Hasta: ${CONFIG.ejecucion.fechaFin.toLocaleDateString()}`);
        console.log("==================================================");

        const proyectosUrl = `${CONFIG.baseUrl}/_apis/projects?api-version=${CONFIG.apiVersion}`;
        const proyectosResponse = await fetch(proyectosUrl, { headers: { 'Authorization': authHeader } });
        if (!proyectosResponse.ok) throw new Error(`Error proyectos: ${proyectosResponse.statusText}`);
       
        const proyectosData = await proyectosResponse.json();
        const listaProyectos = proyectosData.value || [];

        for (const proyecto of listaProyectos) {
            const projectName = proyecto.name;
            const projectId = proyecto.id;

            // Exclusión del proyecto Mas_Bosques en la extracción general
            if (projectName.toLowerCase().includes('masbosques') || projectName.toLowerCase().includes('mas_bosques')) {
                continue;
            }

            let husProyecto = 0;
            let tasksProyecto = 0;
            let bugsProyecto = 0;
            let horasEstimadasProyecto = 0;
            let horasEjecutadasProyecto = 0;

            let huToDo = 0;
            let huInProgress = 0;
            let huDone = 0;

            let taskToDo = 0;
            let taskInProgress = 0;
            let taskDone = 0;

            try {
                const wiqlUrl = `${CONFIG.baseUrl}/${projectId}/_apis/wit/wiql?api-version=${CONFIG.apiVersion}`;
                const wiqlBody = { query: `Select [System.Id] From WorkItems Where [System.TeamProject] = '${projectName}'` };

                const wiqlResponse = await fetch(wiqlUrl, {
                    method: 'POST',
                    headers: { 'Content-Type': 'application/json', 'Authorization': authHeader },
                    body: JSON.stringify(wiqlBody)
                });

                if (!wiqlResponse.ok) continue;

                const wiqlData = await wiqlResponse.json();
                if (!wiqlData.workItems || wiqlData.workItems.length === 0) continue;

                const idsProyecto = wiqlData.workItems.map(item => item.id);
                const chunkDetailRows = await obtenerDetallesMasivos(idsProyecto);
                const usaLogicaMasBosquesProyecto = esProyectoMasBosques(projectName);
                const rangoMasBosquesProyecto = usaLogicaMasBosquesProyecto
                    ? obtenerRangoMasBosquesDesdeEnv(CONFIG.ejecucionEspecifica.fechaInicio, CONFIG.ejecucionEspecifica.fechaFin)
                    : null;
                const sprintObjetivoMasBosques = usaLogicaMasBosquesProyecto
                    ? resolverSprintSemanalMasBosques(
                        chunkDetailRows.map(item => ({
                            fechaCreacionRaw: item?.fields?.["System.CreatedDate"] || null,
                            fechaCambioRaw: item?.fields?.["System.ChangedDate"] || null,
                            sprint: item?.fields?.["System.IterationPath"] || ''
                        })),
                        rangoMasBosquesProyecto.fechaInicio,
                        rangoMasBosquesProyecto.fechaFin,
                        parsearFechaReporte,
                        obtenerSprintLimpioReporte
                    )
                    : '';
               
                const elementosHijos = [];
                const husDict = {};

                chunkDetailRows.forEach(item => {
                    const tipo = (item.fields["System.WorkItemType"] || "").trim();
                    if (tipo === "Feature") return;

                    const esRelevanteBase = cumpleRangoFechasOSprintEspecifico(item);
                    const fechaCreacionItem = parsearFechaReporte(item?.fields?.["System.CreatedDate"] || null);
                    const fechaCambioItem = parsearFechaReporte(item?.fields?.["System.ChangedDate"] || null);
                    const enRangoMasBosquesPorFecha = usaLogicaMasBosquesProyecto && (
                        (fechaCreacionItem && fechaCreacionItem >= rangoMasBosquesProyecto.fechaInicio && fechaCreacionItem <= rangoMasBosquesProyecto.fechaFin)
                        || (fechaCambioItem && fechaCambioItem >= rangoMasBosquesProyecto.fechaInicio && fechaCambioItem <= rangoMasBosquesProyecto.fechaFin)
                    );
                    const sprintItem = obtenerSprintLimpioReporte(item?.fields?.["System.IterationPath"] || '');
                    const enSprintObjetivoMasBosques = usaLogicaMasBosquesProyecto
                        && normalizarSprint(sprintItem) === normalizarSprint(sprintObjetivoMasBosques);
                    const esRelevante = usaLogicaMasBosquesProyecto
                        ? (esRelevanteBase || enSprintObjetivoMasBosques || enRangoMasBosquesPorFecha)
                        : esRelevanteBase;

                    const fechaCambioRaw = item.fields["System.ChangedDate"] || item.fields["System.CreatedDate"];
                    const fechaCreacionRaw = item.fields["System.CreatedDate"] || null;
                    const fechaCambioRealRaw = item.fields["System.ChangedDate"] || null;
                    const ultimaMod = fechaCambioRaw ? new Date(fechaCambioRaw).toLocaleString('es-CO', { timeZone: 'America/Bogota' }) : "No registrada";

                    if (tipo === "Product Backlog Item" || tipo === "Bug") {
                        husDict[item.id] = {
                            id: item.id,
                            titulo: item.fields["System.Title"] || "Sin título",
                            tipo: tipo,
                            estado: item.fields["System.State"] || "Sin estado",
                            sprint: item.fields["System.IterationPath"] || "Sin Iteración",
                            encargado: obtenerNombreEncargado(item.fields["System.AssignedTo"]),
                            totalEstimado: 0,
                            totalEjecutado: 0,
                            cantTareas: 0,
                            cantBugs: 0,
                            ultimaModificacion: ultimaMod,
                            fechaCreacionRaw: fechaCreacionRaw,
                            fechaCambioRaw: fechaCambioRealRaw,
                            fechaRaw: fechaCambioRaw,
                            relevante: esRelevante,
                            subElementos: []
                        };
                        
                        desgloseEstadosPorHU[item.id] = {
                            id: item.id,
                            titulo: item.fields["System.Title"] || "Sin título",
                            proyecto: projectName,
                            conteoEstados: {}
                        };
                    } else {
                        elementosHijos.push({ item, ultimaMod, esRelevante });
                    }
                });

                elementosHijos.forEach(hijoObj => {
                    const hijo = hijoObj.item;
                    const ultimaModHijo = hijoObj.ultimaMod;
                    const esRelevanteHijo = hijoObj.esRelevante;
                    let idPadre = null;

                    if (hijo.relations) {
                        const relacionPadre = hijo.relations.find(r =>
                            r.rel === "System.LinkTypes.Hierarchy-Reverse" || r.attributes?.name === "Parent"
                        );
                        if (relacionPadre && relacionPadre.url) {
                            const urlPartes = relacionPadre.url.split('/');
                            idPadre = parseInt(urlPartes[urlPartes.length - 1], 10);
                        }
                    }

                    if (!idPadre && hijo.fields["System.Parent"]) {
                        idPadre = parseInt(hijo.fields["System.Parent"], 10);
                    }

                    if (idPadre && husDict[idPadre] && esRelevanteHijo) {
                        const tipoHijo = (hijo.fields["System.WorkItemType"] || "").trim();
                        const est = parseFloat(hijo.fields[CAMPOS_HORAS.estimacion]) || 0;
                        const ejec = parseFloat(hijo.fields[CAMPOS_HORAS.ejecucion]) || 0;
                        const encargadoHijo = obtenerNombreEncargado(hijo.fields["System.AssignedTo"]);
                        const estadoHijoRaw = (hijo.fields["System.State"] || "To Do").trim().toLowerCase();

                        husDict[idPadre].relevante = true;

                        if (["done", "completed", "cerrado", "resolved", "closed", "finalizado"].includes(estadoHijoRaw)) {
                            taskDone++;
                        } else if (
                            [
                                "active", 
                                "active/desarrollo", 
                                "desarrollo", 
                                "in progress", 
                                "en progreso", 
                                "asignado", 
                                "doing", 
                                "en desarrollo", 
                                "en pruebas", 
                                "testing"
                            ].includes(estadoHijoRaw) || 
                            estadoHijoRaw.includes("progress") || 
                            estadoHijoRaw.includes("progreso") || 
                            estadoHijoRaw.includes("active")
                        ) {
                            taskInProgress++;
                        } else {
                            taskToDo++;
                        }

                        husDict[idPadre].totalEstimado += est;
                        husDict[idPadre].totalEjecutado += ejec;
                        globales.totalEst += est;
                        globales.totalEjec += ejec;
                        horasEstimadasProyecto += est;
                        horasEjecutadasProyecto += ejec;

                        const estadoHijoBonito = hijo.fields["System.State"] || "To Do";
                        if (desgloseEstadosPorHU[idPadre]) {
                            if (!desgloseEstadosPorHU[idPadre].conteoEstados[estadoHijoBonito]) {
                                desgloseEstadosPorHU[idPadre].conteoEstados[estadoHijoBonito] = 0;
                            }
                            desgloseEstadosPorHU[idPadre].conteoEstados[estadoHijoBonito]++;
                        }

                        let registroEncargado = encargadosPorProyectoAcumulador.find(
                            r => r.Proyecto === projectName && r["Persona Encargada"] === encargadoHijo
                        );
                        if (!registroEncargado) {
                            registroEncargado = { Proyecto: projectName, "Persona Encargada": encargadoHijo, Horas: 0 };
                            encargadosPorProyectoAcumulador.push(registroEncargado);
                        }
                        registroEncargado.Horas += ejec;

                        if (tipoHijo === "Bug") {
                            husDict[idPadre].cantBugs += 1;
                            bugsProyecto++;
                        } else {
                            husDict[idPadre].cantTareas += 1;
                            tasksProyecto++;
                        }

                        let progresoHijo = est > 0 ? (ejec / est) * 100 : 0;
                        if (["Done", "Completed", "Cerrado", "Resolved"].includes(hijo.fields["System.State"])) {
                            progresoHijo = 100;
                        }

                        const fechaHijoRaw = hijo.fields["System.ChangedDate"] || hijo.fields["System.CreatedDate"];
                        const fechaHijoCreacionRaw = hijo.fields["System.CreatedDate"] || null;
                        const fechaHijoCambioRaw = hijo.fields["System.ChangedDate"] || null;
                        const sprintHijo = obtenerSprintLimpioReporte(hijo.fields["System.IterationPath"] || husDict[idPadre].sprint);

                        husDict[idPadre].subElementos.push({
                            huId: idPadre,
                            tareaId: hijo.id,
                            tituloPadre: husDict[idPadre].titulo,
                            titulo: (hijo.fields["System.Title"] || "Sin título").replace(/[,;"'\n\r]/g, " ").trim(),
                            tipo: tipoHijo,
                            estado: estadoHijoBonito,
                            encargado: encargadoHijo,
                            horasEstimadas: est.toFixed(1),
                            horasEjecutadas: ejec.toFixed(1),
                            porcentajeAvance: progresoHijo.toFixed(1),
                            ultimaModificacion: ultimaModHijo,
                            fechaCreacionRaw: fechaHijoCreacionRaw,
                            fechaCambioRaw: fechaHijoCambioRaw,
                            sprint: sprintHijo,
                            fechaRaw: fechaHijoRaw,
                            esSubtarea: true
                        });
                    }
                });

                Object.keys(husDict).forEach(idPadre => {
                    const hu = husDict[idPadre];
                    if (!hu.relevante) return;

                    const estadoPadreRaw = normalizarTextoReporte(hu.estado);
                    if (["done", "completed", "cerrado", "resolved", "closed", "finalizado"].includes(estadoPadreRaw)) {
                        huDone++;
                    } else if (["active", "in progress", "en progreso", "asignado", "doing", "desarrollo", "active/desarrollo"].includes(estadoPadreRaw)) {
                        huInProgress++;
                    } else {
                        huToDo++;
                    }
                    husProyecto++;

                    let progresoHU = hu.totalEstimado > 0 ? (hu.totalEjecutado / hu.totalEstimado) * 100 : 0;
                    if (["Done", "Completed", "Cerrado", "Resolved"].includes(hu.estado)) progresoHU = 100;

                    const nombreSprintLimpio = obtenerSprintLimpioReporte(hu.sprint);

                    rowsConsolidadas.push({
                        proyectoNombre: projectName,
                        sprint: nombreSprintLimpio,
                        huId: hu.id,
                        tareaId: "", 
                        tituloPadre: '',
                        titulo: hu.titulo.replace(/[,;"'\n\r]/g, " ").trim(),
                        tipo: hu.tipo,
                        estado: hu.estado,
                        encargado: hu.encargado,
                        cantTareas: hu.cantTareas,
                        cantBugs: hu.cantBugs,
                        horasEstimadas: hu.totalEstimado.toFixed(1),
                        horasEjecutadas: hu.totalEjecutado.toFixed(1),
                        porcentajeAvance: progresoHU.toFixed(1),
                        ultimaModificacion: hu.ultimaModificacion,
                        fechaCreacionRaw: hu.fechaCreacionRaw,
                        fechaCambioRaw: hu.fechaCambioRaw,
                        fechaRaw: hu.fechaRaw, 
                        esSubtarea: false
                    });

                    hu.subElementos.forEach(hijo => {
                        rowsConsolidadas.push({
                            proyectoNombre: projectName,
                            sprint: hijo.sprint,
                            huId: hijo.huId,
                            tareaId: hijo.tareaId, 
                            tituloPadre: hijo.tituloPadre,
                            titulo: hijo.titulo,
                            tipo: hijo.tipo,
                            estado: hijo.estado,
                            encargado: hijo.encargado,
                            cantTareas: 0,
                            cantBugs: 0,
                            horasEstimadas: hijo.horasEstimadas,
                            horasEjecutadas: hijo.horasEjecutadas,
                            porcentajeAvance: hijo.porcentajeAvance,
                            ultimaModificacion: hijo.ultimaModificacion,
                            fechaCreacionRaw: hijo.fechaCreacionRaw,
                            fechaCambioRaw: hijo.fechaCambioRaw,
                            fechaRaw: hijo.fechaRaw, 
                            esSubtarea: true
                        });
                    });
                });

                logsConsolaProyectos.push({
                    "Proyecto": projectName,
                    "Total HUs": husProyecto,
                    "HU: To Do": huToDo,
                    "HU: In Prog": huInProgress,
                    "HU: Done": huDone,
                    "Total Tasks": tasksProyecto,
                    "Task: To Do": taskToDo,
                    "Task: In Prog": taskInProgress,
                    "Task: Done": taskDone,
                    "Hrs Estim.": `${horasEstimadasProyecto.toFixed(1)}h`,
                    "Hrs Ejecut.": `${horasEjecutadasProyecto.toFixed(1)}h`
                });

            } catch (errProyecto) {
                console.error(`Error procesando proyecto ${projectName}:`, errProyecto.message);
            }
        }

        // AUDITORÍA
        globales.totalHUs = 0;
        globales.totalTasks = 0;
        globales.totalBugsHijos = 0;
        globales.totalSoporte = 0;

        let tablaSoporteConsola = [];

        rowsConsolidadas.forEach(r => {
            const esSoporte = /soporte/i.test(r.titulo);
            if (esSoporte) {
                globales.totalSoporte++;
                tablaSoporteConsola.push({
                    "Proyecto": r.proyectoNombre,
                    "ID": r.tareaId ? `#${r.tareaId}` : `#${r.huId}`,
                    "Título / Actividad (Soporte)": r.titulo.substring(0, 45) + (r.titulo.length > 45 ? "..." : ""),
                    "Tipo": r.tipo,
                    "Estado": r.estado,
                    "Encargado": r.encargado,
                    "Hrs Ejecut.": `${r.horasEjecutadas}h`
                });
            }

            const _mesesNombres = ["Enero", "Febrero", "Marzo", "Abril", "Mayo", "Junio", "Julio", "Agosto", "Septiembre", "Octubre", "Noviembre", "Diciembre"];
            let mesTexto = "Sin Período";
            if (r.sprint) {
                // Parse S-YYYY-MM.NN format by numeric month (highest priority)
                const sprintNum = r.sprint.match(/(\d{4})-(\d{1,2})[.\-]/);
                if (sprintNum) {
                    const yr = Number(sprintNum[1]);
                    const mo = Number(sprintNum[2]);
                    if (mo >= 1 && mo <= 12) mesTexto = `${_mesesNombres[mo - 1]} ${yr}`;
                } else {
                    const s = r.sprint.toLowerCase();
                    if (s.includes("may")) mesTexto = "Mayo 2026";
                    else if (s.includes("jun")) mesTexto = "Junio 2026";
                    else if (s.includes("jul")) mesTexto = "Julio 2026";
                    else if (s.includes("ago") || s.includes("aug")) mesTexto = "Agosto 2026";
                    else if (s.includes("sep")) mesTexto = "Septiembre 2026";
                    else if (s.includes("oct")) mesTexto = "Octubre 2026";
                    else if (s.includes("nov")) mesTexto = "Noviembre 2026";
                    else if (s.includes("dic") || s.includes("dec")) mesTexto = "Diciembre 2026";
                    else if (s.includes("ene") || s.includes("jan")) mesTexto = "Enero 2026";
                    else if (s.includes("feb")) mesTexto = "Febrero 2026";
                    else if (s.includes("mar")) mesTexto = "Marzo 2026";
                    else if (s.includes("abr") || s.includes("apr")) mesTexto = "Abril 2026";
                    else if (r.fechaRaw) {
                        const d = new Date(r.fechaRaw);
                        mesTexto = `${_mesesNombres[d.getMonth()]} ${d.getFullYear()}`;
                    }
                }
            } else if (r.fechaRaw) {
                const d = new Date(r.fechaRaw);
                mesTexto = `${_mesesNombres[d.getMonth()]} ${d.getFullYear()}`;
            }

            if (!conteoMensual[mesTexto]) conteoMensual[mesTexto] = {};
            if (!conteoMensual[mesTexto][r.proyectoNombre]) {
                conteoMensual[mesTexto][r.proyectoNombre] = { hus: 0, tasks: 0, bugs: 0 };
            }
            
            if (!r.esSubtarea) {
                conteoMensual[mesTexto][r.proyectoNombre].hus++;
                globales.totalHUs++;
            } else {
                if (r.tipo === "Bug") {
                    conteoMensual[mesTexto][r.proyectoNombre].bugs++;
                    globales.totalBugsHijos++;
                } else {
                    conteoMensual[mesTexto][r.proyectoNombre].tasks++;
                    globales.totalTasks++;
                }
            }
        });

        // Escritura del archivo CSV
        let contenidoCSV = "Proyecto,Sprint,HU_ID,Titulo,Tipo,Estado,Encargado,Cant_Tareas,Cant_Bugs,Horas_Estimadas,Horas_Ejecutadas,Progreso,Ultima_Modificacion,Tarea_ID\n";
        rowsConsolidadas.forEach(r => {
            contenidoCSV += `"${r.proyectoNombre}","${r.sprint}",${r.huId},"${r.titulo}","${r.tipo}","${r.estado}","${r.encargado}",${r.cantTareas},${r.cantBugs},${r.horasEstimadas},${r.horasEjecutadas},${r.porcentajeAvance},"${r.ultimaModificacion}","${r.tareaId || ''}"\n`;
        });

        const contenidoHTML = generarHTML(rowsConsolidadas, globales, {
            conteoMensual,
            fechaInicio: CONFIG.ejecucion.fechaInicio,
            fechaFin: CONFIG.ejecucion.fechaFin
        });

        // Guardar en report/tables
        if (!fs.existsSync(CONFIG.outputFolderReports)) {
            fs.mkdirSync(CONFIG.outputFolderReports, { recursive: true });
        }
        fs.writeFileSync(path.join(CONFIG.outputFolderReports, CONFIG.outputFile), contenidoCSV, 'utf-8');
        fs.writeFileSync(path.join(CONFIG.outputFolderReports, CONFIG.outputHtmlFile), contenidoHTML, 'utf-8');

        // Guardar en dashboard-ui/public/data
        if (!fs.existsSync(CONFIG.outputFolderDashboardUI)) {
            fs.mkdirSync(CONFIG.outputFolderDashboardUI, { recursive: true });
        }
        fs.writeFileSync(path.join(CONFIG.outputFolderDashboardUI, CONFIG.outputFile), contenidoCSV, 'utf-8');
        fs.writeFileSync(path.join(CONFIG.outputFolderDashboardUI, CONFIG.outputHtmlFile), contenidoHTML, 'utf-8');

        // Guardar en dashboard-ui/src/data
        if (!fs.existsSync(CONFIG.outputFolderDashboardSrc)) {
            fs.mkdirSync(CONFIG.outputFolderDashboardSrc, { recursive: true });
        }
        fs.writeFileSync(path.join(CONFIG.outputFolderDashboardSrc, CONFIG.outputFile), contenidoCSV, 'utf-8');
        fs.writeFileSync(path.join(CONFIG.outputFolderDashboardSrc, CONFIG.outputHtmlFile), contenidoHTML, 'utf-8');

        // Guardar en SharePoint (si existe)
        if (!fs.existsSync(CONFIG.outputFolderSharePoint)) {
            fs.mkdirSync(CONFIG.outputFolderSharePoint, { recursive: true });
        }
        fs.writeFileSync(path.join(CONFIG.outputFolderSharePoint, CONFIG.outputFile), contenidoCSV, 'utf-8');
        fs.writeFileSync(path.join(CONFIG.outputFolderSharePoint, CONFIG.outputHtmlFile), contenidoHTML, 'utf-8');

        // ===================================================
        // GENERACIÓN AUTOMÁTICA DE INFORMES PDF POR PROYECTO
        // ===================================================
        await generarInformesGerencialesPorProyecto(rowsConsolidadas);

        let tablaEncargadosConsola = encargadosPorProyectoAcumulador.map(r => {
            return {
                "Proyecto": r.Proyecto,
                "Persona Encargada": r["Persona Encargada"],
                "Tiempo Execution": `${r.Horas.toFixed(1)} hrs`
            };
        });

        let tablaEstadosPorHUConsola = Object.keys(desgloseEstadosPorHU).map(idHU => {
            const hu = desgloseEstadosPorHU[idHU];
            let estadosTexto = Object.keys(hu.conteoEstados)
                .map(estado => `${estado}: ${hu.conteoEstados[estado]}`)
                .join(' | ');

            return {
                "Proyecto": hu.proyecto,
                "HU ID": `#${hu.id}`,
                "Título de Historia": hu.titulo.substring(0, 35) + (hu.titulo.length > 35 ? "..." : ""),
                "Estados Sub-Tareas": estadosTexto || "❌ Sin sub-elementos asignados"
            };
        });

        const ordenMeses = {
            "Enero 2026": 1, "Febrero 2026": 2, "Marzo 2026": 3, "Abril 2026": 4,
            "Mayo 2026": 5, "Junio 2026": 6, "Julio 2026": 7, "Agosto 2026": 8,
            "Septiembre 2026": 9, "Octubre 2026": 10, "Noviembre 2026": 11, "Diciembre 2026": 12
        };

        let tablaMensualConsola = [];
        Object.keys(conteoMensual)
            .sort((a, b) => (ordenMeses[a] || 99) - (ordenMeses[b] || 99))
            .forEach(mesKey => {
                Object.keys(conteoMensual[mesKey])
                    .sort()
                    .forEach(projName => {
                        const datos = conteoMensual[mesKey][projName];
                        tablaMensualConsola.push({
                            "Mes / Período": mesKey,
                            "Proyecto": projName,
                            "HUs (Historias)": datos.hus,
                            "Tasks (Tareas)": datos.tasks,
                            "Bugs (Errores)": datos.bugs,
                            "Total Items": datos.hus + datos.tasks + datos.bugs
                        });
                    });
            });

        const avanceMuestreo = globales.totalEst > 0 ? (globales.totalEjec / globales.totalEst) * 100 : 0;

        console.clear();

        console.log("\x1b[36m=======================================================================================================================\x1b[0m");
        console.log("\x1b[1m\x1b[33m                             📝 DESCRIPCIÓN Y ALCANCE DE LOS PROYECTOS\x1b[0m");
        console.log("\x1b[36m=======================================================================================================================\x1b[0m");
        logsConsolaProyectos.forEach(p => {
            const desc = DESCRIPCIONES_PROYECTOS[p["Proyecto"]] || DESCRIPCION_ALCANCE_DEFAULT;
            console.log(`📌 \x1b[1mPROYECTO: ${p["Proyecto"]}\x1b[0m`);
            console.log(`    📄 Descripción: ${desc}\n`);
        });

        console.log("\x1b[36m=======================================================================================================================\x1b[0m");
        console.log("\x1b[1m\x1b[32m                             ESTADOS DE HUs Y TAREAS HIJAS POR PROYECTO\x1b[0m");
        console.log("\x1b[36m=======================================================================================================================\x1b[0m");
        console.table(logsConsolaProyectos); 

        console.log("\n\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.log("\x1b[1m\x1b[31m                             🛠️ DETALLE DE ACTIVIDADES DE SOPORTE CONTABILIZADAS\x1b[0m");
        console.log("\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        if (tablaSoporteConsola.length > 0) {
            console.table(tablaSoporteConsola);
        } else {
            console.log("⚠️ No se identificaron tareas o Historias de Usuario con la palabra 'soporte' en el período evaluado.\n");
        }
        
        console.log("\n\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.log("\x1b[1m\x1b[34m                             TIEMPO DE EJECUCIÓN POR PERSONA ENCARGADA\x1b[0m");
        console.log("\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.table(tablaEncargadosConsola); 

        console.log("\n\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.log("\x1b[1m\x1b[35m                             📋 DETALLE DE SUBTAREAS POR CADA HU PADRE\x1b[0m");
        console.log("\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.table(tablaEstadosPorHUConsola);

        console.log("\n\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.log("\x1b[1m\x1b[33m                             📊 VOLUMEN DE ELEMENTOS POR MES Y PROYECTO\x1b[0m");
        console.log("\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.table(tablaMensualConsola);

        console.log("\x1b[36m-----------------------------------------------------------------------------------------------------------------------\x1b[0m");
        console.log(` \x1b[1mHistorias de Usuario / Bugs Totales (Padres):\x1b[0m    ${globales.totalHUs}`);
        console.log(` \x1b[1mTotal de Tareas (Tasks) Asociadas (Hijas):\x1b[0m      ${globales.totalTasks}`);
        console.log(` \x1b[1mTotal de Bugs de Soporte (Hijos):\x1b[0m                 ${globales.totalBugsHijos}`);
        console.log(` \x1b[1mTotal de Actividades de Soporte Filtradas:\x1b[0m     \x1b[31m${globales.totalSoporte}\x1b[0m`);
        console.log(` \x1b[1mTotal de Horas Estimadas Semanales:\x1b[0m             ${globales.totalEst.toFixed(1)} Hrs`);
        console.log(` \x1b[1mTotal de Horas Ejecutadas Semanales:\x1b[0m            ${globales.totalEjec.toFixed(1)} Hrs`);
        console.log(` \x1b[1mBalance de Horas del Período:\x1b[0m                   ${globales.totalEjec.toFixed(1)} / ${globales.totalEst.toFixed(1)} Hrs`);
        console.log(` \x1b[1mPorcentaje de Avance Macroscópico:\x1b[0m              \x1b[32m${avanceMuestreo.toFixed(1)}%\x1b[0m`);
        console.log("\x1b[36m=======================================================================================================================\x1b[0m");

        const tipoDeRango = esRangoPersonalizado ? "RANGO ESTÁTICO (DEFINIDO EN .ENV)" : "RANGO DINÁMICO POR DEFECTO";
        console.log("\x1b[33m=======================================================================================================================\x1b[0m");
        console.log(` \x1b[1m📅 RANGO DE FECHAS AUDITADAS [${tipoDeRango}]:\x1b[0m`);
        console.log(`    Desde: \x1b[1m${CONFIG.ejecucion.fechaInicio.toLocaleString('es-CO')}\x1b[0m | Hasta: \x1b[1m${CONFIG.ejecucion.fechaFin.toLocaleString('es-CO')}\x1b[0m`);
        console.log("\x1b[33m=======================================================================================================================\x1b[0m\n");

    } catch (error) {
        console.error("Error crítico general:", error);
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

exportarDatosDashboard();