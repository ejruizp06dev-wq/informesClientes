function normalizarProyecto(valor) {
    return String(valor || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .trim();
}

function esProyectoMasBosques(nombreProyecto) {
    const nombre = normalizarProyecto(nombreProyecto);
    return nombre.includes('masbosques') || nombre.includes('mas bosques') || nombre === 'mas_bosques';
}

function parsearFechaEnv(valor, finDelDia = false) {
    if (!valor) return null;

    const partes = String(valor).split('-').map(Number);
    if (partes.length !== 3 || partes.some(Number.isNaN)) return null;

    const [ano, mes, dia] = partes;
    return finDelDia
        ? new Date(ano, mes - 1, dia, 23, 59, 59, 999)
        : new Date(ano, mes - 1, dia, 0, 0, 0, 0);
}

function normalizarSprint(valor) {
    return String(valor || '')
        .normalize('NFD')
        .replace(/[\u0300-\u036f]/g, '')
        .toLowerCase()
        .replace(/\s+/g, ' ')
        .trim();
}

function resolverSprintSemanalMasBosques(items, fechaInicio, fechaFin, parsearFechaReporte, obtenerSprintLimpioReporte) {
    const acumulado = new Map();

    items.forEach(item => {
        const fechaCreacionItem = parsearFechaReporte(item.fechaCreacionRaw);
        const fechaCambioItem = parsearFechaReporte(item.fechaCambioRaw);
        const fechaItem = fechaCambioItem || fechaCreacionItem;
        if (!fechaItem) return;
        if (fechaItem < fechaInicio || fechaItem > fechaFin) return;

        const sprintLimpio = obtenerSprintLimpioReporte(item.sprint || '');
        const claveSprint = normalizarSprint(sprintLimpio);
        if (!claveSprint || claveSprint === 'sin iteracion' || claveSprint === 'sin iteracion') return;

        if (!acumulado.has(claveSprint)) {
            acumulado.set(claveSprint, {
                sprint: sprintLimpio,
                conteo: 0,
                ultimaFecha: fechaItem
            });
        }

        const registro = acumulado.get(claveSprint);
        registro.conteo += 1;
        if (fechaItem > registro.ultimaFecha) {
            registro.ultimaFecha = fechaItem;
        }
    });

    const candidatos = [...acumulado.values()].sort((a, b) => {
        if (b.conteo !== a.conteo) return b.conteo - a.conteo;
        return b.ultimaFecha - a.ultimaFecha;
    });

    return candidatos[0]?.sprint || '';
}

function obtenerRangoMasBosquesDesdeEnv(rangoInicioBase, rangoFinBase, env = process.env) {
    const fechaInicio =
        parsearFechaEnv(env.FECHA_INICIO_MASBOSQUES)
        || parsearFechaEnv(env.FECHA_INICIO_ESPECIFICO_MASBOSQUES)
        || parsearFechaEnv(env.FECHA_INICIO_ESPECIFICO_MASBOSUES)
        || new Date(rangoInicioBase);

    const fechaFin =
        parsearFechaEnv(env.FECHA_FIN_MASBOSQUES, true)
        || parsearFechaEnv(env.FECHA_FIN_ESPECIFICO_MASBOSQUES, true)
        || parsearFechaEnv(env.FECHA_FIN_ESPECIFICO_MASBOSUES, true)
        || new Date(rangoFinBase);

    return {
        fechaInicio,
        fechaFin
    };
}

function obtenerFechaInicioHistorialMasBosques(env = process.env) {
    return parsearFechaEnv(env.FECHA_INICIO_HISTORIAL_MASBOSQUES)
        || parsearFechaEnv(env.FECHA_INICIO_HISTORIAL_MASBOSUES)
        || parsearFechaEnv(env.FECHA_INICIO_HISTORIAL)
        || new Date(2026, 4, 1, 0, 0, 0, 0);
}

module.exports = {
    esProyectoMasBosques,
    obtenerRangoMasBosquesDesdeEnv,
    obtenerFechaInicioHistorialMasBosques,
    normalizarSprint,
    resolverSprintSemanalMasBosques
};
