import {http} from "./http_services";

export function guardarInformacionTemporal($data) {
    return http().post(
        '/api/guardar-informacion-personal', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function departamentos() {
    return http().get('/api/departamentos');
}

export function municipios(codigo) {
    return http().get('/api/municipios?codigo='+codigo);
}

export function guardarOrigenIdentidad($data) {
    return http().post(
        '/api/guardar-origen-entidad', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function escolaridad() {
    return http().get('/api/escolaridad');
}

export function guardarEducacion($data) {
    return http().post(
        '/api/guardar-educacion', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function ocupaciones() {
    return http().get('/api/ocupaciones');
}

export function guardarSituacionLaboral($data) {
    return http().post(
        '/api/guardar-situacion-laboral', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function guardarSalud($data) {
    return http().post(
        '/api/guardar-salud', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function guardarCulturaTradiciones($data) {
    return http().post(
        '/api/guardar-cultura-tradiciones', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function guardarViviendaHogar($data) {
    return http().post(
        '/api/guardar-vivienda-hogar', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function jefesHogar() {
    return http().get('/api/jefes-hogar');
}

export function consultarCaracterizados() {
    return http().get('/api/caracterizados');
}

export function consultarCaracterizadosDigitador(idDigitador){
    return http().get('/api/caracterizados-digitador?id_usuario='+idDigitador);
}

export function consultarDatosIndividuo(id) {
    return http().get('/api/datos-individuo?identificacion='+id);
}


export function consolidado(id, pagina) {
    return http().get('/api/consolidado?corregimiento='+id+'&pagina='+pagina);
}

export function paginacionConsolidado(id, pagina) {
    return http().get('/api/paginacion-consolidado?corregimiento='+id+'&pagina='+pagina);
}

export function municipiosConsolidado() {
    return http().get('/api/municipios-consolidado');
}

export function exportarConsolidadoExcel(id, tipo) {
    return http().get('/api/exportar-consolidado?corregimiento='+id+'&tipo_reporte='+tipo);
}

export function consolidadoConcejo(id, pagina) {
    return http().get('/api/consolidado-concejo?concejo='+id+'&pagina='+pagina);
}

export function paginacionConsolidadoConcejo(concejo, pagina) {
    return http().get('/api/paginacion-consolidado-concejo?concejo='+concejo+'&pagina='+pagina);
}


export function exportarConsolidadoExcelConcejo(concejo) {
    return http().get('/api/exportar-consolidado-concejo?concejo='+concejo);
}

export function eliminarCaracterizacion(identificacion) {
    return http().get('/api/eliminar-caracterizacion?identificacion='+identificacion);
}