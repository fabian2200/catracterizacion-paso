import {http} from "./http_services";

export function datosDashboard() {
    return http().get('/api/datos-dashboard');
}

export function datosInforme() {
    return http().get('/api/datos-informe');
}

export function exportarExcel(tipo) {
    return http().get('/api/exportar-excel?tipo='+tipo);
}