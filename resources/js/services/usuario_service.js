import {http} from "./http_services";

export function misDatos() {
    return http().get('/api/mis-datos');
}

export function verificarLogin() {
    return http().get('/api/verificarLogin');
}

export function editar_usuario($data) {
    return http().post(
        '/api/editar-usuario', 
        $data, 
        {
            headers: {
                'Content-Type': 'multipart/form-data',
            }
        }
    );
}

export function cambiar_password($data) {
    return http().post(
        '/api/cambiar-password', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function consultarUsuarios() {
    return http().get('/api/consultar-usuarios');
}

export function registroUsuario($data) {
    return http().post(
        '/api/registrar-usuario', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function editarUsuario($data) {
    return http().post(
        '/api/editar-usuario-admin', 
        $data, 
        {
            headers: {
                'Content-Type': 'application/json'
            }
        }
    );
}

export function cambiarEstadoUsuario(id, estado) {
    return http().get('/api/cambiar-estado-usuario?id='+id+"&estado_actual="+estado);
}

export function cerrarSesion() {
    return http().get('/api/cerrar-sesion');
}

export function listarUsuarios() {
    return http().get('/api/listar-usuarios');
}