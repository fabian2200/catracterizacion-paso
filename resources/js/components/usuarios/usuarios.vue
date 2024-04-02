<template lang="">
    <div>
        <loading :active="loading" 
            :can-cancel="true"
            loader="bars" 
            color="#38b4c5"
            :height=100
            :width=200
            :on-cancel="onCancel"
            :is-full-page="true">
        </loading>
        <div class="col-lg-12">
            <div class="card" style="min-height: 400pt">
                <div class="card-header">
                    <h4 class="card-title" style="font-weight: bold">Lista de usuarios</h4>
                    <hr>
                </div>
                <div class="card-content">
                    <div class="card-body">
                        <button  data-toggle="modal" data-target="#modalRegistroUsuario" class="btn btn-info"><i class="fas fa-user-plus"></i> Nuevo Usuario </button>
                        <div style="margin-top: 40px">
                            <table id="tablaUsuarios" style="width: 100%">
                                <thead>
                                    <tr>
                                        <th></th>
                                        <th>Identificación</th>
                                        <th>Nombre</th>
                                        <th>Usuario</th>
                                        <th>Email</th>
                                        <th>Dirección</th>
                                        <th>Celular</th>
                                        <th style="text-align: center">opciones</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr v-for="(item, index) in lista_usuarios" :key="index">
                                        <td><i :style="item.estado == 1 ? 'color: green' : 'color: red'" class="fas fa-dot-circle"></i></td>
                                        <td>{{ item.identificacion }}</td>
                                        <td>{{ item.nombre }}</td>
                                        <td>{{ item.usuario }}</td>
                                        <td>{{ item.email }}</td>
                                        <td>{{ item.direccion }}</td>
                                        <td>{{ item.celular }}</td>
                                        <td style="display: flex; justify-content: space-around;">
                                            <button @click="datosEditar = item" data-toggle="modal" data-target="#modalEditarUsuario"  class="btn btn-success btn_opciones">
                                                <i class="fas fa-user-edit"></i>
                                            </button>
                                            <button @click="mensajeEliminarUsuario(item)" :class="item.estado == 1 ? 'btn btn-danger btn_opciones' : 'btn btn-warning btn_opciones'">
                                                <i :class="item.estado == 1 ? 'fas fa-trash' : 'fas fa-user-check'"></i>
                                            </button>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        
                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade text-left" id="modalRegistroUsuario" tabindex="-1" role="dialog" aria-labelledby="myModalLabel17" aria-hidden="true">
            <div class="modal-dialog modal-lg modal-dialog-centered" role="document">
                <div class="modal-content">
                    <div class="modal-body">
                        <section class="row flexbox-container">
                            <div class="col-12 d-flex align-items-center justify-content-center">
                                <div class="col-lg-12 col-md-8 col-10  p-0">
                                    <div class="card border-grey border-lighten-3 px-1 py-1 m-0" style="box-shadow: none">
                                        <div class="card-header border-0">
                                            <div class="card-title text-center">
                                                <h2>Registro de usuario</h2>
                                            </div>
                                        
                                        </div>
                                        <div class="card-content">
                                            <p class="card-subtitle line-on-side text-muted text-center font-small-3 mx-2 my-1"><span>Ingrese los siguientes datos</span></p>
                                            <div class="card-body">
                                                <form class="form-horizontal" id="formularioRegistroDocente" method="POST" validate>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datosRegistro.identificacion" name="identificacion" placeholder="Ingrese su identificación" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-id-card"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datosRegistro.nombre" name="nombre" placeholder="Ingrese su nombre" required>
                                                                <div class="form-control-position">
                                                                    <i class="feather icon-user"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosRegistro.usuario" type="text" class="form-control" name="usuario" placeholder="Ingrese su usuario" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-at"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosRegistro.celular" type="text" class="form-control" name="celular" placeholder="Ingrese su número  de teléfono" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-phone-volume"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosRegistro.email" type="text" class="form-control" name="email" placeholder="Ingrese un email" required>
                                                                <div class="form-control-position">
                                                                    <i class="feather icon-mail"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosRegistro.password" type="password" class="form-control" name="password" placeholder="Ingrese una contraseña" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-unlock-alt"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-12">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosRegistro.direccion" type="text" class="form-control" name="direccion" placeholder="Ingrese su dirección" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-street-view"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <hr>
                                                   
                                                </form>
                                            </div>
                                            <div class="card-body" style="display: flex; padding: 0.5rem; justify-content: space-between;">
                                                <button style="width: 48%" @click="registroUsuario" class="btn btn-outline-success btn-block"><i class="feather icon-user"></i> Guardar Usuario</button>
                                                <button data-dismiss="modal" aria-label="Close" style="width: 48%; margin: 0" class="btn btn-outline-danger btn-block"><i class="feather icon-x"></i> Cancelar</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </section>
                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade text-left" id="modalEditarUsuario" tabindex="-1" role="dialog" aria-labelledby="myModalLabel17" aria-hidden="true">
            <div class="modal-dialog modal-lg modal-dialog-centered" role="document">
                <div class="modal-content">
                    <div class="modal-body">
                        <section class="row flexbox-container">
                            <div class="col-12 d-flex align-items-center justify-content-center">
                                <div class="col-lg-12 col-md-8 col-10  p-0">
                                    <div class="card border-grey border-lighten-3 px-1 py-1 m-0" style="box-shadow: none">
                                        <div class="card-header border-0">
                                            <div class="card-title text-center">
                                                <h2>Editar usuario</h2>
                                            </div>
                                        
                                        </div>
                                        <div class="card-content">
                                            <p class="card-subtitle line-on-side text-muted text-center font-small-3 mx-2 my-1"><span>Ingrese los siguientes datos</span></p>
                                            <div class="card-body">
                                                <form class="form-horizontal" id="formularioRegistroDocente" method="POST" validate>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datosEditar.identificacion" name="identificacion" placeholder="Ingrese su identificación" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-id-card"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datosEditar.nombre" name="nombre" placeholder="Ingrese su nombre" required>
                                                                <div class="form-control-position">
                                                                    <i class="feather icon-user"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosEditar.usuario" type="text" class="form-control" name="usuario" placeholder="Ingrese su usuario" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-at"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosEditar.celular" type="text" class="form-control" name="celular" placeholder="Ingrese su número  de teléfono" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-phone-volume"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosEditar.email" type="text" class="form-control" name="email" placeholder="Ingrese un email" required>
                                                                <div class="form-control-position">
                                                                    <i class="feather icon-mail"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosEditar.password" type="password" class="form-control" name="password" placeholder="Ingrese una contraseña" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-unlock-alt"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-12">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input v-model="datosEditar.direccion" type="text" class="form-control" name="direccion" placeholder="Ingrese su dirección" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-street-view"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <hr>
                                                </form>
                                            </div>
                                            <div class="card-body" style="display: flex; padding: 0.5rem; justify-content: space-between;">
                                                <button style="width: 48%" @click="editarUsuario" class="btn btn-outline-warning btn-block"><i class="feather icon-user"></i> Editar Usuario</button>
                                                <button data-dismiss="modal" aria-label="Close" style="width: 48%; margin: 0" class="btn btn-outline-danger btn-block"><i class="feather icon-x"></i> Cancelar</button>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                        </section>
                    </div>
                </div>
            </div>
        </div>
    </div>
</template>
<script>
import * as usuarioService from "../../services/usuario_service";
import Swal from 'sweetalert2';
import Loading from 'vue3-loading-overlay';
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';

export default {
    components: {
        Loading
    },
    mounted() {
        this.consultarUsuarios();
    },
    data() {
        return {
            lista_usuarios: [],
            loading: false,
            datosRegistro: {
                identificacion: "",
                nombre: "",
                usuario: "",
                email: "",
                password: "",
                direccion: "",
                celular: ""
            }, 
            datosEditar: {
                identificacion: "",
                nombre: "",
                usuario: "",
                email: "",
                password: "",
                direccion: "",
                celular: ""
            }
        }
    },
    methods: {
        async consultarUsuarios(){
            this.loading = true;
            try {
                await usuarioService.consultarUsuarios().then(respuesta => {
                    this.lista_usuarios = respuesta.data;
                });
                this.dataTables();
            } catch (error) {
                console.log(error);
            }
        },
        dataTables() {
            this.loading = false;
            $("#tablaUsuarios").DataTable({
                ordering: false,
                language: {
                    "decimal": "",
                    "emptyTable": "No hay información",
                    "info": "Mostrando _START_ a _END_ de _TOTAL_ Registros",
                    "infoEmpty": "Mostrando 0 to 0 of 0 Entradas",
                    "infoFiltered": "(Filtrado de _MAX_ total Registros)",
                    "infoPostFix": "",
                    "thousands": ",",
                    "lengthMenu": "Mostrar _MENU_ Registros",
                    "loadingRecords": "Cargando...",
                    "processing": "Procesando...",
                    "search": "Buscar:",
                    "zeroRecords": "Sin resultados encontrados",
                    "paginate": {
                        "first": "Primero",
                        "last": "Ultimo",
                        "next": "Siguiente",
                        "previous": "Anterior"
                    }
                },
            });
        },
        async registroUsuario(){
            this.loading = true;
            try {
                if(this.datosRegistro.direccion != "" && this.datosRegistro.email != "" && this.datosRegistro.identificacion != "" && this.datosRegistro.nombre != "" && this.datosRegistro.password != "" && this.datosRegistro.celular != "" && this.datosRegistro.usuario != ""){
                    await usuarioService.registroUsuario(this.datosRegistro).then(respuesta => {
                        var respuesta_ok = respuesta.data;
                        if(respuesta_ok[1] == 1){
                            toastr.success(respuesta_ok[0]);
                            $('#modalRegistroUsuario').modal('hide');
                            $('#tablaUsuarios').DataTable().destroy();
                            setTimeout(() => {
                                this.consultarUsuarios();
                            }, 100);
                        }else{
                            toastr.error(respuesta_ok[0]);
                        }
                    });
                }else{
                    toastr.warning("Todos los campos son obligatorios");
                }
            } catch (error) {
                console.log(error);
            }
        },
        async editarUsuario(){
            this.loading = true;
            try {
                if(this.datosEditar.direccion != "" && this.datosEditar.email != "" && this.datosEditar.identificacion != "" && this.datosEditar.nombre != "" && this.datosEditar.celular != "" && this.datosEditar.usuario != ""){
                    await usuarioService.editarUsuario(this.datosEditar).then(respuesta => {
                        var respuesta_ok = respuesta.data;
                        if(respuesta_ok[1] == 1){
                            toastr.success(respuesta_ok[0]);
                            $('#modalEditarUsuario').modal('hide');
                            $('#tablaUsuarios').DataTable().destroy();
                            setTimeout(() => {
                                this.consultarUsuarios();
                            }, 100);
                        }else{
                            toastr.error(respuesta_ok[0]);
                        }
                    });
                }else{
                    toastr.warning("Todos los campos son obligatorios");
                }
            } catch (error) {
                console.log(error);
            }
        },
        mensajeEliminarUsuario(item){
            Swal.fire({
                title: '¿Desea cambiar de estado a este usuario?',
                text: "",
                icon: 'warning',
                showCancelButton: true,
                confirmButtonColor: '#3085d6',
                cancelButtonColor: '#d33',
                confirmButtonText: '¡Sí, cambiar!',
                cancelButtonText: 'Cancelar'
            }).then((result) => {
                if (result.isConfirmed) {
                    item.estado = item.estado == 1 ? 0 : 1;
                    this.cambiarEstadoUsuario(item.id, item.estado);
                }
            })
        },
        async cambiarEstadoUsuario(id, estado){
            this.loading = true;
            await usuarioService.cambiarEstadoUsuario(id, estado).then(respuesta => {
                toastr.success(respuesta.data);
                this.loading = false;
            });
        }
    }
    
}
</script>
<style scoped>
    th, td {
        font-size: 18px !important;
        padding: 10px 10px !important;
    }

    .btn_opciones {
        display: inline-block !important;
        text-align: center !important;
        vertical-align: middle !important;
        cursor: pointer !important;
        user-select: none !important;
        border: 1px solid transparent !important;
        padding: 9px !important;
        font-size: 11px !important;
        line-height: 1.25 !important;
        border-radius: 0.25rem !important;
        transition: background-color 0.15s ease-in-out, border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
    }
</style>