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
                    <h4 class="card-title" style="font-weight: bold">Lista de personas caracterizadas por digitador</h4>
                    <hr>
                </div>
                <div class="card-content">
                    <div class="card-body">
                        <h3 style="font-weight: bold">Seleccione un digitador para ver su lista de caracterizados</h3>
                        <br>
                        <div class="row">
                            <dic style="margin-bottom: 20px" class="col-lg-3" v-for="(item, index) in usuarios" :key="index">
                                <input  type="radio" :id="'control_CU'+index" name="control_CU1" value='' >
                                <label @click="consultarCaracterizados(item.id)" class="lradio" :for="'control_CU'+index">
                                    <br>
                                    <p style="font-size: 17px">{{ item.nombre }}</p>
                                </label>
                            </dic>    
                        </div>
                        <div style="margin-top: 40px">
                            <table id="tablaCaracterizados_2" style="width: 100%">
                                <thead>
                                    <tr>
                                        <th>Estado <br> registro</th>
                                        <th>#</th>
                                        <th>Identificación</th>
                                        <th>Nombre Completo</th>
                                        <th>Rol</th>
                                        <th>Dirección</th>
                                        <th>F. Caracterización</th>
                                        <th style="text-align: center">opciones</th>
                                    </tr>
                                </thead>
                                <tbody>
                                    <tr v-for="(item, index) in lista_caracterizados" :key="index">
                                        <td><span @click="mostrarpestanasFaltantes(item.estado_registro_data)" :class="item.estado_registro_data.cod_estado == 1 ? 'badge badge-success' : 'badge badge-danger'">{{ item.estado_registro_data.estado }}</span></td>
                                        <td>{{ item.numero_caracterizacion }}</td>
                                        <td>{{ item.tipo_identificacion+" : "+item.identificacion }}</td>
                                        <td>{{ item.nombre_completo }}</td>
                                        <td>{{ item.rol }}</td>
                                        <td>{{ item.direccion }}</td>
                                        <td>{{ item.fecha_caracterizacion }}</td>
                                        <td style="display: flex; justify-content: space-around;">
                                            <router-link :to="{ name: 'editarCaracterizacion', params: { identificacion_editar: item.identificacion } }"><button class="btn btn-success btn_opciones"><i class="fas fa-user-edit"></i></button></router-link>
                                            <button @click="eliminarCaracterizacion(item.identificacion)" class="btn btn-danger btn_opciones"><i class="fas fa-trash-alt"></i></button>
                                        </td>
                                    </tr>
                                </tbody>
                            </table>
                        </div>
                        
                    </div>
                </div>
            </div>
        </div>
    </div>

    <div class="modal fade" id="pestanasFaltantes_2" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="exampleModalLabel">Pestañas faltantes</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <h3 id="errores_faltantes_2"></h3>
            </div>
            <div class="modal-footer">
                <button type="button" class="btn btn-danger" data-dismiss="modal">cerrar</button>
            </div>
            </div>
        </div>
        </div>

</template>
<script>

import * as caracterizacionService from "../../services/caracterizacion_service";
import Loading from 'vue3-loading-overlay';
import * as usuarioService from "../../services/usuario_service";
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';

export default {
    components: {
        Loading
    },
    mounted() {
        this.listarUsuarios();
    },
    data() {
        return {
            lista_caracterizados: [],
            loading: true,
            usuarios: []        
        }
    },
    methods: {
        async listarUsuarios(){
            this.loading = true;
            try {
                await usuarioService.listarUsuarios().then(respuesta => {
                   this.usuarios = respuesta.data;
                   this.loading = false;
                   this.dataTables();
                });
            } catch (error) {
                console.log(error);
            }
        },
        async consultarCaracterizados(idUsuario){
            this.loading = true;
            try {
                await caracterizacionService.consultarCaracterizadosDigitador(idUsuario).then(respuesta => {
                    this.lista_caracterizados = respuesta.data;
                    this.loading = false;
                    this.dataTables();
                });
            } catch (error) {
                console.log(error);
            }
        },
        dataTables() {
            $("#tablaCaracterizados_2").DataTable().destroy();
            setTimeout(()=>{
                $("#tablaCaracterizados_2").DataTable({
                ordering: false,
                    "columnDefs": [
                        { "orderable": false, "targets": 0 },
                        { "orderable": false, "targets": 1 },
                        { "orderable": false, "targets": 2 },
                        { "orderable": false, "targets": 3 },
                        { "orderable": false, "targets": 4 },
                        { "orderable": false, "targets": 5 },
                        { "orderable": false, "targets": 6 },
                        { "orderable": false, "targets": 7 }
                    ],
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
                            "next": ">>",
                            "previous": "<<"
                        },
                    },
                });
            }, 200)
        },
        mostrarpestanasFaltantes(item_errores){
            $("#pestanasFaltantes_2").modal("show");
            document.getElementById("errores_faltantes_2").innerHTML = item_errores.mensaje;
        },
    },
}
</script>
<style scoped>
    table {
        border-collapse: collapse;
    }
    th, td {
        font-size: 16px !important;
        padding: 10px 10px !important;
    }

    .btn_opciones {
        display: inline-block !important;
        text-align: center !important;
        vertical-align: middle !important;
        cursor: pointer !important;
        -webkit-user-select: none !important;
        -moz-user-select: none !important;
        user-select: none !important;
        border: 1px solid transparent !important;
        padding: 9px !important;
        font-size: 11px !important;
        line-height: 1.25 !important;
        border-radius: 0.25rem !important;
        transition: background-color 0.15s ease-in-out, border-color 0.15s ease-in-out, box-shadow 0.15s ease-in-out;
    }

    tr {
        border-bottom: 1px solid lightgrey !important;
    }

    .btn-info {
        background-color: #328691 !important;
        border-color: #328691 !important;
    }

    .badge:hover{
        cursor: pointer;
    }

    .nav-item a.nav-link.active {
        color: #2dcee3 !important;
    }
    .nav-item a.nav-link {
        color: #4e5154 !important;
    }

  
    input[type="radio"] {
        display: none;
        &:not(:disabled) ~ label {
            cursor: pointer;
        }
        &:disabled ~ label {
            color: hsla(150, 5%, 75%, 1);
            border-color: hsla(150, 5%, 75%, 1);
            box-shadow: none;
            cursor: not-allowed;
        }
    }

    input[type="checkbox"] {
        display: none;
        &:not(:disabled) ~ label {
            cursor: pointer;
        }
        &:disabled ~ label {
            color: hsla(150, 5%, 75%, 1);
            border-color: hsla(150, 5%, 75%, 1);
            box-shadow: none;
            cursor: not-allowed;
        }
    }

    label.lradio {
        height: 100%;
        display: block;
        background: white;
        border: 2px solid hsla(150, 75%, 50%, 1);
        border-radius: 20px;
        padding: 1rem;
        margin-bottom: 1rem;
        text-align: center;
        box-shadow: 0px 3px 10px -2px hsla(150, 5%, 65%, 0.5);
        position: relative;
        padding-top: 19px;
        font-size: 14px;
        font-weight: bold;
    }

    .lradio p {
        margin-bottom: 0px !important;
    }

    input[type="radio"]:checked + label {
        background: hsla(150, 75%, 50%, 1);
        color: hsla(215, 0%, 100%, 1);
        box-shadow: 0px 0px 20px hsla(150, 100%, 50%, 0.75);
        &::after {
            color: hsla(215, 5%, 25%, 1);
            font-family: FontAwesome;
            border: 2px solid hsla(150, 75%, 45%, 1);
            content: "\f00c";
            font-size: 24px;
            position: absolute;
            top: -25px;
            left: 50%;
            transform: translateX(-50%);
            height: 50px;
            width: 50px;
            line-height: 50px;
            text-align: center;
            border-radius: 50%;
            background: white;
            box-shadow: 0px 2px 5px -2px hsla(0, 0%, 0%, 0.25);
        }
    }

    input[type="checkbox"]:checked + label {
        background: hsla(150, 75%, 50%, 1);
        color: hsla(215, 0%, 100%, 1);
        box-shadow: 0px 0px 20px hsla(150, 100%, 50%, 0.75);
        &::after {
            color: hsla(215, 5%, 25%, 1);
            font-family: FontAwesome;
            border: 2px solid hsla(150, 75%, 45%, 1);
            content: "\f00c";
            font-size: 24px;
            position: absolute;
            top: -25px;
            left: 50%;
            transform: translateX(-50%);
            height: 50px;
            width: 50px;
            line-height: 50px;
            text-align: center;
            border-radius: 50%;
            background: white;
            box-shadow: 0px 2px 5px -2px hsla(0, 0%, 0%, 0.25);
        }
    }

    .alert-primary {
        color: #004085 !important;
        background-color: #cce5ff !important;
        border-color: #b8daff !important;
    }

    td, th {
        padding: 4px !important;
        font-size: 16px !important;
        cursor: pointer;
    }
</style>