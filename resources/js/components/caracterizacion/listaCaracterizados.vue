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
                    <h4 class="card-title" style="font-weight: bold">Lista de personas caracterizadas</h4>
                    <hr>
                </div>
                <div class="card-content">
                    <div class="card-body">
                        <router-link to="/caracterizacion"><button class="btn btn-info"><i class="fas fa-user-check"></i> Nueva Caracterización </button></router-link>
                        <router-link style="margin-left: 20px" to="/consolidado"><button class="btn btn-warning"><i class="fas fa-file-excel"></i> Consolidado </button></router-link>
                        <div style="margin-top: 40px">
                            <table id="tablaCaracterizados" style="width: 100%" v-if="loading == false">
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

    <div class="modal fade" id="pestanasFaltantes" tabindex="-1" role="dialog" aria-labelledby="exampleModalLabel" aria-hidden="true">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
            <div class="modal-header">
                <h5 class="modal-title" id="exampleModalLabel">Pestañas faltantes</h5>
                <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                <span aria-hidden="true">&times;</span>
                </button>
            </div>
            <div class="modal-body">
                <h3 id="errores_faltantes"></h3>
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
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';
import Swal from 'sweetalert2';


export default {
    components: {
        Loading
    },
    mounted() {
        this.consultarCaracterizados();
    },
    data() {
        return {
            lista_caracterizados: [],
            loading: true           
        }
    },
    methods: {
        async consultarCaracterizados(){
            this.loading = true;
            try {
                await caracterizacionService.consultarCaracterizados().then(respuesta => {
                    this.lista_caracterizados = respuesta.data;
                    setTimeout(()=>{
                        this.dataTables();
                    }, 10);
                });
            } catch (error) {
                console.log(error);
            }
        },
        dataTables() {
            this.loading = false;
            setTimeout(() => {
                $("#tablaCaracterizados").DataTable({
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
            }, 10);
        },
        mostrarpestanasFaltantes(item_errores){
            $("#pestanasFaltantes").modal("show");
            document.getElementById("errores_faltantes").innerHTML = item_errores.mensaje;
        },
        eliminarCaracterizacion(identificacion){
            Swal.fire({
                title: "¿Esta seguro de eliminar esta caracterización?",
                text: "No podrá revertir esta acción",
                icon: "warning",
                showCancelButton: true,
                confirmButtonColor: "#3085d6",
                cancelButtonColor: "#d33",
                confirmButtonText: "Si, eliminar",
                cancelButtonText: "No, cancelar"
            }).then((result) => {
                if (result.isConfirmed) {
                    this.confirmarEliminarCaracterizacion(identificacion);
                }
            });
        },
        async confirmarEliminarCaracterizacion(identificacion){
            this.loading = true;
            try {
                await caracterizacionService.eliminarCaracterizacion(identificacion).then(respuesta => {
                    location.reload();
                });
            } catch (error) {
                console.log(error);
            }
        }
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
</style>