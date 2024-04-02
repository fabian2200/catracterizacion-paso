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
                    <h4 class="card-title" style="font-weight: bold">Gestión de corregimientos</h4>
                    <hr>
                </div>
                <div class="card-content">
                    <div class="card-body">
                        <button data-toggle="modal" data-target="#modalRegistroMunicipio" class="btn btn-warning"><i class="fas fa-plus"></i> Nuevo Corregimiento </button>
                        <br>
                        <br>
                        <table id="tabla_corregimientos">
                            <thead>
                                <tr>
                                    <th style="text-align: center">Id</th>
                                    <th style="text-align: center">Departamento</th>
                                    <th style="text-align: center">Municipio</th>
                                    <th>Nombre</th>
                                    <th style="text-align: center">Estado</th>
                                    <th style="text-align: center">Opciones</th>
                                </tr>
                            </thead>
                            <tbody>
                                
                                <tr v-for="(item, index) in corregimientos" :key="index" v-if="corregimientos.length > 0">
                                    <td style="text-align: center">{{ item.id }}</td>
                                    <td style="text-align: center">{{ item.nombre_departamento }}</td>
                                    <td style="text-align: center">{{ item.nombre_municipio }}</td>
                                    <td>{{ item.nombre }}</td>
                                    <td style="text-align: center"><span :class="item.estado == 'Activo' ? 'badge badge-primary' : 'badge badge-danger'">{{item.estado}}</span></td>
                                    <td style="display: flex; justify-content: space-around;">
                                        <button @click="editarCorregimiento(item)" class="btn btn-success btn_opciones"><i class="fas fa-pencil-alt"></i></button>
                                        <button @click="elimianrCorregimiento(item)" :class="item.estado == 'Activo' ? 'btn btn-danger btn_opciones' : 'btn btn-warning btn_opciones'"><i :class="item.estado == 'Activo' ? 'fas fa-trash-alt' : 'fas fa-check'"></i></button>
                                    </td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                </div>
            </div>
        </div>

        <div class="modal fade" id="modalRegistroMunicipio" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered" role="document">
                <div class="modal-content">
                <div class="modal-header">
                    <h5 class="modal-title" id="exampleModalCenterTitle">Gestión Municipio</h5>
                    <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                    <span aria-hidden="true">&times;</span>
                    </button>
                </div>
                <div class="modal-body">
                    <div class="row">
                        <div class="col-lg-6">
                            <label for=""><strong>Departamento</strong></label>
                            <fieldset class="form-group position-relative has-icon-left">
                                <select @change="consultarMunicipios()" v-model="corregimiento.departamento" class="form-control form-control-lg mb-1" name="" id="">
                                    <option value="">Seleccione una opción</option>
                                    <option v-for="(item, index) in departamentos" :key="index" :value="item.codigo">{{ item.descripcion }}</option>
                                </select>                                            
                                <div class="form-control-position">
                                    <i style="font-size: 19px" class="fas fa-map-marked-alt"></i>
                                </div>
                            </fieldset>
                        </div>
                        <div class="col-lg-6">
                            <label for=""><strong>Municipio</strong></label>
                            <fieldset class="form-group position-relative has-icon-left">
                                <select v-model="corregimiento.municipio" class="form-control form-control-lg mb-1" name="" id="">
                                    <option value="">Seleccione una opción</option>
                                    <option v-for="(item, index) in municipios" :key="index" :value="item.id">{{ item.descripcion }}</option>
                                </select>                                            
                                <div class="form-control-position">
                                    <i style="font-size: 19px" class="fas fa-city"></i>
                                </div>
                            </fieldset>
                        </div>   
                        <div class="col-lg-12">
                            <fieldset class="form-group position-relative has-icon-left">
                                <input v-model="corregimiento.nombre" style="text-transform: uppercase" type="text"  placeholder="Nombre del municipio" class="form-control form-control-lg mb-1">                                          
                                <div class="form-control-position">
                                    <i style="font-size: 19px" class="fas fa-city"></i>
                                </div>
                            </fieldset>
                        </div> 
                    </div>
                </div>
                <div class="modal-footer">
                    <button type="button" class="btn btn-danger" data-dismiss="modal">Cancelar</button>
                    <button @click="guardarCorregimiento()" type="button" class="btn btn-success">Guardar Datos</button>
                </div>
                </div>
            </div>
            </div>
    </div>
</template>
<script>
import * as corregimientosService from "../../services/corregimientos";
import * as caracterizacionService from "../../services/caracterizacion_service";
import Loading from 'vue3-loading-overlay';
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';

export default {
    components: {
        Loading
    },
    data() {
        return {
            corregimientos: [],
            departamentos: [],
            municipios: [],
            loading: false,
            corregimiento: {
                id: "",
                departamento: "",
                municipio: "",
                nombre: ""
            }
        }
    },
    mounted() {
        this.listarCorregimientos();
        this.consultarDepartamentos();
    },
    methods: {
        async listarCorregimientos(){
            this.loading = true;
            await corregimientosService.listarCorregimientos().then(respuesta => {
                this.corregimientos = respuesta.data;
                this.dataTables();
            });
        },
        dataTables() {
            this.loading = false;
            setTimeout(()=>{
                $("#tabla_corregimientos").DataTable({
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
            }, 50)
            
        },
        async consultarDepartamentos(){
            await caracterizacionService.departamentos().then(respuesta => {
                this.departamentos = respuesta.data;
            });
        },
        async consultarMunicipios(){
            this.municipio = "";
            await caracterizacionService.municipios(this.corregimiento.departamento).then(respuesta => {
                this.municipios = respuesta.data;
            });
        },
        async guardarCorregimiento(){
            if(this.corregimiento.departamento == "" || this.corregimiento.municipio == "" || this.corregimiento.nombre == "" ){
                toastr.error("Todos los campos son obligatorios");
            }else{
                this.loading = true;
                await corregimientosService.guardarCorregimiento(this.corregimiento).then(respuesta => {
                    var response = respuesta.data;
                    this.loading = false;
                    if(response.code == 1){
                        toastr.success(response.mensaje);
                        $('#modalRegistroMunicipio').modal('hide');
                        this.corregimiento = {
                            id: "",
                            departamento: "",
                            municipio: "",
                            nombre: ""
                        };
                        this.listarCorregimientos();
                    }else{
                        Swal.fire({
                            icon: "error",
                            title: "Información Importante",
                            text: response.mensaje,
                        });
                    }
                });
            }
        },
        async editarCorregimiento(item){
            this.corregimiento = item;
            this.consultarMunicipios();
            $('#modalRegistroMunicipio').modal('show');
        },
        async elimianrCorregimiento(item){
            this.loading = true;
            if(item.estado == "Activo"){
                item.estado = "Inactivo";
            }else{
                item.estado = "Activo";
            }

            await corregimientosService.elimianrCorregimiento(item.id, item.estado).then(respuesta => {
                var response = respuesta.data;
                this.loading = false;
                toastr.success(response.mensaje);
            });
        }
    },
}
</script>
<style scoped>
    table {
        border-collapse: collapse;
        width: 100% !important;
    }

    thead {
        font-weight: bold;
    }

    tr {
        border-bottom: 1px solid grey;
    }

    td, th {
        padding: 10px;
        font-size: 18px !important;
    }

    select, option {
        font-size: 14px !important;
    }
</style>