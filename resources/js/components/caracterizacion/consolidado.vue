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
                    <h4 class="card-title" style="font-weight: bold">Consolidado personas censadas</h4>
                    <button v-if="corregimiento != ''" @click="exportarConsolidadoExcelPregunta" class="btn btn-success boton_exportar"><i class="fas fa-file-excel"></i></button>
                    <button v-if="corregimiento != ''" @click="corregimiento = ''" class="btn btn-danger boton_atras"><i class="fas fa-arrow-left"></i></button>
                    <button v-if="concejo != ''" @click="exportarConsolidadoExcelConcejo" class="btn btn-success boton_exportar"><i class="fas fa-file-excel"></i></button>
                    <button v-if="concejo != ''" @click="concejo = ''" class="btn btn-danger boton_atras"><i class="fas fa-arrow-left"></i></button>
                    <hr>
                </div>
                <div class="card-content">
                    <div class="card-body">
                        <div v-if="corregimiento == '' && concejo == ''">
                            <div class="row">
                                <div class="col-lg-12">
                                    <h2 style="font-weight: bold; color: #008385;">SELECCIONA UNA OPCIÓN PARA GENERAR EL CONSOLIDADO</h2>
                                    <br>
                                </div>
                                <div class="col-lg-12">
                                    <h3 style="font-weight: bold; color: #ff425c;">POR CORREGIMIENTO</h3>
                                    <br>
                                </div>
                                <div style="text-transform: uppercase;" v-for="(item, index) in municipios_consolidados" :key="index" class="col-lg-4">
                                    <div @click="verConsolidado(item.id_corregimiento, item.nombre, 1)" class="card card_hover">
                                        <div class="card-content">
                                            <div class="media align-items-stretch">
                                                <div class="p-2 text-center bg-primary bg-darken-2">
                                                    <i class="icon-user font-large-2 white"></i>
                                                    <hr>
                                                    <h5 style="font-weight: bold !important" class="white mb-0">{{item.cantidad}}</h5>
                                                </div>
                                                <div style="display: flex; flex-direction: column; justify-content: center;" class="p-1 bg-primary white media-body">
                                                    <h5><strong>DEP:</strong> {{item.departamento}}</h5>
                                                    <h5><strong>MUN:</strong> {{item.municipio}}</h5>
                                                    <h5 class="text-bold-400 mb-0"><strong>CORR:</strong> {{item.nombre}}</h5>
                                                </div>
                                            </div>
                                        </div>
                                    </div>                               
                                </div>

                                <div style="text-transform: uppercase;" class="col-lg-4">
                                    <div @click="verConsolidado('todo', 'General', 1)" class="card card_hover">
                                        <div class="card-content">
                                            <div class="media align-items-stretch" style="height: 134px">
                                                <div style="display: flex; align-items: center; background-color: #ff864a !important;" class="p-2 text-center bg-primary bg-darken-2">
                                                    <i class="icon-users font-large-2 white"></i>
                                                </div>
                                                <div style="background-color: #ffa87d !important; display: flex; flex-direction: column; justify-content: center;" class="p-1 bg-primary white media-body">
                                                    <h5><strong>Consolidado <br> General</strong> </h5>
                                                </div>
                                            </div>
                                        </div>
                                    </div>    
                                </div>
                            </div>
                            <br>
                            <div class="row">
                                <div class="col-lg-12">
                                    <h3 style="font-weight: bold; color: #ff425c;">POR CONCEJO COMUNITARIO</h3>
                                    <br>
                                </div>
                                <div style="text-transform: uppercase;" class="col-lg-4">
                                    <div @click="verConsolidado('todo', 1, 2)" class="card card_hover">
                                        <div class="card-content">
                                            <div class="media align-items-stretch" style="height: 134px">
                                                <div style="display: flex; align-items: center; background-color: #ff425c !important;" class="p-2 text-center bg-primary bg-darken-2">
                                                    <i class="icon-home font-large-2 white"></i>
                                                </div>
                                                <div style="background-color: #ff7588 !important; display: flex; flex-direction: column; justify-content: center;" class="p-1 bg-primary white media-body">
                                                    <h5><strong>CONCEJO COMUNITARIO <br> "ALEJO DURAN"</strong> </h5>
                                                </div>
                                            </div>
                                        </div>
                                    </div>    
                                </div>
                                <div style="text-transform: uppercase;" class="col-lg-4">
                                    <div @click="verConsolidado('todo', 2, 2)" class="card card_hover">
                                        <div class="card-content">
                                            <div class="media align-items-stretch" style="height: 134px">
                                                <div style="display: flex; align-items: center; background-color: #ff425c !important;" class="p-2 text-center bg-primary bg-darken-2">
                                                    <i class="icon-home font-large-2 white"></i>
                                                </div>
                                                <div style="background-color: #ff7588 !important; display: flex; flex-direction: column; justify-content: center;" class="p-1 bg-primary white media-body">
                                                    <h5><strong>CONCEJO COMUNITARIO <br> "JULIO CESAR ALTAMAR MUÑOS"</strong> </h5>
                                                </div>
                                            </div>
                                        </div>
                                    </div>    
                                </div>
                            </div>
                        </div>
                            
                        <div v-if="corregimiento != ''">
                            <div>
                                <table style=" width: 100% !important;">
                                    <thead>
                                        <tr>
                                            <td rowspan="2" style="width: 220px; text-align: center; padding: 0px">
                                                <img style="width: 150px" src="/imagenes/censo.png" alt="">
                                            </td>
                                            <td colspan="2" style="text-align: center">MACROPROCESO: AGENCIA NACIONAL DE TIERRAS <br> PROCESO: DIRECCIÓN DE ASUNTOS ÉTNICOS</td>
                                            <td style="padding-left: 20px">
                                                CODIGO:
                                                <br> 
                                                FECHA: 
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2" style="text-align: center">
                                                FORMATO: CENSO POBLACIÓN COMUNIDADES NEGRAS
                                            </td>
                                            <td style="padding-left: 20px">
                                                PAGINA: 1/1
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="text-align: center" colspan="4">
                                                DATOS DE IDENTIFICACION
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>Departamento: CESAR</td>
                                            <td>Municipio: EL PASO</td>
                                            <td>Corregimiento: {{ corregimiento }}</td>
                                            <td>Fecha: {{currentDate()}}</td>
                                        </tr>
                                    </thead>
                                </table>
                                <br>
                                <div style=" width: 100%; overflow-x: auto">
                                    <table id="table2"  style=" width: 100% !important;">
                                        <thead>
                                            <tr>
                                                <td style="text-align: center; background-color: #8cee8c; font-weight: bold;" colspan="2">Listado de personas</td>
                                                <td style="text-align: center; background-color: #43d7f0; font-weight: bold;" colspan="3">Datos</td>
                                                <td style="text-align: center; background-color: #cdf043; font-weight: bold;" colspan="6"></td>
                                                <td style="text-align: center; background-color: #43d7f0; font-weight: bold;" colspan="2"></td>
                                                <td style="text-align: center; background-color: #8cee8c; font-weight: bold;" colspan="4">Tenencia de Tierras</td>
                                            </tr>
                                            <tr>
                                                <td class="vertical-text" style="text-align: center;  background-color: #8cee8c; font-weight: bold;">Nº Familia</td>
                                                <td class="vertical-text" style="text-align: center;  background-color: #8cee8c; font-weight: bold;">Nº Orden</td>
                                                <td style="text-align: center; background-color: #43d7f0; font-weight: bold;">Nombre</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;">Tipo</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;"># Identificación</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Sexo</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Edad</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Estado Civil</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Parentesco</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Ocupación</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Escolaridad</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;">Tipo Vivienda</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;">Tenencia Vivienda</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Posesión <br> Baldíos</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Área Total (ha)</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Área Total (m2)</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Área Utilizada (m2)</td>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <tr style="background-color: #f5f2ee;" v-for="(item, index) in consolidado" :key="index">
                                                <td style="text-align: center" v-if="index == 0 || consolidado[index].numero_familia != consolidado[index-1].numero_familia" :rowspan="item.rowspan">{{ item.numero_familia }}</td>
                                                <td style="text-align: center" >{{ item.numero_orden }}</td>
                                                <td>{{ item.nombre_completo }}</td>
                                                <td>{{ item.tipo_identificacion }}</td>
                                                <td>{{ item.identificacion }}</td>
                                                <td>{{ item.sexo == "Masculino" ? "M" : "F" }}</td>
                                                <td></td>
                                                <td>{{ item.estado_civil }}</td>
                                                <td>{{ item.rol }}</td>
                                                <td>{{ item.ocupacion }}</td>
                                                <td  style="text-align: center">{{ item.escolaridad }}</td>
                                                <td class="vertical-text" v-if="index == 0 || consolidado[index].numero_familia != consolidado[index-1].numero_familia" :rowspan="item.rowspan" style="text-align: center">{{ item.rol == "Jefe de hogar" ? item.tipo_vivienda : "" }}</td>
                                                <td class="vertical-text" v-if="index == 0 || consolidado[index].numero_familia != consolidado[index-1].numero_familia" :rowspan="item.rowspan" style="text-align: center">{{ item.rol == "Jefe de hogar" ? item.tenencia : "" }}</td>
                                                <td v-if="index == 0 || consolidado[index].numero_familia != consolidado[index-1].numero_familia" :rowspan="item.rowspan" style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.tipo_vivienda == 'Finca' || item.tipo_vivienda == 'Parcela') ? item.posecion_baldia : "N.A" }}</td>
                                                <td v-if="index == 0 || consolidado[index].numero_familia != consolidado[index-1].numero_familia" :rowspan="item.rowspan" style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.tipo_vivienda == 'Finca' || item.tipo_vivienda == 'Parcela') ? (item.area_total / 10000) : "N.A" }}</td>
                                                <td v-if="index == 0 || consolidado[index].numero_familia != consolidado[index-1].numero_familia" :rowspan="item.rowspan" style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.tipo_vivienda == 'Finca' || item.tipo_vivienda == 'Parcela') ? item.area_total : "N.A" }}</td>
                                                <td v-if="index == 0 || consolidado[index].numero_familia != consolidado[index-1].numero_familia" :rowspan="item.rowspan" style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.tipo_vivienda == 'Finca' || item.tipo_vivienda == 'Parcela') ? item.area_destinada_ocupada : "N.A" }}</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                                <p class="px-1"></p>
                                <br>
                                <div class="row pagination" v-if="paginacion.numero_registros > 0">
                                    <div class="col-lg-12" style="display: flex; justify-content: center; align-items: center" >
                                        <ul class="pagination pagination_style">
                                            <!-- flecha anterior -->
                                            <li class="page-item">
                                                <button class="page-link banterior" @click="decrementarPagina" aria-label="Next">
                                                    <i class="fa fa-chevron-left"></i> Anterior
                                                </button>
                                            </li>

                                            <!-- mostrar página 1 y los ... si la primera del grupo no es la página 1 -->
                                            <li v-if="primera !== 1" class="page-item">
                                                <button class="page-link" @click="irAPagina(1)" aria-label="Previous">
                                                <span aria-hidden="true">1</span>
                                                </button>
                                            </li>
                                            <li v-if="primera !== 1" class="page-item disabled">
                                                <button class="page-link" disabled aria-label="Previous">
                                                <span aria-hidden="true">...</span>
                                                </button>
                                            </li>

                                            <!-- mostrar página actual, las 5 anteriores y las 5 posteriores -->
                                            <li v-for="i in ultima - primera + 1" :key="i" :class="{ 'page-item': true, 'active': pagina === i + primera - 1 }">
                                                <button class="page-link" @click="irAPagina(i + primera - 1)" aria-label="Previous">
                                                <span aria-hidden="true">{{ i + primera - 1 }}</span>
                                                </button>
                                            </li>

                                            <!-- mostrar la última y los ... si la última del grupo no es la última -->
                                            <li v-if="ultima !== paginacion.paginas_consulta" class="page-item disabled">
                                                <button class="page-link" disabled aria-label="Previous">
                                                <span aria-hidden="true">...</span>
                                                </button>
                                            </li>
                                            <li v-if="ultima !== paginacion.paginas_consulta" class="page-item">
                                                <button class="page-link" @click="irAPagina(paginacion.paginas_consulta)" aria-label="Previous">
                                                <span aria-hidden="true">{{ paginacion.paginas_consulta }}</span>
                                                </button>
                                            </li>

                                            <!-- flecha siguiente -->
                                            <li class="page-item">
                                                <button class="page-link bsiguiente" @click="incrementarPagina" aria-label="Next">
                                                    Siguiente <i class="fa fa-chevron-right"></i> 
                                                </button>
                                            </li>
                                            </ul>
                                    </div>
                                </div>
                            </div>
                        </div>

                        <div v-if="concejo != ''">
                            <div>
                                <table style=" width: 100% !important;">
                                    <thead>
                                        <tr>
                                            <td rowspan="2" style="width: 220px; text-align: center; padding: 0px">
                                                <img style="width: 150px" src="/imagenes/censo.png" alt="">
                                            </td>
                                            <td colspan="2" style="text-align: center">MACROPROCESO: AGENCIA NACIONAL DE TIERRAS <br> PROCESO: DIRECCIÓN DE ASUNTOS ÉTNICOS</td>
                                            <td style="padding-left: 20px">
                                                CODIGO:
                                                <br> 
                                                FECHA: 
                                            </td>
                                        </tr>
                                        <tr>
                                            <td colspan="2" style="text-align: center">
                                                FORMATO: CENSO POBLACIÓN COMUNIDADES NEGRAS
                                            </td>
                                            <td style="padding-left: 20px">
                                                PAGINA: 1/1
                                            </td>
                                        </tr>
                                        <tr>
                                            <td style="text-align: center" colspan="4">
                                                DATOS DE IDENTIFICACION
                                            </td>
                                        </tr>
                                        <tr>
                                            <td>Departamento: CESAR</td>
                                            <td>Municipio: EL PASO</td>
                                            <td>Concejo: {{ concejo }}</td>
                                            <td>Fecha: {{currentDate()}}</td>
                                        </tr>
                                    </thead>
                                </table>
                                <br>
                                <div style=" width: 100%; overflow-x: auto">
                                    <table id="table2"  style=" width: 100% !important;">
                                        <thead>
                                            <tr>
                                                <td style="text-align: center; background-color: #43d7f0; font-weight: bold;" colspan="4">Datos</td>
                                                <td style="text-align: center; background-color: #cdf043; font-weight: bold;" colspan="6"></td>
                                                <td style="text-align: center; background-color: #43d7f0; font-weight: bold;" colspan="2"></td>
                                                <td style="text-align: center; background-color: #8cee8c; font-weight: bold;" colspan="4">Tenencia de Tierras</td>
                                            </tr>
                                            <tr>
                                                <td  class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;">N° Familia</td>
                                                <td style="text-align: center; background-color: #43d7f0; font-weight: bold;">Nombre</td>
                                                <td style="text-align: center; background-color: #43d7f0; font-weight: bold;">Tipo</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;">N° Identificación</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Sexo</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Edad (Años)</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Estado Civil</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Parentesco</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Ocupación</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #cdf043; font-weight: bold;">Escolaridad</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;">Tipo Vivienda</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #43d7f0; font-weight: bold;">Tenencia Vivienda</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Posesión <br> Baldíos</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Área Total (ha)</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Área Total (m2)</td>
                                                <td class="vertical-text" style="text-align: center; background-color: #8cee8c; font-weight: bold;">Área Utilizada (m2)</td>
                                            </tr>
                                        </thead>
                                        <tbody>
                                            <tr style="background-color: #f5f2ee;" v-for="(item, index) in consolidado" :key="index">
                                                <td style="text-align: center">{{ item.numero_familia }}</td>
                                                <td>{{ item.nombre_completo }}</td>
                                                <td>{{ item.tipo_identificacion }}</td>
                                                <td>{{ item.identificacion }}</td>
                                                <td>{{ item.sexo == "Masculino" ? "M" : "F" }}</td>
                                                <td>{{ item.edad }}</td>
                                                <td>{{ item.estado_civil }}</td>
                                                <td>{{ item.rol }}</td>
                                                <td>{{ item.ocupacion }}</td>
                                                <td style="text-align: center">{{ item.escolaridad }}</td>
                                                <td class="vertical-text">{{ item.rol == "Jefe de hogar" ? item.vivienda_hogar.tipo_vivienda : "" }}</td>
                                                <td class="vertical-text">{{ item.rol == "Jefe de hogar" ? item.vivienda_hogar.tenencia : "" }}</td>
                                                <td style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.vivienda_hogar.tipo_vivienda == 'Finca' || item.vivienda_hogar.tipo_vivienda == 'Parcela') ? item.vivienda_hogar.posecion_baldia : "N.A" }}</td>
                                                <td style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.vivienda_hogar.tipo_vivienda == 'Finca' || item.vivienda_hogar.tipo_vivienda == 'Parcela') ? (item.vivienda_hogar.area_total / 10000) : "N.A" }}</td>
                                                <td style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.vivienda_hogar.tipo_vivienda == 'Finca' || item.vivienda_hogar.tipo_vivienda == 'Parcela') ? item.vivienda_hogar.area_total : "N.A" }}</td>
                                                <td style="text-align: center">{{ item.rol == "Jefe de hogar" && (item.vivienda_hogar.tipo_vivienda == 'Finca' || item.vivienda_hogar.tipo_vivienda == 'Parcela') ? item.area_destinada_ocupada : "N.A" }}</td>
                                            </tr>
                                        </tbody>
                                    </table>
                                </div>
                                <p class="px-1"></p>
                                <br>
                                <div class="row pagination" v-if="paginacion.numero_registros > 0">
                                    <div class="col-lg-12" style="display: flex; justify-content: center; align-items: center" >
                                        <ul class="pagination pagination_style">
                                            <!-- flecha anterior -->
                                            <li class="page-item">
                                                <button class="page-link banterior" @click="decrementarPaginaConcejo" aria-label="Next">
                                                    <i class="fa fa-chevron-left"></i> Anterior
                                                </button>
                                            </li>

                                            <!-- mostrar página 1 y los ... si la primera del grupo no es la página 1 -->
                                            <li v-if="primera !== 1" class="page-item">
                                                <button class="page-link" @click="irAPaginaConcejo(1)" aria-label="Previous">
                                                <span aria-hidden="true">1</span>
                                                </button>
                                            </li>
                                            <li v-if="primera !== 1" class="page-item disabled">
                                                <button class="page-link" disabled aria-label="Previous">
                                                <span aria-hidden="true">...</span>
                                                </button>
                                            </li>

                                            <!-- mostrar página actual, las 5 anteriores y las 5 posteriores -->
                                            <li v-for="i in ultima - primera + 1" :key="i" :class="{ 'page-item': true, 'active': pagina === i + primera - 1 }">
                                                <button class="page-link" @click="irAPaginaConcejo(i + primera - 1)" aria-label="Previous">
                                                <span aria-hidden="true">{{ i + primera - 1 }}</span>
                                                </button>
                                            </li>

                                            <!-- mostrar la última y los ... si la última del grupo no es la última -->
                                            <li v-if="ultima !== paginacion.paginas_consulta" class="page-item disabled">
                                                <button class="page-link" disabled aria-label="Previous">
                                                <span aria-hidden="true">...</span>
                                                </button>
                                            </li>
                                            <li v-if="ultima !== paginacion.paginas_consulta" class="page-item">
                                                <button class="page-link" @click="irAPaginaConcejo(paginacion.paginas_consulta)" aria-label="Previous">
                                                <span aria-hidden="true">{{ paginacion.paginas_consulta }}</span>
                                                </button>
                                            </li>

                                            <!-- flecha siguiente -->
                                            <li class="page-item">
                                                <button class="page-link bsiguiente" @click="incrementarPaginaConcejo" aria-label="Next">
                                                    Siguiente <i class="fa fa-chevron-right"></i> 
                                                </button>
                                            </li>
                                            </ul>
                                    </div>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>
    </div>
</template>
<script>
import * as caracterizacionService from "../../services/caracterizacion_service";
import Swal from 'sweetalert2';
import Loading from 'vue3-loading-overlay';
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';

export default {
    components: {
        Loading
    },
    data() {
        return {
            corregimiento: "",
            municipios_consolidados: [],
            consolidado: [],
            loading: false,
            //pagination
            pagina: 1,
            paginacion: [],
            primera: 0,
            ultima: 0,
            idBuscar: "",
            concejo: ""
        }
    },
    mounted() {
        this.municipiosConsolidado();
    },
    methods: {
        currentDate() {
            const current = new Date();
            const date = `${current.getDate()}/${current.getMonth()+1}/${current.getFullYear()}`;
            return date;
        },
        async municipiosConsolidado(){
            this.loading = true;
            await caracterizacionService.municipiosConsolidado().then(respuesta => {
                this.municipios_consolidados = respuesta.data;
                this.loading = false;
            });
        },
        async paginador() {
            await caracterizacionService.paginacionConsolidado(this.idBuscar, this.pagina).then(respuesta => {
                this.paginacion = respuesta.data;
                this.pagina = parseInt(this.pagina);
                this.primera = this.pagina - 4 > 1 ? this.pagina - 4 : 1;
                this.ultima = this.pagina + 4 < this.paginacion.paginas_consulta ? this.pagina + 4 : this.paginacion.paginas_consulta;
            });
        },
        verConsolidado(id, nombre, tipo_consolidado){
            this.loading = true;
            this.pagina = 1;
            if(tipo_consolidado == 1){
                this.corregimiento = nombre;
                this.idBuscar = id;
                this.consolidadoFunction();
            }else{
                if(nombre == 1){
                    this.concejo = 'CONCEJO COMUNITARIO "ALEJO DURAN"';
                }else{
                    this.concejo = 'CONCEJO COMUNITARIO "JULIO CESAR ALTAMAR MUÑOS"';
                }
                
                this.consolidadoFunctionConcejo();
            }
            
        },
        async consolidadoFunction(){
            this.loading = true;
            await caracterizacionService.consolidado(this.idBuscar, this.pagina).then(respuesta => {
                this.consolidado = respuesta.data;
                this.loading = false;
            });
            this.paginador();
        },
        incrementarPagina() {
            if (this.pagina < this.paginacion.paginas_consulta) {
                this.pagina++;
            }
            this.consolidadoFunction();            
        },
        decrementarPagina() {
            if (this.pagina > 1) {
                this.pagina--;
            }
            this.consolidadoFunction();
        },
        irAPagina(pagina) {
            this.pagina = pagina;
            this.consolidadoFunction();
        },
        exportarConsolidadoExcelPregunta(){
            Swal.fire({
                title: "¿Que tipo de reporte desea?",
                showDenyButton: true,
                showCancelButton: false,
                confirmButtonText: "General",
                denyButtonText: "Solo Jefes"
                }).then((result) => {
                /* Read more about isConfirmed, isDenied below */
                if (result.isConfirmed) {
                    this.exportarConsolidadoExcel("general");
                } else if (result.isDenied) {
                    this.exportarConsolidadoExcel("jefes");
                }
            });    
        },
        async exportarConsolidadoExcel(tipoBuscar){
            this.loading = true;
            await caracterizacionService.exportarConsolidadoExcel(this.idBuscar, tipoBuscar).then(respuesta => {
                this.loading = false;
                Swal.fire({
                    icon: "success",
                    title: "Reporte Generado Correctamente",
                    text: "",
                    showConfirmButton: false,
                    footer: '<a id="descargarBtn" download class="btn btn-success" href="'+respuesta.data.nombre+'">Descargar archivo <i class="fas fa-file-excel"></i></a>'
                });


                document.getElementById('descargarBtn').addEventListener('click', () => {
                    Swal.close();
                });
            });
        },
        async consolidadoFunctionConcejo(){
            this.loading = true;
            await caracterizacionService.consolidadoConcejo(this.concejo, this.pagina).then(respuesta => {
                this.consolidado = respuesta.data;
                this.loading = false;
            });
            this.paginadorConcejo();
        },
        async paginadorConcejo() {
            await caracterizacionService.paginacionConsolidadoConcejo(this.concejo, this.pagina).then(respuesta => {
                this.paginacion = respuesta.data;
                this.pagina = parseInt(this.pagina);
                this.primera = this.pagina - 4 > 1 ? this.pagina - 4 : 1;
                this.ultima = this.pagina + 4 < this.paginacion.paginas_consulta ? this.pagina + 4 : this.paginacion.paginas_consulta;
            });
        },
        incrementarPaginaConcejo() {
            if (this.pagina < this.paginacion.paginas_consulta) {
                this.pagina++;
            }
            this.consolidadoFunctionConcejo();            
        },
        decrementarPaginaConcejo() {
            if (this.pagina > 1) {
                this.pagina--;
            }
            this.consolidadoFunctionConcejo();
        },
        irAPaginaConcejo(pagina) {
            this.pagina = pagina;
            this.consolidadoFunctionConcejo();
        },
        async exportarConsolidadoExcelConcejo(){
            this.loading = true;
            await caracterizacionService.exportarConsolidadoExcelConcejo(this.concejo).then(respuesta => {
                this.loading = false;
                Swal.fire({
                    icon: "success",
                    title: "Reporte Generado Correctamente",
                    text: "",
                    showConfirmButton: false,
                    footer: '<a id="descargarBtn" download class="btn btn-success" href="'+respuesta.data.nombre+'">Descargar archivo <i class="fas fa-file-excel"></i></a>'
                });


                document.getElementById('descargarBtn').addEventListener('click', () => {
                    Swal.close();
                });
            });
        }
    }
}
</script>
<style scoped>
    table {
        border-collapse: collapse;
    }

    thead {
        background-color: transparent !important;
        font-weight: bold;
    }

    td, th {
        border: 1px solid grey;
        padding: 10px;
    }

    #table2 td, #table2 th {
        border: 3px solid rgb(255, 255, 255) !important;
        padding: 10px;
    }

    .item_consolidado {
        background-color: #009199 ;
        border-radius: 15px;
        padding: 15px;
        color: #ffff;
        font-weight: bold;
        font-size: 18px;
        cursor: pointer;
    }

    .item_consolidado:hover {
        background-color: #027075 ;
    }

    .item_consolidado p {
        margin-bottom: 3px;
    }

    hr {
        margin-top: 1rem;
        margin-bottom: 1rem;
        border: 0;
        border-top: 1px solid rgb(69 65 65 / 59%);
    }

    .card_hover{
        transition: transform .3s linear;
    }

    .card_hover:hover {
        cursor: pointer;
        transform: scale(1.02);
    }

    .vertical-text {
      writing-mode: vertical-rl; 
      text-orientation: sideways;
      white-space: nowrap; 
      transform: rotate(180deg);
    }

    .pagination li {
        margin: 10px !important;
    }

    .pagination li.active button {
        width: 50px !important;
        height: 50px !important;
        border-radius: 50% !important;
    }

    .pagination li  button {
        width: 40px ;
        height: 40px !important;
        border-radius: 50% !important;
        font-size: 14px;
        font-weight: bold;
    }

    .pagination_style {
        width: auto !important;
        justify-content: center !important;
        align-items: center !important;
        background-color: #f3f3f3 !important;
        border-radius: 40px !important;
        border: 1px solid #eaeaea !important;
    }

    .bsiguiente {
        background-color: transparent !important;
        border: none !important;
        width: 100px !important;   
        display: flex;   
        align-items: center; 
    }

    .bsiguiente i{
        padding-top: 4px;
        margin-left: 10px;
    }

    .banterior {
        background-color: transparent !important;
        border: none !important;
        width: 100px !important;   
        display: flex;   
        align-items: center; 
    }

    .banterior i{
        padding-top: 4px;
        margin-right: 10px;
    }

    .boton_atras {
        position: absolute;
        right: 10px;
        border-radius: 50%;
        top: 20px;
        width: 70px;
        height: 70px;
        border: 8px solid white;
        font-size: 23px;
    }

    .boton_exportar {
        position: absolute;
        right: 80px;
        border-radius: 50%;
        top: 20px;
        width: 70px;
        height: 70px;
        border: 8px solid white;
        font-size: 23px;
    }

    .boton_atras:hover {
        border-color: #ffffff !important;
        background-color: #ff0022 !important;
        color: #FFF !important;
    }

    .boton_exportar:hover {
        border-color: #ffffff !important;
        background-color: #11a578 !important;
        color: #FFF !important;
    }
</style>