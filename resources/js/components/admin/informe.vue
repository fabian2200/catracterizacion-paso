<template lang="">
    <div style="background-color: white">

        <loading :active="loading" 
            :can-cancel="true"
            loader="bars" 
            color="#38b4c5"
            :height=100
            :width=200
            :on-cancel="onCancel"
            :is-full-page="true">
        </loading>

        <div class="card">
                <div class="card-header">
                    <h4 class="card-title" style="font-weight: bold">INFORME SOCIODEMOGRAFICO - POBLACIÓN AFRO</h4>
                </div>
                <div class="card-content">
                    <div class="card-body">
                        <ul class="nav nav-tabs nav-underline" role="tablist">
                            <li style ="font-size: 14px" class="nav-item">
                                <a class="nav-link active" id="baseIcon-tab20" data-toggle="tab" aria-controls="tabIcon20" href="#tabIcon20" role="tab" aria-selected="false">
                                    <i style ="font-size: 18px" class="fas fa-child"></i>
                                    <p style ="font-size: 18px">Afrocolombianos y Negritudes</p>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item">
                                <a class="nav-link" id="baseIcon-tab21" data-toggle="tab" aria-controls="tabIcon21" href="#tabIcon21" role="tab" aria-selected="false">
                                    <i style ="font-size: 18px" class="fas fa-user"></i>
                                    <p style ="font-size: 18px">Sociodemografico</p>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item">
                                <a class="nav-link" id="baseIcon-tab22" data-toggle="tab" aria-controls="tabIcon22" href="#tabIcon22" role="tab" aria-selected="false">
                                    <i style ="font-size: 18px" class="fas fa-suitcase"></i>
                                    <p style ="font-size: 18px">Empleo y Educación</p>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item">
                                <a class="nav-link" id="baseIcon-tab23" data-toggle="tab" aria-controls="tabIcon23" href="#tabIcon23" role="tab" aria-selected="false">
                                    <i style ="font-size: 18px" class="fas fa-school"></i>
                                    <p style ="font-size: 18px">Vivienda y Hogar</p>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item">
                                <a class="nav-link" id="baseIcon-tab24" data-toggle="tab" aria-controls="tabIcon24" href="#tabIcon24" role="tab" aria-selected="false">
                                    <i style ="font-size: 18px" class="fas fa-ambulance"></i>
                                    <p style ="font-size: 18px">Salud</p>
                                </a>
                            </li>
                        </ul>
                        <div style="position: relative" class="tab-content px-1 pt-1">
                            <div class="tab-pane active" id="tabIcon20" role="tabpanel" aria-labelledby="baseIcon-tab20">
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 100px; top: 20px" class="btn btn-danger" @click="generateReportAfro">
                                    <i class="fas fa-2x fa-file-pdf"></i>
                                </button>
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 20px; top: 20px" class="btn btn-success" @click="seleccionarOpcionesExportar(1)" data-toggle="modal" data-target="#modalPagina1">
                                    <i class="fas fa-2x fa-file-excel"></i>
                                </button>
                                <br>
                                <br>
                                <br>
                                <br>
                                <div class="row">
                                    <div class="col-lg-2"></div>
                                    <div class="col-lg-8">
                                        <div id="pestana_afrocolombianos">
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">DISTRIBUCIÓN POBLACIÓN POR CONCEJO</h3>
                                            <br>
                                            <h3 style="line-height: 1.6;" v-if="afrocolombianos_negritudes != null">
                                                Se tiene que de un total de <strong>{{afrocolombianos_negritudes.cantidad_poblacion}}</strong> personas caracterizadas, se tiene que:
                                            </h3>
                                            <h5 v-if="afrocolombianos_negritudes != null" v-for="(item, index) in afrocolombianos_negritudes.por_concejo" :key="index">
                                                <strong>{{ item.cantidad }} personas</strong> pertenecen al concejo comunitario <strong>{{ item.concejo }}</strong>.  
                                            </h5>
                                            <br><br>
                                            <div>
                                                <h3 style="width: 100%; text-align: center; font-weight: bold;">PORCENTAJE DE POBLACIÓN POR CONCEJO COMUNITARIO</h3>
                                                <div id="grafica_concejo" style="height: 300px"></div>
                                            </div>
                                            <br><br>
                                            <div>
                                                <h3 style="width: 100%; text-align: center; font-weight: bold;">PORCENTAJE DE POBLACIÓN POR ETNIA</h3>
                                                <div id="grafica_etnia" style="height: 300px"></div>
                                            </div>
                                            <br>
                                            <div>
                                                <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">PRÁCTICAS CULTURALES Y RELIGIOSAS</h3>
                                                <br>
                                                <table style="width: 100%" class="tabla_informe" v-if="afrocolombianos_negritudes != null">
                                                    <thead>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">Actividad Cultural y religiosa</th>
                                                            <th style="background-color: #5df3c5;"># Personas</th>
                                                            <th style="background-color: #5df3c5;">% Total</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr v-for="(item, index) in afrocolombianos_negritudes.por_practicas_religiosas" :key="index">
                                                            <td>{{ item.practicas_religiosas }}</td>
                                                            <td>{{ item.cantidad }}</td>
                                                            <td>{{ ((item.cantidad / afrocolombianos_negritudes.cantidad_poblacion) * 100).toFixed(2) }} %</td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                            <br>
                                            <div id="grafica_practicas" style="height: 300px"></div>
                                            <br>
                                            <div>
                                                <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">PERSONAS QUE HABLAN Y NO HABLAN ALGUNA LENGUA AFROCOLOMBIANA</h3>
                                                <br>
                                                <table style="width: 100%" class="tabla_informe" v-if="afrocolombianos_negritudes != null">
                                                    <thead>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">¿Habla alguna lengua afrocolombiana?</th>
                                                            <th style="background-color: #5df3c5;">Respuesta</th>
                                                            <th style="background-color: #5df3c5;">% Total</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr v-for="(item, index) in afrocolombianos_negritudes.habla_lenguas" :key="index">
                                                            <td>{{ item.habla_lengua }}</td>
                                                            <td>{{ item.cantidad }}</td>
                                                            <td>{{ ((item.cantidad / afrocolombianos_negritudes.cantidad_poblacion) * 100).toFixed(2) }} %</td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                            <br>
                                            <div>
                                                <h3 style="width: 100%; text-align: center; font-weight: bold; color: #ff425c; ">PORCENTAJE DE POBLACIÓN</h3>
                                                <h5 style="width: 100%; text-align: center; font-weight: bold;">(Habla lengua afrocolombiana)</h5>
                                                <div id="grafica_lengua" style="height: 300px"></div>
                                            </div>
                                        </div>
                                    </div>
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon21" role="tabpanel" aria-labelledby="baseIcon-tab21">
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 100px; top: 20px" class="btn btn-danger" @click="generateReport">
                                    <i class="fas fa-2x fa-file-pdf"></i>
                                </button>
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 20px; top: 20px" class="btn btn-success" @click="seleccionarOpcionesExportar(2)" data-toggle="modal" data-target="#modalPagina1">
                                    <i class="fas fa-2x fa-file-excel"></i>
                                </button>
                                <br>
                                <br>
                                <br>
                                <br>
                                <div class="row">
                                    <div class="col-lg-2"></div>
                                    <div class="col-lg-8">
                                        <div id="pestana_sociodemografico">
                                            <div>
                                                <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">DISTRIBUCIÓN POR GRUPO DE EDAD</h3>
                                                <br>
                                                <table class="tabla_informe" style="width: 100%" v-if="datos_edad.length > 0">
                                                    <thead>
                                                        <tr>
                                                            <th colspan="7">
                                                                Distribución por Grupos de Edad
                                                            </th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr style="background-color: #5df3c5">
                                                            <th>EDAD</th>
                                                            <th>Hombres</th>
                                                            <th>%</th>
                                                            <th>Mujeres</th>
                                                            <th>%</th>
                                                            <th>Total</th>
                                                            <th>%Total</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 0 A 4</th>
                                                            <th>{{datos_edad[0][0]}}</th>
                                                            <th>{{(datos_edad[0][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[0][1]}}</th>
                                                            <th>{{(datos_edad[0][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[0][0]+datos_edad[0][1]}}</th>
                                                            <th>{{((datos_edad[0][0]+datos_edad[0][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 5 A 9</th>
                                                            <th>{{datos_edad[1][0]}}</th>
                                                            <th>{{(datos_edad[1][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[1][1]}}</th>
                                                            <th>{{(datos_edad[1][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[1][0]+datos_edad[1][1]}}</th>
                                                            <th>{{((datos_edad[1][0]+datos_edad[1][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 10 A 14</th>
                                                            <th>{{datos_edad[2][0]}}</th>
                                                            <th>{{(datos_edad[2][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[2][1]}}</th>
                                                            <th>{{(datos_edad[2][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[2][0]+datos_edad[2][1]}}</th>
                                                            <th>{{((datos_edad[2][0]+datos_edad[2][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 15 A 19</th>
                                                            <th>{{datos_edad[3][0]}}</th>
                                                            <th>{{(datos_edad[3][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[3][1]}}</th>
                                                            <th>{{(datos_edad[3][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[3][0]+datos_edad[3][1]}}</th>
                                                            <th>{{((datos_edad[3][0]+datos_edad[3][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 20 A 24</th>
                                                            <th>{{datos_edad[4][0]}}</th>
                                                            <th>{{(datos_edad[4][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[4][1]}}</th>
                                                            <th>{{(datos_edad[4][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[4][0]+datos_edad[4][1]}}</th>
                                                            <th>{{((datos_edad[4][0]+datos_edad[4][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 25 A 29</th>
                                                            <th>{{datos_edad[5][0]}}</th>
                                                            <th>{{(datos_edad[5][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[5][1]}}</th>
                                                            <th>{{(datos_edad[5][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[5][0]+datos_edad[5][1]}}</th>
                                                            <th>{{((datos_edad[5][0]+datos_edad[5][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 30 A 34</th>
                                                            <th>{{datos_edad[6][0]}}</th>
                                                            <th>{{(datos_edad[6][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[6][1]}}</th>
                                                            <th>{{(datos_edad[6][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[6][0]+datos_edad[1][1]}}</th>
                                                            <th>{{((datos_edad[6][0]+datos_edad[6][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 35 A 39</th>
                                                            <th>{{datos_edad[7][0]}}</th>
                                                            <th>{{(datos_edad[7][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[7][1]}}</th>
                                                            <th>{{(datos_edad[7][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[7][0]+datos_edad[7][1]}}</th>
                                                            <th>{{((datos_edad[7][0]+datos_edad[7][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 40 A 44</th>
                                                            <th>{{datos_edad[8][0]}}</th>
                                                            <th>{{(datos_edad[8][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[8][1]}}</th>
                                                            <th>{{(datos_edad[8][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[8][0]+datos_edad[8][1]}}</th>
                                                            <th>{{((datos_edad[8][0]+datos_edad[8][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 45 A 49</th>
                                                            <th>{{datos_edad[9][0]}}</th>
                                                            <th>{{(datos_edad[9][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[9][1]}}</th>
                                                            <th>{{(datos_edad[9][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[9][0]+datos_edad[9][1]}}</th>
                                                            <th>{{((datos_edad[9][0]+datos_edad[9][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 50 A 54</th>
                                                            <th>{{datos_edad[10][0]}}</th>
                                                            <th>{{(datos_edad[10][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[10][1]}}</th>
                                                            <th>{{(datos_edad[10][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[10][0]+datos_edad[10][1]}}</th>
                                                            <th>{{((datos_edad[10][0]+datos_edad[10][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 55 A 59</th>
                                                            <th>{{datos_edad[11][0]}}</th>
                                                            <th>{{(datos_edad[11][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[11][1]}}</th>
                                                            <th>{{(datos_edad[11][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[11][0]+datos_edad[11][1]}}</th>
                                                            <th>{{((datos_edad[11][0]+datos_edad[11][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 60 A 64</th>
                                                            <th>{{datos_edad[12][0]}}</th>
                                                            <th>{{(datos_edad[12][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[12][1]}}</th>
                                                            <th>{{(datos_edad[12][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[12][0]+datos_edad[12][1]}}</th>
                                                            <th>{{((datos_edad[12][0]+datos_edad[12][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 65 A 69</th>
                                                            <th>{{datos_edad[13][0]}}</th>
                                                            <th>{{(datos_edad[13][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[13][1]}}</th>
                                                            <th>{{(datos_edad[13][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[13][0]+datos_edad[13][1]}}</th>
                                                            <th>{{((datos_edad[13][0]+datos_edad[13][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 70 A 74</th>
                                                            <th>{{datos_edad[14][0]}}</th>
                                                            <th>{{(datos_edad[14][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[14][1]}}</th>
                                                            <th>{{(datos_edad[14][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[14][0]+datos_edad[14][1]}}</th>
                                                            <th>{{((datos_edad[14][0]+datos_edad[14][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>De 75 Y MAS</th>
                                                            <th>{{datos_edad[15][0]}}</th>
                                                            <th>{{(datos_edad[15][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{datos_edad[15][1]}}</th>
                                                            <th>{{(datos_edad[15][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad[15][0]+datos_edad[15][1]}}</th>
                                                            <th>{{((datos_edad[15][0]+datos_edad[15][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th style="background-color: #5df3c5">TOTAL</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad_general.numero_masculino}}</th>
                                                            <th style="background-color: #5df3c5">{{(datos_edad_general.numero_masculino / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{datos_edad_general.numero_femenino}}</th>
                                                            <th style="background-color: #5df3c5">{{(datos_edad_general.numero_femenino / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th style="background-color: #5df3c5">{{ numero_personas }}</th>
                                                            <th>100 %</th>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <br>
                                                <table style="width: 100%" class="tabla_informe" v-if="datos_edad.length > 0">
                                                    <thead>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;" colspan="3">Menores de 15 años</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">H</th>
                                                            <th style="background-color: #5df3c5;">M</th>
                                                            <th style="background-color: #5df3c5; width: 33%; font-size: 20px" rowspan="5"> &lt; 15 </th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{(datos_edad[16][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{(datos_edad[16][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{datos_edad[16][0]}}</th>
                                                            <th>{{datos_edad[16][1]}}</th>
                                                        </tr>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">Total</th>
                                                            <th style="background-color: #5df3c5;">Total %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{datos_edad[16][0] + datos_edad[16][1]}}</th>
                                                            <th>{{((datos_edad[16][0] + datos_edad[16][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                            <br>
                                            <div>
                                                <table style="width: 100%" class="tabla_informe" v-if="datos_edad.length > 0">
                                                    <thead>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;" colspan="3">De 15 a 64 años</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">H</th>
                                                            <th style="background-color: #5df3c5;">M</th>
                                                            <th style="background-color: #5df3c5; width: 33%; font-size: 20px" rowspan="5"> &gt;= 15 &lt;= 64 </th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{(datos_edad[17][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{(datos_edad[17][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{datos_edad[17][0]}}</th>
                                                            <th>{{datos_edad[17][1]}}</th>
                                                        </tr>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">Total</th>
                                                            <th style="background-color: #5df3c5;">Total %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{datos_edad[17][0] + datos_edad[17][1]}}</th>
                                                            <th>{{((datos_edad[17][0] + datos_edad[17][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                                <br>
                                                <table style="width: 100%" class="tabla_informe" v-if="datos_edad.length > 0">
                                                    <thead>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;" colspan="3">De 65 años en adelante</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">H</th>
                                                            <th style="background-color: #5df3c5;">M</th>
                                                            <th style="background-color: #5df3c5; width: 33%;  font-size: 20px" rowspan="5"> &gt;= 65  </th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{(datos_edad[18][0] / numero_personas * 100).toFixed(2)}} %</th>
                                                            <th>{{(datos_edad[18][1] / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{datos_edad[18][0]}}</th>
                                                            <th>{{datos_edad[18][1]}}</th>
                                                        </tr>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">Total</th>
                                                            <th style="background-color: #5df3c5;">Total %</th>
                                                        </tr>
                                                        <tr>
                                                            <th>{{datos_edad[18][0] + datos_edad[18][1]}}</th>
                                                            <th>{{((datos_edad[18][0] + datos_edad[18][1]) / numero_personas * 100).toFixed(2)}} %</th>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                            <br><br>
                                            <div>
                                                <h3 style="width: 100%; text-align: center; font-weight: bold;">PORCENTAJE DE LA POBLACIÓN POR GENERO</h3>
                                                <div id="grafica_sexo_informe" style="height: 300px"></div>
                                            </div>
                                            <br>
                                            <div>
                                                <h3 style="width: 100%; text-align: center; font-weight: bold;">PIRÁMIDE POBLACIONAL</h3>
                                                <br>
                                                <div id="grafico_edad_informe" style="height: 465px; width: 104%"></div>
                                            </div>
                                            <br>
                                            <br>
                                            <div>
                                                <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">POBLACIÓN DESPLAZADA</h3>
                                                <br>
                                                <table style="width: 100%" class="tabla_informe" v-if="datos_edad.length > 0">
                                                    <thead>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">Grupo Edad</th>
                                                            <th style="background-color: #5df3c5;">Masculinos</th>
                                                            <th style="background-color: #5df3c5;">Femeninos</th>
                                                            <th style="background-color: #5df3c5;">Total</th>
                                                            <th style="background-color: #5df3c5;">% Total</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr>
                                                            <td><strong>{{ datos_desplazados.desplazados_m15[2] }}</strong></td>
                                                            <td>{{ datos_desplazados.desplazados_m15[0] }}</td>
                                                            <td>{{ datos_desplazados.desplazados_m15[1] }}</td>
                                                            <td>{{ datos_desplazados.desplazados_m15[0] + datos_desplazados.desplazados_m15[1]}}</td>
                                                            <td>{{ (((datos_desplazados.desplazados_m15[0] + datos_desplazados.desplazados_m15[1]) / datos_desplazados.cantidad_poblacion_desplazada) * 100).toFixed(2) }} %</td>
                                                        </tr>
                                                        <tr>
                                                            <td><strong>{{ datos_desplazados.desplazados_15a64[2] }}</strong></td>
                                                            <td>{{ datos_desplazados.desplazados_15a64[0] }}</td>
                                                            <td>{{ datos_desplazados.desplazados_15a64[1] }}</td>
                                                            <td>{{ datos_desplazados.desplazados_15a64[0] + datos_desplazados.desplazados_15a64[1]}}</td>
                                                            <td>{{ (((datos_desplazados.desplazados_15a64[0] + datos_desplazados.desplazados_15a64[1]) / datos_desplazados.cantidad_poblacion_desplazada) * 100).toFixed(2) }} %</td>
                                                        </tr>
                                                        <tr>
                                                            <td><strong>{{ datos_desplazados.desplazados_mas65[2] }}</strong></td>
                                                            <td>{{ datos_desplazados.desplazados_mas65[0] }}</td>
                                                            <td>{{ datos_desplazados.desplazados_mas65[1] }}</td>
                                                            <td>{{ datos_desplazados.desplazados_mas65[0] + datos_desplazados.desplazados_mas65[1]}}</td>
                                                            <td>{{ (((datos_desplazados.desplazados_mas65[0] + datos_desplazados.desplazados_mas65[1]) / datos_desplazados.cantidad_poblacion_desplazada) * 100).toFixed(2) }} %</td>
                                                        </tr>
                                                        <tr>
                                                            <th style="background-color: #5df3c5;">Total</th>
                                                            <th style="background-color: #5df3c5;">{{ datos_desplazados.masculinos_desplazados }}</th>
                                                            <th style="background-color: #5df3c5;">{{ datos_desplazados.femeninos_desplazados }}</th>
                                                            <th style="background-color: #5df3c5;">{{ datos_desplazados.cantidad_poblacion_desplazada }}</th>
                                                            <th style="background-color: #5df3c5;">100%</th>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                            <br>
                                            <div class="page-break"></div>
                                            <div>
                                                <br>
                                                <h3 style="width: 100%; text-align: center; font-weight: bold;">DISTRIBUCIÓN POBLACIÓN DESPLAZADAS</h3>
                                                <h5 style="width: 100%; text-align: center; font-weight: bold;">(por grupo de edad)</h5>
                                                <br>
                                                <div id="grafico_desplazados" style="margin-top: 10px; height: 350px; width: 104%"></div>
                                            </div>
                                            <div style="margin-top: 15px">
                                                <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">DISTRIBUCIÓN POR ESTADO CIVIL</h3>
                                                <br>
                                                <table class="tabla_informe" style="width: 100%" v-if="datos_edad.length > 0">
                                                    <thead>
                                                        <tr>
                                                            <th colspan="3">
                                                                ESTADO CIVIL
                                                            </th>
                                                        </tr>
                                                        <tr>
                                                            <th>ESTADO CIVIL</th>
                                                            <th>PERSONAS</th>
                                                            <th>%</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr v-for="(item, index) in datos_edad_general.estado_civil" :key="index">
                                                            <th style="text-transform: uppercase">{{item[0]}}</th>
                                                            <td>{{item[1]}}</td>
                                                            <td>{{item[2]}} %</td>
                                                        </tr>
                                                    </tbody>
                                                    <thead>
                                                        <tr>
                                                            <th>TOTAL</th>
                                                            <th>{{numero_personas}}</th>
                                                            <th>100%</th>
                                                        </tr>
                                                    </thead>
                                                </table>
                                            </div>
                                            <br>
                                            <br>
                                            <div>
                                                <h3 style="width: 100%; text-align: center; font-weight: bold;">DISTRIBUCIÓN DE PERSONAS POR ESTADO CIVIL</h3>
                                                <div id="grafico_estado_civil_por" style="margin-top: 30px; height: 390px; width: 104%"></div>
                                            </div>
                                        </div>
                                    </div> 
                                    <div class="col-lg-2"></div>     
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon22" role="tabpanel" aria-labelledby="baseIcon-tab22">
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 100px; top: 20px" class="btn btn-danger" @click="generateReport2">
                                    <i class="fas fa-2x fa-file-pdf"></i>
                                </button>
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 20px; top: 20px" class="btn btn-success" @click="seleccionarOpcionesExportar(3)" data-toggle="modal" data-target="#modalPagina1">
                                    <i class="fas fa-2x fa-file-excel"></i>
                                </button>
                                <br>
                                <br>
                                <br>
                                <br>
                                <div class="row">
                                    <div class="col-lg-2"></div>  
                                    <div class="col-lg-8">
                                        <div id="pestana_empleo_educacion">
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">POBLACION ECONOMICAMENTE ACTIVA</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="poblacion_e_activa != null">
                                                <thead>
                                                    <tr>
                                                        <th colspan="6">POBLACION ECONOMICAMENTE ACTIVA</th>
                                                    </tr>
                                                    <tr>
                                                        <th>HOMBRES</th>
                                                        <th>%</th>
                                                        <th>MUJERES</th>
                                                        <th>%</th>
                                                        <th>TOTAL</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr style="border-left: 1px solid grey; border-right: 1px solid grey">
                                                        <td>{{poblacion_e_activa[0]}}</td>
                                                        <td>{{(poblacion_e_activa[0]/poblacion_e_activa[2] * 100).toFixed(2)}}%</td>
                                                        <td>{{poblacion_e_activa[1]}}</td>
                                                        <td>{{(poblacion_e_activa[1]/poblacion_e_activa[2] * 100).toFixed(2)}}%</td>
                                                        <td>{{poblacion_e_activa[0] + poblacion_e_activa[1]}}</td>
                                                        <td>100%</td>
                                                    </tr>
                                                    <tr style="border-left: 1px solid grey; border-right: 1px solid grey">
                                                        <td colspan="6" style="text-align: center; padding: 10px">
                                                            También se le llama fuerza laboral y está conformada por las personas en edad de trabajar que trabajan o están buscando empleo. 
                                                        </td>
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th colspan="6">POBLACION ECONOMICAMENTE INACTIVA</th>
                                                    </tr>
                                                    <tr>
                                                        <th>HOMBRES</th>
                                                        <th>%</th>
                                                        <th>MUJERES</th>
                                                        <th>%</th>
                                                        <th>TOTAL</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr style="border-left: 1px solid grey; border-right: 1px solid grey">
                                                        <td>{{poblacion_e_inactiva[0]}}</td>
                                                        <td>{{(poblacion_e_inactiva[0]/poblacion_e_inactiva[2] * 100).toFixed(2)}}%</td>
                                                        <td>{{poblacion_e_inactiva[1]}}</td>
                                                        <td>{{(poblacion_e_inactiva[1]/poblacion_e_inactiva[2] * 100).toFixed(2)}}%</td>
                                                        <td>{{poblacion_e_inactiva[0] + poblacion_e_inactiva[1]}}</td>
                                                        <td>100%</td>
                                                    </tr>
                                                    <tr style="border-left: 1px solid grey; border-right: 1px solid grey; border-bottom: 1px solid grey">
                                                        <td colspan="6" style="text-align: center; padding: 10px">
                                                            Comprende a todas las personas en edad de trabajar que en la semana de referencia no participan en la producción de bienes y servicios porque no necesitan, no pueden o no están interesadas en tener actividad remunerada. A este grupo pertenecen estudiantes, amas de casa, pensionados, jubilados, rent istas, inválidos (incapacitados Permanentemente para trabajar), personas que no les llama la atención o creen que no vale la pena trabajar.
                                                        </td>
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <br><br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">(P.E.A) PORCENTAJE POR GENERO</h3>
                                            <div id="grafica_pae" style="height: 300px;"></div>
                                            <br><br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">(P.E.I) PORCENTAJE POR GENERO</h3>
                                            <div id="grafica_pei" style="height: 300px"></div>
                                            <br><br>
                                            <div class="page-break"></div>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">DISTRIBUCIÓN DE PERSONAS POR OCUPACIÓN</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="datos_edad.length > 0">
                                                <thead>
                                                    <tr>
                                                        <th colspan="3">OCUPACIÓN PRINCIPAL</th>
                                                    </tr>
                                                    <tr>
                                                        <th>OCUPACIÓN PRINCIPAL</th>
                                                        <th># PERSONAS</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in datos_por_ocupacion" :key="index">
                                                        <td>{{item.ocupacion}}</td>
                                                        <td>{{item.cantidad}}</td>
                                                        <td>{{(item.cantidad / numero_personas * 100).toFixed(2)}} %</td>                                    
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th>Total</th>
                                                        <th>{{numero_personas}}</th>
                                                        <th>100 %</th>
                                                    </tr>
                                                </thead>
                                            </table>
                                            <div class="page-break"></div>
                                            <br>
                                            <br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">5 PRINCIPALES OCUPACIONES</h3>
                                            <div id="grafica_principales_ocupaciones" style="height: 350px"></div>
                                            <br><br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">SITUACIÓN LABORAL</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="datos_edad.length > 0">
                                                <thead>
                                                    <tr>
                                                        <th>SITUACIÓN LABORAL</th>
                                                        <th>HOMBRES</th>
                                                        <th>MUJERES</th>
                                                        <th># PERSONAS</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr>
                                                        <th>Trabajo Formal</th>
                                                        <td>{{situacion_laboral.trabajo_formal_m}}</td>
                                                        <td>{{situacion_laboral.trabajo_formal_f}}</td>
                                                        <td>{{situacion_laboral.trabajo_formal}}</td>
                                                    </tr>
                                                    <tr>
                                                        <th>Trabajo Informal</th>
                                                        <td>{{situacion_laboral.trabajo_informal_m}}</td>
                                                        <td>{{situacion_laboral.trabajo_informal_f}}</td>
                                                        <td>{{situacion_laboral.trabajo_informal}}</td>
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th>TOTAL</th>
                                                        <th>{{situacion_laboral.trabajo_formal_m + situacion_laboral.trabajo_informal_m}}</th>
                                                        <th>{{situacion_laboral.trabajo_formal_f + situacion_laboral.trabajo_informal_f}}</th>
                                                        <th>{{situacion_laboral.trabajo_formal + situacion_laboral.trabajo_informal}}</th>
                                                    </tr>
                                                </thead>
                                            </table>
                                            <br>
                                            <br>
                                            <br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">TRABAJO INFORMAL vs TRABAJO FORMAL</h3>
                                            <div id="grafica_situacion_laboral" style="height: 350px"></div>
                                            <br><br>
                                            <br>
                                            <div class="page-break"></div>
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">DISTRIBUCIÓN DE PERSONAS POR NIVEL ACADÉMICO</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="datos_edad.length > 0">
                                                <thead>
                                                    <tr>
                                                        <th colspan="3">NIVEL ACADÉMICO</th>
                                                    </tr>
                                                    <tr>
                                                        <th>ESCOLARIDAD</th>
                                                        <th>PERSONAS</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in datos_nivel_academico" :key="index">
                                                        <td>{{item.descripcion}}</td>
                                                        <td>{{item.cantidad}}</td>
                                                        <td>{{(item.cantidad / numero_personas * 100).toFixed(2)}} %</td>
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th>TOTAL</th>
                                                        <th>{{numero_personas}}</th>
                                                        <th>100%</th>
                                                    </tr>
                                                </thead>
                                            </table>
                                            <br><br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">PORCENTAJE DE PERSONAS POR NIVEL ACADÉMICO</h3>
                                            <div id="grafica_nivel_academico" style="height: 300px"></div>
                                            <br><br>
                                            <div class="page-break"></div>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">DISTRIBUCIÓN DEL NIVEL ACADÉMICO DE LA POBLACIÓN POR SEXO</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="datos_edad.length > 0">
                                                <thead>
                                                    <tr>
                                                        <th>ESCOLARIDAD</th>
                                                        <th>HOMBRES</th>
                                                        <th>%</th>
                                                        <th>MUJERES</th>
                                                        <th>%</th>
                                                        <th>TOTAL</th>
                                                        <th>% TOTAL</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in datos_nivel_academico_sexo" :key="index">
                                                        <td>{{item.nivel_educativo}}</td>
                                                        <td>{{item.hombres}}</td>
                                                        <td>{{(item.hombres / personas_nivel_educativo.total * 100).toFixed(2)}} %</td>
                                                        <td>{{item.mujeres}}</td>
                                                        <td>{{(item.mujeres / personas_nivel_educativo.total * 100).toFixed(2)}} %</td>
                                                        <td>{{item.hombres + item.mujeres}}</td>
                                                        <td>{{((item.mujeres +item.hombres)/ personas_nivel_educativo.total * 100).toFixed(2)}} %</td>
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th>TOTAL</th>
                                                        <th>{{personas_nivel_educativo.total_hombres}}</th>
                                                        <th>{{(personas_nivel_educativo.total_hombres / personas_nivel_educativo.total * 100).toFixed(2)}} %</th>
                                                        <th>{{personas_nivel_educativo.total_mujeres}}</th>
                                                        <th>{{(personas_nivel_educativo.total_mujeres / personas_nivel_educativo.total * 100).toFixed(2)}} %</th>
                                                        <th>{{personas_nivel_educativo.total_hombres + personas_nivel_educativo.total_mujeres}}</th>
                                                        <th>100 %</th>
                                                    </tr>
                                                </thead>
                                                <thead>
                                                    <tr>
                                                        <th colspan="2">Frente al total de la población</th>
                                                        <th>{{(personas_nivel_educativo.total_hombres / numero_personas * 100).toFixed(2)}} %</th>
                                                        <td></td>
                                                        <th>{{(personas_nivel_educativo.total_mujeres / numero_personas * 100).toFixed(2)}} %</th>
                                                        <td></td>
                                                        <th>{{((personas_nivel_educativo.total_mujeres + personas_nivel_educativo.total_hombres) / numero_personas * 100).toFixed(2)}} %</th>
                                                    </tr>
                                                </thead>
                                            </table>
                                            <br>
                                            <br>
                                            <br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">DISTRIBUCIÓN DE LA POBLACIÓN POR NIVEL ACADÉMICO Y SEXO</h3>
                                            <div id="grafica_nivel_academico_sexo" style="height: 600px"></div>
                                        </div>
                                    </div>
                                    <div class="col-lg-2"></div>  
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon23" role="tabpanel" aria-labelledby="baseIcon-tab23">
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 100px; top: 20px" class="btn btn-danger" @click="generateReport3">
                                    <i class="fas fa-2x fa-file-pdf"></i>
                                </button>
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 20px; top: 20px" class="btn btn-success" @click="seleccionarOpcionesExportar(4)" data-toggle="modal" data-target="#modalPagina1">
                                    <i class="fas fa-2x fa-file-excel"></i>
                                </button>
                                <br>
                                <br>
                                <br>
                                <br>
                                <div class="row">
                                    <div class="col-lg-2"></div>  
                                    <div class="col-lg-8">
                                        <div id="pestana_vivienda_hogar">
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">ACCESO A SERVICIOS BÁSICOS</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="posesion_vivienda != null">
                                                <thead>
                                                    <tr>
                                                        <th colspan="3">VIVIENDAS CON SERVICIOS BÁSICOS</th>
                                                    </tr>
                                                    <tr>
                                                        <th>SERVICIO</th>
                                                        <th>SI</th>
                                                        <th>NO</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in servicios" :key="index">
                                                        <td>{{ item[2] }}</td>
                                                        <td>{{ item[0] }}</td>
                                                        <td>{{ item[1] }}</td>  
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <br><br>
                                            <div style="width: 100%; text-align: center">
                                                <h4><strong>ACCESO A SERVICIOS PÚBLICOS</strong></h4>
                                                <div id="grafico_servicios" style="height: 350px"></div>
                                            </div>
                                            <br>
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">DISTRIBUCIÓN DE FAMILIAS POR TIPO DE VIVIENDA</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="posesion_vivienda != null">
                                                <thead>
                                                    <tr>
                                                        <th colspan="3">TIPOS DE VIVIENDA</th>
                                                    </tr>
                                                    <tr>
                                                        <th>TIPO</th>
                                                        <th>#FAMILIAS</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in tipo_vivienda" :key="index">
                                                        <td>{{ item.tipo_vivienda }}</td>
                                                        <td>{{ item.cantidad }}</td>
                                                        <td>{{ ((item.cantidad / cantidad_viviendas) * 100).toFixed(2) }} %</td>  
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <br><br>
                                            <div class="page-break"></div>
                                            <br>
                                            <div style="width: 100%; text-align: center">
                                                <h4><strong>DISTRIBUCION DE FAMILIAS</strong></h4>
                                                <h5><strong>por tipo de vivienda</strong></h5>
                                                <div id="grafica_tipo_vivienda" style="height: 350px"></div>
                                            </div>
                                            <br>
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">DISTRIBUCIÓN DE FAMILIAS POR POSESIÓN DE VIVIENDA PROPIA</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="posesion_vivienda != null">
                                                <thead>
                                                    <tr>
                                                        <th colspan="3">POSESIÓN DE VIVIENDA</th>
                                                    </tr>
                                                    <tr>
                                                        <th>VIVIENDA</th>
                                                        <th>FAMILIAS</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr>
                                                        <td>POSEEN</td>
                                                        <td>{{posesion_vivienda.poseen}}</td>
                                                        <td>{{(posesion_vivienda.poseen / (posesion_vivienda.no_poseen + posesion_vivienda.poseen) * 100).toFixed(2)}} %</td>                                    
                                                    </tr>
                                                    <tr>
                                                        <td>NO POSEEN</td>
                                                        <td>{{posesion_vivienda.no_poseen}}</td>
                                                        <td>{{(posesion_vivienda.no_poseen / (posesion_vivienda.no_poseen + posesion_vivienda.poseen) * 100).toFixed(2)}} %</td>                                    
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th>Total</th>
                                                        <th>{{posesion_vivienda.no_poseen + posesion_vivienda.poseen}}</th>
                                                        <th>100 %</th>
                                                    </tr>
                                                </thead>
                                            </table>
                                            <br><br>
                                            <h3 style="width: 100%; text-align: center">PORCENTAJE DE FAMILIA POR TIPO DE POSESIÓN</h3>
                                            <div id="grafica_posesion" style="height: 350px"></div>
                                            <br><br>
                                            <div class="page-break"></div>
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">TENENCIA DE TIERRAS</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="posesion_vivienda != null">
                                                <thead>
                                                    <tr>
                                                        <th colspan="5">TENENCIA DE TIERRAS</th>
                                                    </tr>
                                                    <tr>
                                                        <th>RESPUESTA</th>
                                                        <th>POSESIÓN BALDÍOS</th>
                                                        <th>PROPIEDAD CON TITULO</th>
                                                        <th>AREA TOTAL (m<sup>2</sup>)</th>
                                                        <th>AREA TOTAL (Ha)</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr>
                                                        <td>Si</td>
                                                        <td>{{posesion_baldios.baldios_si}}</td>
                                                        <td>{{posesion_baldios.titulo_si}}</td> 
                                                        <td rowspan="2">{{posesion_baldios.total_area}} <strong>m <sup>2</sup></strong>  </td>
                                                        <td rowspan="2">{{posesion_baldios.total_area / 10000}} <strong>Hta</strong></td>                                   
                                                    </tr>
                                                    <tr>
                                                        <td>No</td>
                                                        <td>{{posesion_baldios.baldios_no}}</td>
                                                        <td>{{posesion_baldios.titulo_no}}</td> 
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <br><br>
                                            <h3 style="width: 100%; text-align: center">PRINCIPALES CULTIVOS</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="posesion_vivienda != null">
                                                <thead>
                                                    <tr>
                                                        <th colspan="4">CULTIVOS</th>
                                                    </tr>
                                                    <tr>
                                                        <th>LINEA</th>
                                                        <th>ACTIVIDAD</th>
                                                        <th>AREA TOTAL (m<sup>2</sup>)</th>
                                                        <th>AREA TOTAL (Ha)</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in cultivos" :key="index">
                                                        <td>{{item.linea}}</td>
                                                        <td>{{item.actividad}}</td>
                                                        <td>{{item.area_destinada}}</td> 
                                                        <td>{{item.area_destinada / 10000}}</td> 
                                                    </tr>
                                                </tbody>
                                            </table>
                                            <div class="page-break"></div>
                                            <br>
                                            <h3 style="width: 100%; text-align: center">PORCENTAJE DE HECTÁREAS OCUPADAS</h3>
                                            <h5 style="width: 100%; text-align: center; font-weight: bold;">por linea economica</h5>
                                            <div id="grafica_linea_economica" style="height: 300px"></div>
                                            <br><br>
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">PORCENTAJE DE HECTÁREAS OCUPADAS</h3>
                                            <h5 style="width: 100%; text-align: center; font-weight: bold;">por linea economica y actividad</h5>
                                            <br>
                                            <div v-for="(item, index) in cultivosPorLinea" :key="index">
                                                <br><br>
                                                <h3 style="height: 45px; width: 100%; text-align: center">{{ item[0].linea }}</h3>
                                                <br>
                                                <div :id="'grafica_linea_'+index" style="height: 200px"></div>
                                                <br><br>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-lg-2"></div>  
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon24" role="tabpanel" aria-labelledby="baseIcon-tab24">
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 100px; top: 20px" class="btn btn-danger" @click="generateReportSalud">
                                    <i class="fas fa-2x fa-file-pdf"></i>
                                </button>
                                <button style="border-radius: 50%; width: 60px; height: 60px; position: absolute; right: 20px; top: 20px" class="btn btn-success" @click="seleccionarOpcionesExportar(5)" data-toggle="modal" data-target="#modalPagina1">
                                    <i class="fas fa-2x fa-file-excel"></i>
                                </button>
                                <br>
                                <br>
                                <br>
                                <br>
                                <div class="row">
                                    <div class="col-lg-2"></div>  
                                    <div class="col-lg-8">
                                        <div id="pestana_salud">
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">ESTADO DE SALUD GENERAL DE LA POBLACIÓN</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="salud != null">
                                                <thead>
                                                    <tr>
                                                        <th>Estado de salud</th>
                                                        <th>Numero de personas</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in salud.estado_salud" :key="index">
                                                        <td>{{ item.estado_salud }}</td>
                                                        <td>{{ item.cantidad }}</td>
                                                        <td>{{ ((item.cantidad / salud.cantidad_poblacion) * 100).toFixed(2) }}</td>  
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th>Total</th>
                                                        <th>{{ salud.cantidad_poblacion }}</th>
                                                        <th>100%</th>
                                                    </tr>
                                                </thead>
                                            </table>
                                            <br>
                                            <br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">PORCENTAJE DE PERSONAS POR ESTADO DE SALUD</h3>
                                            <div id="grafica_estado_salud" style="height: 300px"></div>
                                            <br>
                                            <h3 style="color: #ff425c; width: 100%; text-align: center; font-weight: bold;">PERSONAS POR TIPO DE ACCESO A LA SALUD</h3>
                                            <div id="grafica_acceso_salud" style="height: 300px"></div>
                                            <br>
                                            <h3 style="color: #ff425c; width: 100%; text-align: center; font-weight: bold;">PERSONAS POR TIPO DE ACCESO REGIMEN EN SALUD</h3>
                                            <div id="grafica_regimen_salud" style="height: 300px"></div>
                                            <br>
                                            <h3 style="width: 100%; color: #ff425c; text-align: center; font-weight: bold;">PORCENTAJE DE PERSONAS QUE POSEE ALGUNA DISCAPACIDAD</h3>
                                            <br>
                                            <table style="width: 100%" class="situacion_laboral" v-if="salud != null">
                                                <thead>
                                                    <tr>
                                                        <th>Discapacidad</th>
                                                        <th>Numero de personas</th>
                                                        <th>%</th>
                                                    </tr>
                                                </thead>
                                                <tbody>
                                                    <tr v-for="(item, index) in salud.condicion_discapacidad" :key="index">
                                                        <td>{{ item.condicion_discapacidad }}</td>
                                                        <td>{{ item.cantidad }}</td>
                                                        <td>{{ ((item.cantidad / salud.cantidad_poblacion) * 100).toFixed(2) }}</td>  
                                                    </tr>
                                                </tbody>
                                                <thead>
                                                    <tr>
                                                        <th>Total</th>
                                                        <th>{{ salud.cantidad_poblacion }}</th>
                                                        <th>100%</th>
                                                    </tr>
                                                </thead>
                                            </table>
                                            <br>
                                            <br>
                                            <h3 style="width: 100%; text-align: center; font-weight: bold;">NUMERO DE PERSONAS POR DISCAPACIDAD</h3>
                                            <div id="grafica_discapacidad" style="height: 300px"></div>
                                        </div>
                                    </div>
                                    <div class="col-lg-2"></div>  
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
        </div>


        <div class="modal fade text-left" id="modalPagina1" tabindex="-1" role="dialog" aria-labelledby="myModalLabel12" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered" role="document">
                <div class="modal-content">
                    <div class="modal-header bg-warning white">
                        <h4 class="modal-title" id="myModalLabel12"><i class="fa fa-file-excel"></i> Exportar Datos a Excel</h4>
                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                            <span aria-hidden="true">&times;</span>
                        </button>
                    </div>
                    <div class="modal-body">
                        <h2><i class="fa fa-arrow-right"></i> Seleccione un indicador</h2>
                        <hr>
                        <div class="row">
                            <div v-for="(item, index) in opciones_exportar" :key="index" class="col-lg-6" style="margin-bottom: 20px">
                                <input v-model="tipoExportarExcelSeleccionado" type="radio" :id="'control_e'+index" name="selecte1" :value='item.value' >
                                <label class="lradio" :for="'control_e'+index">
                                    <p>{{ item.label }}</p>
                                </label>
                            </div>
                        </div>
                    </div>
                    <div class="modal-footer">
                        <button type="button" class="btn grey btn-outline-danger" data-dismiss="modal">Cerrar</button>
                        <button type="button" @click="exportarDatosExcel" class="btn btn-outline-success">Exportar</button>
                    </div>
                </div>
            </div>
        </div>

    </div>
</template>
<script>
import Vue3Html2pdf from 'vue3-html2pdf'
import * as dashboardService from "../../services/dashboard_service";
import * as am4core from "@amcharts/amcharts4/core";
import * as am4charts from "@amcharts/amcharts4/charts";
import am4themes_animated from "@amcharts/amcharts4/themes/animated";
import Loading from 'vue3-loading-overlay';
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';

import Swal from 'sweetalert2';

am4core.useTheme(am4themes_animated);
export default {
    components: {
        Vue3Html2pdf,
        Loading
    },
    data() {
        return {
            htmlToPdfOptions: {
                margin: 0.3,

                filename: `informe_general.pdf`,

                image: {
                    type: 'jpeg', 
                    quality: 0.98
                },

                enableLinks: false,

                html2canvas: {
                    scale: 1,
                    useCORS: true
                },

                jsPDF: {
                    unit: 'in',
                    format: 'B4',
                    orientation: 'portrait'
                },
            },
            datos_edad: [],
            numero_personas: 0,
            datos_edad_general: null,
            chart_edad: null,
            chart_sexo: null,
            datos_desplazados: null,
            chart_desplazados: null,
            char_estado_civil: null,
            char_estado_civil_por: null,
            poblacion_e_activa: null,
            poblacion_e_inactiva: null,
            char_poblacion_activa: null,
            char_poblacion_inactiva: null,
            chart_ocupaciones: null,
            situacion_laboral: null,
            loading: false,
            datos_nivel_academico: [],
            personas_nivel_educativo: 0,
            chart_nivel_academico: null,
            chart_nivel_academico_sexo: null,
            datos_por_ocupacion: null,
            posesion_vivienda: null,
            chart_posesion: null,
            posesion_baldios: null,
            cultivos: null,
            cultivosPorLinea : [],
            servicios: [],
            tipo_vivienda: null,
            cantidad_viviendas: null,
            afrocolombianos_negritudes: null,
            salud: null,
            headerAfro: '<table style="width: 100%" class="tabla_informe"><tbody><tr><td colspan="2"><img src="/imagenes/header.png" style="width: 100%" alt=""></td></tr><tr><td colspan="2">Caracterización de población afrodescendiente</td></tr><tr><td><strong>El paso - Cesar</strong></td><td><strong>Informe:</strong> Afrocolombianos y Negritudes</td></tr></tbody></table><hr>',
            headerSocio: '<table style="width: 100%" class="tabla_informe"><tbody><tr><td colspan="2"><img src="/imagenes/header.png" style="width: 100%" alt=""></td></tr><tr><td colspan="2">Caracterización de población afrodescendiente</td></tr><tr><td><strong>El paso - Cesar</strong></td><td><strong>Informe:</strong> Sociodemografico </td></tr></tbody></table><hr>',
            headerEmpleo: '<table style="width: 100%" class="tabla_informe"><tbody><tr><td colspan="2"><img src="/imagenes/header.png" style="width: 100%" alt=""></td></tr><tr><td colspan="2">Caracterización de población afrodescendiente</td></tr><tr><td><strong>El paso - Cesar</strong></td><td><strong>Informe:</strong> Empleo y Educación </td></tr></tbody></table><hr>',
            headerVivienda: '<table style="width: 100%" class="tabla_informe"><tbody><tr><td colspan="2"><img src="/imagenes/header.png" style="width: 100%" alt=""></td></tr><tr><td colspan="2">Caracterización de población afrodescendiente</td></tr><tr><td><strong>El paso - Cesar</strong></td><td><strong>Informe:</strong> Vivienda y Hogar </td></tr></tbody></table><hr>',
            headerSalud: '<table style="width: 100%" class="tabla_informe"><tbody><tr><td colspan="2"><img src="/imagenes/header.png" style="width: 100%" alt=""></td></tr><tr><td colspan="2">Caracterización de población afrodescendiente</td></tr><tr><td><strong>El paso - Cesar</strong></td><td><strong>Informe:</strong> Salud </td></tr></tbody></table><hr>',
            exportar_pagina_1: [
                    {
                        "label": "POBLACIÓN POR CONCEJO",
                        "value": "e11"
                    },
                    {
                        "label": "POBLACIÓN POR ETNIA",
                        "value": "e12"
                    },
                    {
                        "label": "POBLACIÓN POR PRÁCTICAS CULTURALES Y RELIGIOSAS",
                        "value": "e13"
                    },
                    {
                        "label": "POBLACIÓN QUE HABLA Y NO HABLA ALGUNA LENGUA AFROCOLOMBIANA",
                        "value": "e14"
                    },
            ],
            exportar_pagina_2: [
                    {
                        "label": "POBLACIÓN POR GRUPO DE EDAD",
                        "value": "e21"
                    },
                    {
                        "label": "POBLACIÓN POR GENERO",
                        "value": "e22"
                    },
                    {
                        "label": "POBLACIÓN DESPLAZADA",
                        "value": "e23"
                    },
                    {
                        "label": "POBLACIÓN POR ESTADO CIVIL",
                        "value": "e24"
                    },
            ],
            exportar_pagina_3: [
                    {
                        "label": "POBLACIÓN ECONÓMICAMENTE ACTIVA",
                        "value": "e31"
                    },
                    {
                        "label": "POBLACIÓN ECONÓMICAMENTE ACTIVA",
                        "value": "e32"
                    },
                    {
                        "label": "POBLACIÓN POR OCUPACIÓN",
                        "value": "e33"
                    },
                    {
                        "label": "POBLACIÓN POR SITUACIÓN LABORAL",
                        "value": "e34"
                    },
                    {
                        "label": "POBLACIÓN POR NIVEL ACADÉMICO",
                        "value": "e35"
                    },
            ],
            exportar_pagina_4: [
                    {
                        "label": "VIVIENDAS CON A SERVICIOS BÁSICOS",
                        "value": "e41"
                    },
                    {
                        "label": "FAMILIAS POR TIPO DE VIVIENDA",
                        "value": "e42"
                    },
                    {
                        "label": "FAMILIAS POR POSESIÓN DE VIVIENDA PROPIA",
                        "value": "e43"
                    },
                    {
                        "label": "TENENCIA DE TIERRAS",
                        "value": "e44"
                    },
                    {
                        "label": "PRINCIPALES CULTIVOS",
                        "value": "e45"
                    },
            ],
            exportar_pagina_5: [
                    {
                        "label": "ESTADO DE SALUD GENERAL DE LA POBLACIÓN",
                        "value": "e51"
                    },
                    {
                        "label": "PERSONAS POR TIPO DE ACCESO A LA SALUD",
                        "value": "e52"
                    },
                    {
                        "label": "PERSONAS POR TIPO DE ACCESO REGIMEN EN SALUD",
                        "value": "e53"
                    },
                    {
                        "label": "LISTA DE PERSONAS QUE POSEE ALGUNA DISCAPACIDAD",
                        "value": "e54"
                    }
            ],
            opciones_exportar: [],
            tipoExportarExcelSeleccionado: ""
        }
    },
    mounted() {
        this.consultarDatos();
    },
    methods: {
        generateReport () {
            this.loading = true;
            setTimeout(()=>{
                $("#pestana_sociodemografico").printThis({
                    debug: false,    
                    pageTitle: "",    
                    importCSS: true,            
                    importStyle: true,        
                    printContainer: true,
                    header: this.headerSocio,
                    afterPrint: this.loading = false,
                });
            }, 100)
        },
        generateReport2 () {
            this.loading = true;
            setTimeout(()=>{
                $("#pestana_empleo_educacion").printThis({
                    debug: false,       
                    importCSS: true,            
                    importStyle: true,        
                    printContainer: true,
                    afterPrint: this.loading = false,
                    header: this.headerEmpleo
                });
            }, 100)
        },
        generateReport3 () {
            this.loading = true;
            setTimeout(()=>{
                $("#pestana_vivienda_hogar").printThis({
                    debug: false,       
                    importCSS: true,            
                    importStyle: true,        
                    printContainer: true,
                    afterPrint: this.loading = false,
                    header: this.headerVivienda
                });
            }, 100)
        },
        generateReportAfro() {
            this.loading = true;
            setTimeout(()=>{
                $("#pestana_afrocolombianos").printThis({
                    debug: false,     
                    header: this.headerAfro,   
                    importCSS: true,            
                    importStyle: true,        
                    printContainer: true,
                    pageTitle: null,
                    afterPrint: this.loading = false,
                });
            }, 100)
        },
        generateReportSalud() {
            this.loading = true;
            setTimeout(()=>{
                $("#pestana_salud").printThis({
                    debug: false,       
                    importCSS: true,            
                    importStyle: true, 
                    printContainer: true,  
                    afterPrint: this.loading = false,
                    header: this.headerSalud
                });
            }, 100)     
        },
        seleccionarOpcionesExportar(opcion){
            if(opcion == 1){
                this.opciones_exportar = this.exportar_pagina_1;
            }

            if(opcion == 2){
                this.opciones_exportar = this.exportar_pagina_2;
            }

            if(opcion == 3){
                this.opciones_exportar = this.exportar_pagina_3;
            }

            if(opcion == 4){
                this.opciones_exportar = this.exportar_pagina_4;
            }

            if(opcion == 5){
                this.opciones_exportar = this.exportar_pagina_5;
            }
        },
        async  exportarDatosExcel(){
            if(this.tipoExportarExcelSeleccionado == ""){
                Swal.fire({
                    width: 400,
                    icon: "error",
                    title: "Seleccione un indicador",
                });
            }else{
                await dashboardService.exportarExcel(this.tipoExportarExcelSeleccionado).then(respuesta => {
                    this.tipoExportarExcelSeleccionado = "";
                });
            }
        },
        async consultarDatos(){
            this.loading = true;
            await dashboardService.datosInforme().then(respuesta => {
                this.datos_edad_general = respuesta.data;
                this.datos_edad = respuesta.data.piramide_edad;
                this.numero_personas = respuesta.data.numero_personas;
                this.datos_desplazados = respuesta.data.datos_desplazados;
                this.generarGraficoEdad();
                this.generarGraficoSexo();
                this.generarGraficoDesplazados();
                this.generarGraficoEstadiCivilPie();
                this.poblacion_e_activa = respuesta.data.poblacion_e_activa;
                this.poblacion_e_inactiva = respuesta.data.poblacion_e_inactiva;
                this.generarGraficoPAE();
                this.generarGraficoPEI();
                this.datos_nivel_academico = respuesta.data.nivel_educativo;
                this.datos_nivel_academico_sexo = respuesta.data.nivel_educativo_sexo;
                this.personas_nivel_educativo = respuesta.data.personas_nivel_educativo;
                this.generarGraficoNivelAcademico();
                this.generarGraficoNivelAcademicoSexo();
                this.datos_por_ocupacion = respuesta.data.por_ocupacion;
                this.situacion_laboral = respuesta.data.situacion_laboral;
                this.generarGraficoPrincipalesOcupaciones();
                this.generarGraficoSituacionLaboral();
                this.posesion_vivienda = respuesta.data.posesion_vivienda;
                this.generarGraficoPosesion();
                this.posesion_baldios = respuesta.data.posesion_baldios[0];
                this.cultivos = respuesta.data.cultivos;
                this.cultivosPorLinea = this.agruparPorLinea(this.cultivos);
                this.generarGraficaLinea();
                setTimeout(()=>{
                    this.generarGraficasLinea();
                }, 1000);

                this.servicios = respuesta.data.servicios;
                this.generarGraficasServicios();
                this.tipo_vivienda = respuesta.data.tipo_vivienda;
                this.cantidad_viviendas = respuesta.data.cantidad_viviendas;
                this.generarGraficoTipoVivienda();
                this.afrocolombianos_negritudes = respuesta.data.afrocolombianos_negritudes;
                this.generarGraficoConcejo();
                this.generarGraficoEtnia();
                this.generarGraficoPracticas();
                this.generarGraficoLenguaHabla();
                this.salud = respuesta.data.salud;
                this.generarGraficoEstadoSalud();
                this.generarGraficoDiscapacidad();
                this.generarGraficaAccesoSalud();
                this.generarGraficaRegimenSalud();
                this.loading = false;
            });
        },
        agruparPorLinea(data) {
            const groupedByLinea = data.reduce((result, currentItem) => {
            const key = currentItem.linea;

            const groupIndex = result.findIndex(group => group[0].linea === key);

            if (groupIndex === -1) {
                result.push([currentItem]);
            } else {
                result[groupIndex].push(currentItem);
            }

            return result;
            }, []);

            return groupedByLinea;
        },
        generarGraficoEdad() {
            var chart = am4core.create("grafico_edad_informe", am4charts.XYChart);
            chart.data = [];
        
            var maximo = 0;

            for (let index = 0; index <= 15; index++) {
                const element = this.datos_edad[index];
                chart.data.push({
                    "age": element[2],
                    "male": (-1) * element[0],
                    "female": element[1]
                });

                if(element[0] >= maximo){
                    maximo = element[0];
                }

                if(element[1] >= maximo){
                    maximo = element[1];
                }
            }

            // Use only absolute numbers
            chart.numberFormatter.numberFormat = "#.#s";

            // Create axes
            var categoryAxis = chart.yAxes.push(new am4charts.CategoryAxis());
            categoryAxis.dataFields.category = "age";
            categoryAxis.renderer.grid.template.location = 0;
            categoryAxis.renderer.inversed = true;
            categoryAxis.renderer.minGridDistance = 10;
            

            var valueAxis = chart.xAxes.push(new am4charts.ValueAxis());
            valueAxis.extraMin = 0.1;
            valueAxis.extraMax = 0.1;
            valueAxis.min = -1 * (maximo);
            valueAxis.max = maximo;
            valueAxis.renderer.minGridDistance = 40;
            valueAxis.renderer.ticks.template.length = 5;
            valueAxis.renderer.ticks.template.disabled = false;
            valueAxis.renderer.ticks.template.strokeOpacity = 0.4;
            valueAxis.renderer.labels.template.adapter.add("text", function(text) {
            return text == "Male" || text == "Female" ? text : text;
            })

            // Create series
            var male = chart.series.push(new am4charts.ColumnSeries());
            male.dataFields.valueX = "male";
            male.dataFields.categoryY = "age";
            male.clustered = false;

            var maleLabel = male.bullets.push(new am4charts.LabelBullet());
            maleLabel.label.text = "{valueX}";
            maleLabel.label.hideOversized = false;
            maleLabel.label.truncate = false;
            maleLabel.label.horizontalCenter = "right";
            maleLabel.label.dx = -10;

            var female = chart.series.push(new am4charts.ColumnSeries());
            female.dataFields.valueX = "female";
            female.dataFields.categoryY = "age";
            female.clustered = false;

            var femaleLabel = female.bullets.push(new am4charts.LabelBullet());
            femaleLabel.label.text = "{valueX}";
            femaleLabel.label.hideOversized = false;
            femaleLabel.label.truncate = false;
            femaleLabel.label.horizontalCenter = "left";
            femaleLabel.label.dx = 10;
           
            var maleRange = valueAxis.axisRanges.create();
            maleRange.value = -1 * (maximo);
            maleRange.endValue = 0;
            maleRange.label.text = "Masculino";
            maleRange.label.fill = chart.colors.list[0];
            maleRange.label.dy = 20;
            maleRange.label.fontWeight = '600';
            maleRange.grid.strokeOpacity = 0;
            maleRange.grid.stroke = male.stroke;

            var femaleRange = valueAxis.axisRanges.create();
            femaleRange.value = 0;
            femaleRange.endValue = maximo;
            femaleRange.label.text = "Femenino";
            femaleRange.label.fill = chart.colors.list[1];
            femaleRange.label.dy = 20;
            femaleRange.label.fontWeight = '600';
            femaleRange.grid.strokeOpacity = 0;
            femaleRange.grid.stroke = female.stroke;

            this.chart_edad = chart;
        },
        generarGraficoSexo() {
            var chart = am4core.create("grafica_sexo_informe", am4charts.PieChart);

            // Add data
            chart.data = [
                {
                    "country": "Masculino",
                    "litres": this.datos_edad_general.numero_masculino,
                    "color": am4core.color("#ff7588") // Specify the color for Masculino
                },
                {
                    "country": "Femenino",
                    "litres": this.datos_edad_general.numero_femenino,
                    "color": am4core.color("#ffa87d") // Specify the color for Femenino
                }
            ];

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;


            // Set the slice colors based on the "color" field in the data
            pieSeries.slices.template.propertyFields.fill = "color";

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.chart_sexo = chart;
        },
        generarGraficoDesplazados(){
            var chart = am4core.create('grafico_desplazados', am4charts.XYChart)
            chart.colors.step = 2;

            chart.legend = new am4charts.Legend()
            chart.legend.position = 'top'
            chart.legend.paddingBottom = 20
            chart.legend.labels.template.maxWidth = 95

            var xAxis = chart.xAxes.push(new am4charts.CategoryAxis())
            xAxis.dataFields.category = 'category'
            xAxis.renderer.cellStartLocation = 0.1
            xAxis.renderer.cellEndLocation = 0.9
            xAxis.renderer.grid.template.location = 0;

            var yAxis = chart.yAxes.push(new am4charts.ValueAxis());
            yAxis.min = 0;

            function createSeries(value, name) {
                var series = chart.series.push(new am4charts.ColumnSeries())
                series.dataFields.valueY = value
                series.dataFields.categoryX = 'category'
                series.name = name

                series.events.on("hidden", arrangeColumns);
                series.events.on("shown", arrangeColumns);

                var bullet = series.bullets.push(new am4charts.LabelBullet())
                bullet.interactionsEnabled = false
                bullet.dy = 30;
                bullet.label.text = '{valueY}'
                bullet.label.fill = am4core.color('#ffffff')

                return series;
            }

            chart.data = [
                {
                    category: this.datos_desplazados.desplazados_m15[2],
                    first: this.datos_desplazados.desplazados_m15[0],
                    second: this.datos_desplazados.desplazados_m15[1],
                },
                {
                    category: this.datos_desplazados.desplazados_15a64[2],
                    first: this.datos_desplazados.desplazados_15a64[0],
                    second: this.datos_desplazados.desplazados_15a64[1],
                },
                {
                    category: this.datos_desplazados.desplazados_mas65[2],
                    first: this.datos_desplazados.desplazados_mas65[0],
                    second: this.datos_desplazados.desplazados_mas65[1],
                },
            ]


            createSeries('first', 'Masculino');
            createSeries('second', 'Femenino');

            function arrangeColumns() {
                var series = chart.series.getIndex(0);
                var w = 1 - xAxis.renderer.cellStartLocation - (1 - xAxis.renderer.cellEndLocation);
                if (series.dataItems.length > 1) {
                    var x0 = xAxis.getX(series.dataItems.getIndex(0), "categoryX");
                    var x1 = xAxis.getX(series.dataItems.getIndex(1), "categoryX");
                    var delta = ((x1 - x0) / chart.series.length) * w;
                    if (am4core.isNumber(delta)) {
                        var middle = chart.series.length / 2;

                        var newIndex = 0;
                        chart.series.each(function(series) {
                            if (!series.isHidden && !series.isHiding) {
                                series.dummyData = newIndex;
                                newIndex++;
                            }
                            else {
                                series.dummyData = chart.series.indexOf(series);
                            }
                        })
                        var visibleCount = newIndex;
                        var newMiddle = visibleCount / 2;

                        chart.series.each(function(series) {
                            var trueIndex = chart.series.indexOf(series);
                            var newIndex = series.dummyData;

                            var dx = (newIndex - trueIndex + middle - newMiddle) * delta

                            series.animate({ property: "dx", to: dx }, series.interpolationDuration, series.interpolationEasing);
                            series.bulletsContainer.animate({ property: "dx", to: dx }, series.interpolationDuration, series.interpolationEasing);
                        })
                    }
                }
            }

            this.chart_desplazados = chart;

        },
        generarGraficoEstadiCivilPie() {
            var chart = am4core.create("grafico_estado_civil_por", am4charts.PieChart);

            // Add data
            chart.data = [];

            for (let index = 0; index <= 5; index++) {
                const element = this.datos_edad_general.estado_civil[index];
                chart.data.push({
                    "country": element[0],
                    "litres": element[1]
                });
            }


            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.char_estado_civil_por = chart;
        },
        generarGraficoPAE() {
            var chart = am4core.create("grafica_pae", am4charts.PieChart);

            // Add data
            chart.data = [
                {
                    "country": "Masculino",
                    "litres": this.poblacion_e_activa[0],
                },
                {
                    "country": "Femenino",
                    "litres": this.poblacion_e_activa[1],
                }
            ];

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.char_poblacion_activa = chart;
        },
        generarGraficoPEI() {
            var chart = am4core.create("grafica_pei", am4charts.PieChart);

            // Add data
            chart.data = [
                {
                    "country": "Masculino",
                    "litres": this.poblacion_e_inactiva[0],
                },
                {
                    "country": "Femenino",
                    "litres": this.poblacion_e_inactiva[1],
                }
            ];

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.char_poblacion_inactiva = chart;
        },
        generarGraficoPrincipalesOcupaciones() {
            var chart = am4core.create("grafica_principales_ocupaciones", am4charts.PieChart);

            // Add data
            chart.data = [];

            if(this.datos_por_ocupacion.length < 5){
                this.datos_por_ocupacion.forEach(element => {
                    chart.data.push({
                        "country": element.ocupacion,
                        "litres": element.cantidad,
                    });
                });
            }else{
                for (let index = 0; index < this.datos_por_ocupacion.length; index++) {
                    var element = this.datos_por_ocupacion[index];
                    if(element.ocupacion != "NO APLICA"){
                        chart.data.push({
                            "country": element.ocupacion,
                            "litres": element.cantidad,
                        });
                    }

                    if(chart.data.length == 5){
                        break;
                    }   
                }
            }


            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 110;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.chart_ocupaciones = chart;
        },
        generarGraficoSituacionLaboral() {
            var chart = am4core.create("grafica_situacion_laboral", am4charts.PieChart);

            // Add data
            chart.data = [
                {
                    "country": "Trabajo Formal",
                    "litres": this.situacion_laboral.trabajo_formal,
                },
                {
                    "country": "Trabajo Inormal",
                    "litres": this.situacion_laboral.trabajo_informal,
                }
            ];


            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 110;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

        },
        generarGraficoNivelAcademico() {
            var chart = am4core.create("grafica_nivel_academico", am4charts.PieChart);

            // Add data
            chart.data = [];

            this.datos_nivel_academico.forEach(element => {
                chart.data.push({
                    "country": element.descripcion,
                    "litres": element.cantidad,
                });
            });

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.char_poblacion_inactiva = chart;
        },
        generarGraficoNivelAcademicoSexo() {
            var chart = am4core.create('grafica_nivel_academico_sexo', am4charts.XYChart);
            chart.colors.step = 2;

            chart.legend = new am4charts.Legend();
            chart.legend.position = 'top';
            chart.legend.paddingBottom = 20;
            chart.legend.labels.template.maxWidth = 95;

            // Interchange x-axis and y-axis
            var yAxis = chart.xAxes.push(new am4charts.ValueAxis());
            yAxis.min = 0;

            var xAxis = chart.yAxes.push(new am4charts.CategoryAxis());
            xAxis.dataFields.category = 'category';
            xAxis.renderer.cellStartLocation = 0.1;
            xAxis.renderer.cellEndLocation = 0.9;
            xAxis.renderer.grid.template.location = 0;
            xAxis.renderer.minGridDistance = 25;

            function createSeries(value, name) {
                var series = chart.series.push(new am4charts.ColumnSeries());
                series.dataFields.valueX = value; 
                series.dataFields.categoryY = 'category';
                series.name = name;

                var bullet = series.bullets.push(new am4charts.LabelBullet());
                bullet.interactionsEnabled = false;
                bullet.dx = 20;
                bullet.label.text = '{valueX}';
                bullet.label.fill = am4core.color('#000000');

                return series;
            }

            chart.data = [];

            this.datos_nivel_academico_sexo.forEach(element => {
                chart.data.push({
                    category: element.nivel_educativo,
                    first: element.hombres,
                    second: element.mujeres,
                });
            });

            createSeries('first', 'Hombres');
            createSeries('second', 'Mujeres');

        },
        generarGraficoPosesion() {
            var chart = am4core.create("grafica_posesion", am4charts.PieChart);

            // Add data
            chart.data = [
                {
                    "country": "POSEEN",
                    "litres": this.posesion_vivienda.poseen,
                },
                {
                    "country": "NO POSEEN",
                    "litres": this.posesion_vivienda.no_poseen,
                }
            ];

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.char_poblacion_inactiva = chart;
        },
        generarGraficaLinea(){
            var chart = am4core.create("grafica_linea_economica", am4charts.PieChart);
            chart.data = [];
            var nombre_linea = "";
            var cantidad_por_linea = 0;

            this.cultivosPorLinea.forEach(element => { 
                cantidad_por_linea = 0;

                nombre_linea = element[0].linea;  
                element.forEach(element2 => {
                    cantidad_por_linea += element2.area_destinada
                });

                chart.data.push({
                    "country": nombre_linea,
                    "litres": cantidad_por_linea
                });
            });

            chart.innerRadius = am4core.percent(50);

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeWidth = 2;
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 130;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;
        },
        generarGraficasLinea() {
            var index = 0;
            this.cultivosPorLinea.forEach(element => {
                var chart = am4core.create("grafica_linea_"+index, am4charts.PieChart);
                chart.data = [];

                element.forEach(element2 => {
                    chart.data.push({
                        "country": element2.actividad,
                        "litres": element2.area_destinada
                    });
                });

                // Set inner radius
                chart.innerRadius = am4core.percent(50);

                // Add and configure Series
                var pieSeries = chart.series.push(new am4charts.PieSeries());
                pieSeries.dataFields.value = "litres";
                pieSeries.dataFields.category = "country";
                pieSeries.slices.template.stroke = am4core.color("#fff");
                pieSeries.slices.template.strokeWidth = 2;
                pieSeries.slices.template.strokeOpacity = 1;

                pieSeries.labels.template.maxWidth = 110;
                pieSeries.labels.template.wrap = true;

                // This creates initial animation
                pieSeries.hiddenState.properties.opacity = 1;
                pieSeries.hiddenState.properties.endAngle = -90;
                pieSeries.hiddenState.properties.startAngle = -90;

                index++;
            });
        },
        generarGraficasServicios() {
            var chart = am4core.create('grafico_servicios', am4charts.XYChart)
            chart.colors.step = 2;

            chart.legend = new am4charts.Legend()
            chart.legend.position = 'top'
            chart.legend.paddingBottom = 20
            chart.legend.labels.template.maxWidth = 95

            var xAxis = chart.xAxes.push(new am4charts.CategoryAxis())
            xAxis.dataFields.category = 'category'
            xAxis.renderer.cellStartLocation = 0.1
            xAxis.renderer.cellEndLocation = 0.9
            xAxis.renderer.grid.template.location = 0;

            var yAxis = chart.yAxes.push(new am4charts.ValueAxis());
            yAxis.min = 0;

            function createSeries(value, name) {
                var series = chart.series.push(new am4charts.ColumnSeries())
                series.dataFields.valueY = value
                series.dataFields.categoryX = 'category'
                series.name = name

                series.events.on("hidden", arrangeColumns);
                series.events.on("shown", arrangeColumns);

                var bullet = series.bullets.push(new am4charts.LabelBullet())
                bullet.interactionsEnabled = false
                bullet.dy = 30;
                bullet.label.text = '{valueY}'
                bullet.label.fill = am4core.color('#ffffff')

                return series;
            }

            chart.data = [
                {
                    category: this.servicios[0][2],
                    first: this.servicios[0][0],
                    second: this.servicios[0][1],
                },
                {
                    category: this.servicios[1][2],
                    first: this.servicios[1][0],
                    second: this.servicios[1][1],
                },
                {
                    category: this.servicios[2][2],
                    first: this.servicios[2][0],
                    second: this.servicios[2][1],
                },
                {
                    category: this.servicios[3][2],
                    first: this.servicios[3][0],
                    second: this.servicios[3][1],
                },
                {
                    category: this.servicios[4][2],
                    first: this.servicios[4][0],
                    second: this.servicios[4][1],
                },
                {
                    category: this.servicios[5][2],
                    first: this.servicios[5][0],
                    second: this.servicios[5][1],
                },
            ]

            
            createSeries('first', 'Si');
            createSeries('second', 'No');

            function arrangeColumns() {
                var series = chart.series.getIndex(0);
                var w = 1 - xAxis.renderer.cellStartLocation - (1 - xAxis.renderer.cellEndLocation);
                if (series.dataItems.length > 1) {
                    var x0 = xAxis.getX(series.dataItems.getIndex(0), "categoryX");
                    var x1 = xAxis.getX(series.dataItems.getIndex(1), "categoryX");
                    var delta = ((x1 - x0) / chart.series.length) * w;
                    if (am4core.isNumber(delta)) {
                        var middle = chart.series.length / 2;

                        var newIndex = 0;
                        chart.series.each(function(series) {
                            if (!series.isHidden && !series.isHiding) {
                                series.dummyData = newIndex;
                                newIndex++;
                            }
                            else {
                                series.dummyData = chart.series.indexOf(series);
                            }
                        })
                        var visibleCount = newIndex;
                        var newMiddle = visibleCount / 2;

                        chart.series.each(function(series) {
                            var trueIndex = chart.series.indexOf(series);
                            var newIndex = series.dummyData;

                            var dx = (newIndex - trueIndex + middle - newMiddle) * delta

                            series.animate({ property: "dx", to: dx }, series.interpolationDuration, series.interpolationEasing);
                            series.bulletsContainer.animate({ property: "dx", to: dx }, series.interpolationDuration, series.interpolationEasing);
                        })
                    }
                }
            }

            this.chart_desplazados = chart; 
        },
        generarGraficoTipoVivienda() {
            var chart = am4core.create("grafica_tipo_vivienda", am4charts.PieChart);

            // Add data
            chart.data = [];

            this.tipo_vivienda.forEach(element => {
                chart.data.push(
                    {
                        "country": element.tipo_vivienda,
                        "litres": element.cantidad,
                    }
                );
            });

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);

            this.char_poblacion_inactiva = chart;
        },
        generarGraficoConcejo() {
            var chart = am4core.create("grafica_concejo", am4charts.PieChart);

            // Add data
            chart.data = [];

            this.afrocolombianos_negritudes.por_concejo.forEach(element => {
                chart.data.push(
                    {
                        "country": element.concejo,
                        "litres": element.cantidad,
                    }
                );
            });

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 160;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);
        },
        generarGraficoEtnia() {
            var chart = am4core.create("grafica_etnia", am4charts.PieChart);

            // Add data
            chart.data = [];

            this.afrocolombianos_negritudes.por_etnia.forEach(element => {
                chart.data.push(
                    {
                        "country": element.etnia,
                        "litres": element.cantidad,
                    }
                );
            });

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 160;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);
        },
        generarGraficoPracticas() {
            var chart = am4core.create("grafica_practicas", am4charts.XYChart);
           
            var categoryAxis = chart.yAxes.push(new am4charts.CategoryAxis());
            categoryAxis.renderer.grid.template.location = 0;
            categoryAxis.dataFields.category = "category";
            categoryAxis.renderer.minGridDistance = 1;
            categoryAxis.renderer.inversed = true;
            categoryAxis.renderer.grid.template.disabled = true;

            var valueAxis = chart.xAxes.push(new am4charts.ValueAxis());
            valueAxis.min = 0;

            var series = chart.series.push(new am4charts.ColumnSeries());
            series.dataFields.categoryY = "category";
            series.dataFields.valueX = "first";
            series.tooltipText = "{valueX.value}"
            series.columns.template.strokeOpacity = 0;
            series.columns.template.column.cornerRadiusBottomRight = 5;
            series.columns.template.column.cornerRadiusTopRight = 5;

            var labelBullet = series.bullets.push(new am4charts.LabelBullet())
            labelBullet.label.horizontalCenter = "left";
            labelBullet.label.dx = 10;
            labelBullet.label.text = "{values.valueX.workingValue}";
            labelBullet.label.fill = am4core.color("#ffffff"); // Set label text color to white
            labelBullet.label.fontWeight = "bold";
            labelBullet.locationX = 1;

            // as by default columns of the same series are of the same color, we add adapter which takes colors from chart.colors color set
            series.columns.template.adapter.add("fill", function(fill, target){
                return chart.colors.getIndex(target.dataItem.index);
            });

            categoryAxis.sortBySeries = series;

            chart.data = [];

            this.afrocolombianos_negritudes.por_practicas_religiosas.forEach(element => {
                chart.data.push({
                    category: element.practicas_religiosas,
                    first: element.cantidad,
                });
            });

        },
        generarGraficoLenguaHabla() {
            var chart = am4core.create("grafica_lengua", am4charts.PieChart);

            // Add data
            chart.data = [];

            this.afrocolombianos_negritudes.cual_lengua_habla.forEach(element => {
                chart.data.push(
                    {
                        "country": element.cual_lengua,
                        "litres": element.cantidad,
                    }
                );
            });

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 160;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);
        },
        generarGraficoEstadoSalud() {
            var chart = am4core.create("grafica_estado_salud", am4charts.PieChart);

            // Add data
            chart.data = [];

            this.salud.estado_salud.forEach(element => {
                chart.data.push(
                    {
                        "country": element.estado_salud,
                        "litres": element.cantidad,
                    }
                );
            });

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 160;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);
        },
        generarGraficoDiscapacidad() {
            var chart = am4core.create("grafica_discapacidad", am4charts.XYChart);
           
            var categoryAxis = chart.yAxes.push(new am4charts.CategoryAxis());
            categoryAxis.renderer.grid.template.location = 0;
            categoryAxis.dataFields.category = "category";
            categoryAxis.renderer.minGridDistance = 1;
            categoryAxis.renderer.inversed = true;
            categoryAxis.renderer.grid.template.disabled = true;

            var valueAxis = chart.xAxes.push(new am4charts.ValueAxis());
            valueAxis.min = 0;

            var series = chart.series.push(new am4charts.ColumnSeries());
            series.dataFields.categoryY = "category";
            series.dataFields.valueX = "first";
            series.tooltipText = "{valueX.value}"
            series.columns.template.strokeOpacity = 0;
            series.columns.template.column.cornerRadiusBottomRight = 5;
            series.columns.template.column.cornerRadiusTopRight = 5;

            var labelBullet = series.bullets.push(new am4charts.LabelBullet())
            labelBullet.label.horizontalCenter = "left";
            labelBullet.label.dx = 10;
            labelBullet.label.text = "{values.valueX.workingValue}";
            labelBullet.label.fill = am4core.color("#ffffff"); // Set label text color to white
            labelBullet.label.fontWeight = "bold";
            labelBullet.locationX = 1;

            // as by default columns of the same series are of the same color, we add adapter which takes colors from chart.colors color set
            series.columns.template.adapter.add("fill", function(fill, target){
                return chart.colors.getIndex(target.dataItem.index);
            });

            categoryAxis.sortBySeries = series;

            chart.data = [];

            this.salud.condicion_discapacidad.forEach(element => {
                chart.data.push({
                    category: element.condicion_discapacidad,
                    first: element.cantidad,
                });
            });

        },
        generarGraficaAccesoSalud() {
            var chart = am4core.create("grafica_acceso_salud", am4charts.XYChart);

            chart.data = [];

            this.salud.acceso_salud.forEach(element => {
                chart.data.push({
                    category: element.acceso_salud,
                    first: element.cantidad,
                });
            });

            // Create axes

            var categoryAxis = chart.xAxes.push(new am4charts.CategoryAxis());
            categoryAxis.dataFields.category = "category";
            categoryAxis.renderer.grid.template.location = 0;
            categoryAxis.renderer.minGridDistance = 30;

            categoryAxis.renderer.labels.template.adapter.add("dy", function(dy, target) {
            if (target.dataItem && target.dataItem.index & 2 == 2) {
                return dy + 25;
            }
            return dy;
            });

            var valueAxis = chart.yAxes.push(new am4charts.ValueAxis());

            // Create series
            var series = chart.series.push(new am4charts.ColumnSeries());
            series.dataFields.valueY = "first";
            series.dataFields.categoryX = "category";
            series.name = "first";
            series.columns.template.tooltipText = "{categoryX}: [bold]{valueY}[/]";
            series.columns.template.fillOpacity = 1;
            series.columns.template.column.cornerRadiusBottomRight = 5;
            series.columns.template.column.cornerRadiusTopRight = 5;

            // Add labelBullet
            var labelBullet = series.bullets.push(new am4charts.LabelBullet());
            labelBullet.label.text = "{valueY}";
            labelBullet.label.dy = -15; 

            // as by default columns of the same series are of the same color, we add adapter which takes colors from chart.colors color set
            series.columns.template.adapter.add("fill", function(fill, target){
                return chart.colors.getIndex(target.dataItem.index);
            });

            var columnTemplate = series.columns.template;
            columnTemplate.strokeWidth = 2;
            columnTemplate.strokeOpacity = 1;                                                       
        },
        generarGraficaRegimenSalud() {
            var chart = am4core.create("grafica_regimen_salud", am4charts.PieChart);

            // Add data
            chart.data = [];

            this.salud.regimen.forEach(element => {
                chart.data.push(
                    {
                        "country": element.regimen,
                        "litres": element.cantidad,
                    }
                );
            });

            var pieSeries = chart.series.push(new am4charts.PieSeries());
            pieSeries.dataFields.value = "litres";
            pieSeries.dataFields.category = "country";
            pieSeries.slices.template.stroke = am4core.color("#fff");
            pieSeries.slices.template.strokeOpacity = 1;

            pieSeries.labels.template.maxWidth = 160;
            pieSeries.labels.template.wrap = true;

            pieSeries.hiddenState.properties.opacity = 1;
            pieSeries.hiddenState.properties.endAngle = -90;
            pieSeries.hiddenState.properties.startAngle = -90;

            chart.hiddenState.properties.radius = am4core.percent(0);                                                  
        }
    }
}
</script>
<style>
    .pdf-preview {
        position: fixed !important;
        width: 100% !important;
        min-width: 600px !important;
        height: 100% !important;
        top: 0px !important;
        z-index: 9999999 !important;
        left: 50% !important;
        transform: translateX(-50%) !important;
        border-radius: 5px !important;
        box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.2823529412) !important;
        display: flex;
        justify-content: center;
        align-items: center;
        background-color: rgba(61, 61, 61, 0.514);
    }

    .pdf-preview iframe {
        width: 60% !important;
        height: 80% !important;
    }

    .vue-html2pdf .pdf-preview button {
        position: absolute !important;
        top: 16px !important;
        left: 94% !important;
        width: 58px !important;
        height: 55px !important;
        background: #ff0023 !important;
        border: 0 !important;
        box-shadow: 0px 0px 10px rgba(0, 0, 0, 0.2823529412) !important;
        border-radius: 50% !important;
        color: #fff !important;
        display: flex !important;
        align-items: center !important;
        justify-content: center !important;
        font-size: 30px !important;
        cursor: pointer !important;
        font-weight: bold !important;
        border: 4px solid #ffff !important;
    }

    .tabla_informe {
        border-collapse: collapse;
    }

    .tabla_informe thead {
        font-weight: bold;
    }

    .tabla_informe td, .tabla_informe th {
        border: 1px solid grey;
        padding: 5px !important;
        text-align: center;
        font-size: 14px;
    }

    .tabla_informe td, .tabla_informe th {
        border: 1px solid rgb(173, 173, 173) !important;
        padding: 10px;
    }

    thead th {
       background-color: #5df3c5 !important;
    }

    .situacion_laboral th, .situacion_laboral td{
        border-right: 1px solid grey;
        border-left: 1px solid grey;
        border-bottom: 1px solid grey;
        border-top: 1px solid grey;
        text-align: center;
    }

    .page-break {
        page-break-before: always; /* Add a page break before elements with this class */
    }

    @media print {
        @page {
            size: A4;
        }
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
        display: flex;
        justify-content: center;
        align-items: center;
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
</style>