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
            <div class="card">
                <div class="card-header">
                    <h4 class="card-title" style="font-weight: bold">FORMULARIO DE CARACTERIZACIÓN POBLACIÓN AFRODESCENDIENTES</h4>
                </div>
                <div class="card-content">
                    <div class="card-body">
                        <ul class="nav nav-tabs nav-underline" role="tablist">
                            <li style ="font-size: 14px" class="nav-item" data-toggle="tooltip" data-placement="top" title="" data-original-title="Información personal">
                                <a class="nav-link active" id="baseIcon-tab21" data-toggle="tab" aria-controls="tabIcon21" href="#tabIcon21" role="tab" aria-selected="false">
                                    <i style ="font-size: 25px" class="fas fa-user-edit"></i>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item" data-toggle="tooltip" data-placement="top" title="" data-original-title="Origen e Identidad">
                                <a class="nav-link" id="baseIcon-tab22" data-toggle="tab" aria-controls="tabIcon22" href="#tabIcon22" role="tab" aria-selected="false">
                                    <i style ="font-size: 25px" class="fas fa-globe-americas"></i>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item" data-toggle="tooltip" data-placement="top" title="" data-original-title="Educación">
                                <a class="nav-link" id="baseIcon-tab23" data-toggle="tab" aria-controls="tabIcon23" href="#tabIcon23" role="tab" aria-selected="false">
                                    <i style ="font-size: 25px" class="fas fa-school"></i>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item" data-toggle="tooltip" data-placement="top" title="" data-original-title="Situación laboral">
                                <a class="nav-link" id="baseIcon-tab24" data-toggle="tab" aria-controls="tabIcon24" href="#tabIcon24" role="tab" aria-selected="false">
                                    <i style ="font-size: 25px" class="fas fa-user-tie"></i>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item" data-toggle="tooltip" data-placement="top" title="" data-original-title="Salud">
                                <a class="nav-link" id="baseIcon-tab25" data-toggle="tab" aria-controls="tabIcon25" href="#tabIcon25" role="tab" aria-selected="false">
                                    <i style ="font-size: 25px" class="fas fa-hospital-user"></i>
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item" data-toggle="tooltip" data-placement="top" title="" data-original-title="Cultura y Tradiciones">
                                <a class="nav-link" id="baseIcon-tab26" data-toggle="tab" aria-controls="tabIcon26" href="#tabIcon26" role="tab" aria-selected="false">
                                    <i style ="font-size: 25px" class="fas fa-people-carry"></i> 
                                </a>
                            </li>
                            <li style ="font-size: 14px" class="nav-item" data-toggle="tooltip" data-placement="top" title="" data-original-title="Vivienda y Hogar">
                                <a class="nav-link" id="baseIcon-tab27" data-toggle="tab" aria-controls="tabIcon27" href="#tabIcon27" role="tab" aria-selected="false">
                                    <i style ="font-size: 25px" class="fas fa-house-user"></i>
                                </a>
                            </li>
                        </ul>
                        <div class="tab-content px-1 pt-1">
                            <div class="tab-pane active" id="tabIcon21" role="tabpanel" aria-labelledby="baseIcon-tab21">
                                <br>
                                <h2 style="color: #2f8a96"> INFORMACIÓN PERSONAL </h2>
                                <hr>
                                <br>
                                <div class="row">
                                    <div class="col-lg-6">
                                        <label for=""><strong>Rol del individuo</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="informacion_personal.rol" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value=""><strong>Seleccione el rol del individuo</strong></option>
                                                <option value="Jefe de hogar">Jefe de Hogar</option>
                                                <option value="Integrante">Integrante de hogar</option>
                                            </select>
                                            <div class="form-control-position">
                                                <i style ="font-size: 14px" class="feather icon-unlock  font-medium-4"></i>
                                            </div>
                                        </fieldset>
                                    </div>                                    
                                    <div class="col-lg-6" v-if="informacion_personal.rol != 'Jefe de hogar'">
                                        <label for=""><strong>Jefe de Hogar</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input data-toggle="modal" data-target="#exampleModalCenter"  v-model="informacion_personal.nombre_jefe" style="cursor: pointer" type="text" class="form-control form-control-lg mb-1"  placeholder="Seleccione su jefe de hogar">
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-user-tie"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12">
                                        <label for=""><strong>Parentesco</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="informacion_personal.parentesco" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value=""><strong>Seleccione el parentesco del individuo</strong></option>
                                                <option value="Esposo(a), Compañero(a)">Esposo(a), Compañero(a)</option>
                                                <option value="Hijo(a)">Hijo(a)</option>
                                                <option value="Hijastro(a)">Hijastro(a)</option>
                                                <option value="Yerno(a)">Yerno(a)</option>
                                                <option value="Cuñado(a)">Cuñado(a)</option>
                                                <option value="Nieto(a)">Nieto(a)</option>
                                                <option value="Padre">Padre</option>
                                                <option value="Padre soltero">Padre soltero</option>
                                                <option value="Madre">Madre</option>
                                                <option value="Madre Soltera">Madre Soltera</option>
                                                <option value="Padrastro o Madrastra">Padrastro o Madrastra</option>
                                                <option value="Abuelo(a)">Abuelo(a)</option>
                                                <option value="Suegro(a)">Suegro(a)</option>
                                                <option value="Hermano(a)">Hermano(a)</option>
                                                <option value="Hermanastro(a)">Hermanastro(a)</option>
                                                <option value="Otro Pariente">Otro Pariente</option>
                                            </select>
                                            <div class="form-control-position">
                                                <i style ="font-size: 14px" class="feather icon-unlock  font-medium-4"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12">
                                        <div class="row">
                                            <div class="col-lg-4">
                                                <label for=""><strong>Departamento</strong></label>
                                                <fieldset class="form-group position-relative has-icon-left">
                                                    <select :disabled="informacion_personal.rol != 'Jefe de hogar'" @change="consultarMunicipiosIP()" v-model="informacion_personal.departamento" class="form-control form-control-lg mb-1" name="" id="">
                                                        <option value="">Seleccione una opción</option>
                                                        <option v-for="(item, index) in departamentos" :key="index" :value="item.codigo">{{ item.descripcion }}</option>
                                                    </select>                                            
                                                    <div class="form-control-position">
                                                        <i style="font-size: 19px" class="fas fa-map-marked-alt"></i>
                                                    </div>
                                                </fieldset>
                                            </div>
                                            <div class="col-lg-4">
                                                <label for=""><strong>Municipio</strong></label>
                                                <fieldset class="form-group position-relative has-icon-left">
                                                    <select :disabled="informacion_personal.rol != 'Jefe de hogar'" @change="listarCorregimientosPorMuni()" v-model="informacion_personal.municipio" class="form-control form-control-lg mb-1" name="" id="">
                                                        <option value="">Seleccione una opción</option>
                                                        <option v-for="(item, index) in municipiosIP" :key="index" :value="item.id">{{ item.descripcion }}</option>
                                                    </select>                                            
                                                    <div class="form-control-position">
                                                        <i style="font-size: 19px" class="fas fa-city"></i>
                                                    </div>
                                                </fieldset>
                                            </div>
                                            <div class="col-lg-4">
                                                <label for=""><strong>Corregimiento</strong></label>
                                                <fieldset class="form-group position-relative has-icon-left">
                                                    <select :disabled="informacion_personal.rol != 'Jefe de hogar'" v-model="informacion_personal.corregimiento" class="form-control form-control-lg mb-1" name="" id="">
                                                        <option value="">Seleccione una opción</option>
                                                        <option v-for="(item, index) in corregimientos" :key="index" :value="item.id">{{ item.nombre }}</option>
                                                    </select>                                            
                                                    <div class="form-control-position">
                                                        <i style="font-size: 19px" class="fas fa-city"></i>
                                                    </div>
                                                </fieldset>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-lg-12">
                                        <label for=""><strong>Dirección</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input :disabled="informacion_personal.rol != 'Jefe de hogar'" v-model="informacion_personal.direccion" type="text" class="form-control form-control-lg mb-1"  placeholder="Dirección">
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-map-marked-alt"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12">
                                        <label for=""><strong>Nombre completo</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="informacion_personal.nombre_completo" type="text" class="form-control form-control-lg mb-1"  placeholder="Nombre Completo">
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-id-card"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Fecha de nacimiento</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="informacion_personal.fecha_nacimiento" type="date" class="form-control form-control-lg mb-1"  placeholder="Fecha de nacimiento">
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-calendar-alt"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Tipo de identificación</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="informacion_personal.tipo_identificacion" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Tipo de Identificación</option>
                                                <option value="CC">Cédula de Ciudadanía</option>
                                                <option value="TI">Tarjeta de  Identidad</option>
                                                <option value="Pasaporte">Pasaporte</option>
                                                <option value="Registro Civil">Registro Civil</option>
                                                <option value="Permiso Especial de Permanencia ">Permiso Especial de Permanencia </option>
                                                <option value="No Tiene ">No Tiene </option>
                                            </select>                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-id-card-alt"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Número de identificación</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="informacion_personal.numero_identificacion" type="number" class="form-control form-control-lg mb-1"  placeholder="Número de Identificación">
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-hashtag"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Dirección de residencia</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input :disabled="informacion_personal.rol != 'Jefe de hogar'" v-model="informacion_personal.direccion_residencia" type="text" class="form-control form-control-lg mb-1"  placeholder="Dirección de residencia">
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-map-marked-alt"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Sexo</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="informacion_personal.sexo" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Femenino">Femenino</option>
                                                <option value="Masculino">Masculino</option>
                                                <option value="Otros">Otros</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-venus-mars"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6" v-if="informacion_personal.sexo == 'Otros'">
                                        <label for=""><strong>¿Cual?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="informacion_personal.cual_sexo" type="text" class="form-control form-control-lg mb-1"  placeholder="Cual sexo?">                                    
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Orientación sexual</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="informacion_personal.orientacion_sexual" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Heterosexual">Heterosexual</option>
                                                <option value="Homosexual">Homosexual</option>
                                                <option value="Bisexual">Bisexual</option>
                                                <option value="Pansexual">Pansexual</option>
                                                <option value="Asexual">Asexual</option>
                                                <option value="Queer">Queer</option>
                                                <option value="Cuarsosexual">Cuarsosexual </option>
                                                <option value="Otro">Otro</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-users"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Estado civil</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="informacion_personal.estado_civil" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Soltero">Soltero</option>
                                                <option value="Casado">Casado</option>
                                                <option value="Divorciado">Divorciado</option>
                                                <option value="Viudo(a)">Viudo(a)</option>
                                                <option value="Unión Libre">Unión Libre</option>
                                                <option value="Otro">Otro</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-user-friends"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6" v-if="informacion_personal.estado_civil == 'Otro'">
                                        <label for=""><strong>¿Cual?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="informacion_personal.cual_estado_civil" type="text" class="form-control form-control-lg mb-1"  placeholder="Cual estado civil?">                                    
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>Hace cuanto tiempo vive en el municipio? (Meses)</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="informacion_personal.tiempo_municipo" type="number" class="form-control form-control-lg mb-1"  placeholder="Tiempo en meses">                                    
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-calendar-day"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>¿Es usted desplazado?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="informacion_personal.desplazado" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-people-arrows"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12 text-right" style="padding-top: 20px">
                                        <button @click="validarCamposInformacionPersonal()" class="btn btn-success">Guardar y continuar <i style="font-size: 13px" class="fas fa-arrow-right"></i></button>
                                    </div>
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon22" role="tabpanel" aria-labelledby="baseIcon-tab22">
                                <br>
                                <h2 style="color: #2f8a96"> ORIGEN E IDENTIDAD </h2>
                                <hr>
                                <br>
                                <label for=""><strong>1. Lugar de Nacimiento </strong></label>
                                <div class="row" style="margin-top: 37px">
                                    <div class="col-lg-4">
                                        <label for=""><strong>País</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="origen_identidad.pais" class="form-control form-control-lg mb-1" name="" id="">
                                                <option v-for="(item, index) in paises_latinoamericanos" :key="index" :value="item">{{item}}</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-globe-africa"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-4">
                                        <label for=""><strong>Departamento</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select :disabled="origen_identidad.pais !== 'COLOMBIA'" @change="consultarMunicipios()" v-model="origen_identidad.departamento" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option v-for="(item, index) in departamentos" :key="index" :value="item.codigo">{{ item.descripcion }}</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-map-marked-alt"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-4">
                                        <label for=""><strong>Municipio</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select :disabled="origen_identidad.pais !== 'COLOMBIA'" v-model="origen_identidad.municipio" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option v-for="(item, index) in municipios" :key="index" :value="item.id">{{ item.descripcion }}</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-city"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                </div>
                                <br><br>
                                <label for=""><strong>2. Selecciona la etnia a la que pertenece</strong></label>
                                <div class="row" style="margin-top: 37px">
                                    <div class="col-lg-4" style="margin-bottom: 20px">
                                        <input v-model="origen_identidad.etnia" type="radio" id="control_01" name="select" value="Negritudes" >
                                        <label class="lradio" for="control_01">
                                            <p> Negritudes</p>
                                        </label>
                                    </div>
                                    <div class="col-lg-4" style="margin-bottom: 20px">
                                        <input v-model="origen_identidad.etnia" type="radio" id="control_02" name="select" value="Afrocolombianos">
                                        <label class="lradio" for="control_02">
                                            <p>Afrocolombianos</p>
                                        </label>
                                    </div>
                                    <br>
                                    <div class="col-lg-12 text-right" style="padding-top: 20px">
                                        <button @click="eliminarActive('baseIcon-tab21', 'tabIcon21')" class="btn btn-danger" style="margin-right: 20px"><i style="font-size: 13px" class="fas fa-arrow-left"></i> Anterior</button>
                                        <button @click="validarCamposOrigenIdentidad()" class="btn btn-success">Guardar y continuar <i style="font-size: 13px" class="fas fa-arrow-right"></i></button>
                                    </div>
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon23" role="tabpanel" aria-labelledby="baseIcon-tab23">
                                <br>
                                <h2 style="color: #2f8a96"> EDUCACIÓN </h2>
                                <hr>
                                <br>
                                <div class="row">
                                    <div class="col-lg-6">
                                        <label for=""><strong>1. Nivel educativo </strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="educacion.nivel_educativo" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option v-for="(item, index) in escolaridades" :key="index" :value="item.id">{{ item.descripcion }}</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-school"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6" v-if="educacion.nivel_educativo == 23">
                                        <label for=""><strong>¿Cual?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="educacion.cual_nivel_educativo" type="text" class="form-control form-control-lg mb-1"  placeholder="Cual nivel educativo?">                                    
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>2. ¿ha recibido educación específica sobre la cultura afrocolombiana?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select  v-model="educacion.educacion_especifica" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>3. ¿ha enfrentado barreras para acceder a la educación?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select  v-model="educacion.barreras_educacion" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>4. ¿ha experimentado discriminación racial o étnica en el entorno educativo? </strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select  v-model="educacion.descriminacion_racial" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                                <option value="N.A">N.A</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>5. ¿recibe algún tipo de apoyo educativo adicional?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select  v-model="educacion.apoyo_educativo" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                                <option value="N.A">N.A</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12 text-right" style="padding-top: 20px">
                                        <button @click="eliminarActive('baseIcon-tab22', 'tabIcon22')" class="btn btn-danger" style="margin-right: 20px"><i style="font-size: 13px" class="fas fa-arrow-left"></i> Anterior</button>
                                        <button @click="validarCamposOrigenEducacion()" class="btn btn-success">Guardar y continuar <i style="font-size: 13px" class="fas fa-arrow-right"></i></button>
                                    </div>
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon24" role="tabpanel" aria-labelledby="baseIcon-tab24">
                                <br>
                                <h2 style="color: #2f8a96"> SITUACIÓN LABORAL </h2>
                                <hr>
                                <br>    
                                <div class="row">
                                    <div class="col-lg-6">
                                        <label for="input-datalist"><strong>1. Ocupación</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <div class="form-group">
                                                <input v-model="situacion_laboral.ocupacion" type="text" class="form-control form-control-lg mb-1" placeholder="Seleccione una ocupación" list="list-timezone" id="input-datalist">
                                                <datalist id="list-timezone">
                                                    <option v-for="(item, index) in ocupaciones" :key="index">{{ item.descripcion }}</option>
                                                </datalist>
                                            </div>
                                            <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                <i style ="font-size: 19px" class="fas fa-suitcase"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for="input-datalist"><strong>2. Situación laboral</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="situacion_laboral.situacion_laboral" class="form-control form-control-lg mb-1">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Empleado/a tiempo completo">Empleado/a tiempo completo</option>
                                                <option value="Empleado/a tiempo parcial">Empleado/a tiempo parcial</option>
                                                <option value="Desempleado/a (en búsqueda activa de empleo)">Desempleado/a (en búsqueda activa de empleo)</option>
                                                <option value="Desempleado/a (sin búsqueda activa de empleo en el momento)">Desempleado/a (sin búsqueda activa de empleo en el momento)</option>
                                                <option value="Empresario/a dueño/a de negocio">Empresario/a dueño/a de negocio</option>
                                                <option value="Trabajador/a independiente">Trabajador/a independiente</option>
                                                <option value="Jubilado/a (retirado/a del empleo remunerado)">Jubilado/a (retirado/a del empleo remunerado)</option>
                                                <option value="Voluntario/a (sin compensación financiera)">Voluntario/a (sin compensación financiera)</option>
                                                <option value="Trabajador informal">Trabajador informal</option>
                                                <option value="N.A">N.A</option>
                                            </select>
                                            <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                <i style ="font-size: 19px" class="fas fa-suitcase"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>3. ¿ha experimentado discriminación en el ámbito laboral?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="situacion_laboral.discriminacion_laboral" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                                <option value="N.A">N.A</option>
                                                <option value="No responde">No responde</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for="input-datalist"><strong>4. ¿Se generan Ingresos?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="situacion_laboral.ingreso_mensual" class="form-control form-control-lg mb-1">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                                <option value="N.A">N.A</option>
                                                <option value="No responde">No responde</option>
                                            </select>
                                            <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                <i style ="font-size: 19px" class="fas fa-dollar-sign"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12 text-right" style="padding-top: 20px">
                                        <button @click="eliminarActive('baseIcon-tab23', 'tabIcon23')" class="btn btn-danger" style="margin-right: 20px"><i style="font-size: 13px" class="fas fa-arrow-left"></i> Anterior</button>
                                        <button @click="validarCamposSituacionLaboral()" class="btn btn-success">Guardar y continuar <i style="font-size: 13px" class="fas fa-arrow-right"></i></button>
                                    </div>
                                </div>                        
                            </div>
                            <div class="tab-pane" id="tabIcon25" role="tabpanel" aria-labelledby="baseIcon-tab25">
                                <br>
                                <h2 style="color: #2f8a96"> SALUD </h2>
                                <hr>
                                <br>
                                <div class="row">
                                    <div class="col-lg-6">
                                        <label for="input-datalist"><strong>1. Estado de salud general</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="salud.estado_salud" class="form-control form-control-lg mb-1">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Bueno">Bueno</option>
                                                <option value="Regular">Regular</option>
                                                <option value="Malo">Malo</option>
                                                <option value="No informa">No informa</option>
                                            </select>
                                            <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                <i style ="font-size: 19px" class="fas fa-heart"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for="input-datalist"><strong>2. ¿Posee alguna condición de discapacidad? </strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="salud.condicion_discapacidad" class="form-control form-control-lg mb-1">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Discapacidad Física">Discapacidad Física</option>
                                                <option value="Discapacidad Sensorial">Discapacidad Sensorial</option>
                                                <option value="Discapacidad Cognitiva">Discapacidad Cognitiva</option>
                                                <option value="Discapacidad Psicosocial o Mental">Discapacidad Psicosocial o Mental</option>
                                                <option value="Discapacidad del Habla o del Lenguaje">Discapacidad del Habla o del Lenguaje</option>
                                                <option value="Discapacidad Múltiple o Combinada">Discapacidad Múltiple o Combinada</option>
                                                <option value="Discapacidad Temporal">Discapacidad Temporal</option>
                                                <option value="Ninguna">Ninguna</option>
                                            </select>
                                            <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                <i style ="font-size: 19px" class="fas fa-blind"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for="input-datalist"><strong>3. Acceso a servicios de salud</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="salud.acceso_salud" class="form-control form-control-lg mb-1">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Frecuente">Frecuente</option>
                                                <option value="Ocasional">Ocasional</option>
                                                <option value="No tiene acceso">No tiene acceso</option>
                                            </select>
                                            <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                <i style ="font-size: 19px" class="fas fa-medkit"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for="input-datalist"><strong>4. Régimen de salud al que pertenece</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="salud.regimen" class="form-control form-control-lg mb-1">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Contributivo">Contributivo</option>
                                                <option value="Subsidiado">Subsidiado</option>
                                                <option value="Vinculado/ Beneficiario">Vinculado/ Beneficiario</option>
                                                <option value="Régimen Especial">Régimen Especial</option>
                                                <option value="Particular o Privado">Particular o Privado</option>
                                                <option value="No conoce">No conoce</option>
                                                <option value="Población pobre no asegurada">Población pobre no asegurada</option>
                                            </select>
                                            <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                <i style ="font-size: 19px" class="fas fa-laptop-medical"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>5. ¿Ha enfrentado discriminación en el ámbito de la salud?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="salud.discriminacion_salud" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                                <option value="N.A">N.A</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12 text-right" style="padding-top: 20px">
                                        <button @click="eliminarActive('baseIcon-tab24', 'tabIcon24')" class="btn btn-danger" style="margin-right: 20px"><i style="font-size: 13px" class="fas fa-arrow-left"></i> Anterior</button>
                                        <button @click="validarCamposSalud()" class="btn btn-success">Guardar y continuar <i style="font-size: 13px" class="fas fa-arrow-right"></i></button>
                                    </div>
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon26" role="tabpanel" aria-labelledby="baseIcon-tab26">
                                <br>
                                <h2 style="color: #2f8a96"> CULTURA Y TRADICIONES </h2>
                                <hr>
                                <br>
                                <div class="row">
                                    <div class="col-lg-6">
                                        <label for=""><strong>1. ¿Participa en actividades culturales afrocolombianas?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="culturaTradiciones.practica_actividades" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="Nunca">Nunca</option>
                                                <option value="Casi nunca">Casi nunca</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>    
                                    <div class="col-lg-6">
                                        <label for=""><strong>2. ¿habla alguna lengua afrocolombiana?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="culturaTradiciones.habla_lengua" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Si">Si</option>
                                                <option value="No">No</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>    
                                    <div class="col-lg-6" v-if="culturaTradiciones.habla_lengua == 'Si'">
                                        <label for=""><strong>¿Cual lengua habla?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="culturaTradiciones.cual_lengua" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Creole">Creole</option>
                                                <option value="Palenquero">Palenquero</option>
                                            </select>  
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div> 
                                    <div class="col-lg-6">
                                        <label for=""><strong>3. Prácticas culturales y religiosas</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="culturaTradiciones.practicas_religiosas" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Música y Danza">Música y Danza</option>
                                                <option value="Sincretismo Religioso">Sincretismo Religioso</option>
                                                <option value="Santería y Palenque">Santería y Palenque</option>
                                                <option value="Prácticas médicas tradicionales">Prácticas médicas tradicionales</option>
                                                <option value="Ninguna">Ninguna</option>
                                            </select>  
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>    
                                    <div class="col-lg-12">
                                        <label for=""><strong>4. ¿A cual concejo comunitario pertenece?</strong></label>
                                        <br><br>
                                        <div class="row">
                                            <div class="col-lg-4" style="margin-bottom: 20px">
                                                <input v-model="culturaTradiciones.concejo" type="radio" id="control_41" name="select2" value='CONCEJO COMUNITARIO "JULIO CESAR ALTAMAR MUÑOS"' >
                                                <label class="lradio" for="control_41">
                                                    <p> CONCEJO COMUNITARIO "JULIO CESAR ALTAMAR MUÑOS"</p>
                                                </label>
                                            </div>
                                            <div class="col-lg-4" style="margin-bottom: 20px">
                                                <input v-model="culturaTradiciones.concejo" type="radio" id="control_42" name="select2" value='CONCEJO COMUNITARIO "ALEJO DURAN"'>
                                                <label class="lradio" for="control_42">
                                                    <p>CONCEJO COMUNITARIO "ALEJO DURAN"</p>
                                                </label>
                                            </div>
                                        </div>
                                    </div>      
                                    <div class="col-lg-12 text-right" style="padding-top: 20px">
                                        <button @click="eliminarActive('baseIcon-tab25', 'tabIcon25')" class="btn btn-danger" style="margin-right: 20px"><i style="font-size: 13px" class="fas fa-arrow-left"></i> Anterior</button>
                                        <button @click="validarCamposCulturaTradiciones()" class="btn btn-success">Guardar y continuar <i style="font-size: 13px" class="fas fa-arrow-right"></i></button>
                                    </div>   
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon27" role="tabpanel" aria-labelledby="baseIcon-tab27">
                                <br>
                                <h2 style="color: #2f8a96"> VIVIENDA Y HOGAR </h2>
                                <hr>
                                <br>
                                <div class="row">
                                    <div class="col-lg-6">
                                        <label for=""><strong>1. Tipo de vivienda</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="viviendaHogar.tipo_vivienda" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Casa">Casa</option>
                                                <option value="Apartamento">Apartamento</option>
                                                <option value="Improvisada">Improvisada</option>
                                                <option value="Finca">Finca</option>
                                                <option value="Parcela">Parcela</option>
                                                <option value="Otro">Otro</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-home"></i>
                                            </div>
                                        </fieldset>
                                    </div> 
                                    <div class="col-lg-6" v-if="viviendaHogar.tipo_vivienda != 'Otro'">

                                    </div>
                                    <div class="col-lg-6" v-if="viviendaHogar.tipo_vivienda == 'Otro'">
                                        <label for=""><strong>¿Cual?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="viviendaHogar.cual_tipo_vivienda" type="text" class="form-control form-control-lg mb-1"  placeholder="Cual tipo de vivienda?">                                    
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-question"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-6">
                                        <label for=""><strong>2. Tenencia de la vivienda</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="viviendaHogar.tenencia" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Propia">Propia</option>
                                                <option value="Alquilada">Alquilada</option>
                                                <option value="Familiar">Familiar</option>
                                                <option value="Copropiedad">Copropiedad</option>
                                                <option value="Tenencia Gratuita">Tenencia Gratuita</option>
                                                <option value="N.A">N.A</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-home"></i>
                                            </div>
                                        </fieldset>
                                    </div>  
                                    <div class="col-lg-6">
                                        <label for=""><strong>3. Número de personas que conforman el hogar</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="viviendaHogar.numero_personas_hogar" class="form-control form-control-lg mb-1" placeholder="# de personas" type="number">                                           
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-users"></i>
                                            </div>
                                        </fieldset>
                                    </div>    
                                    <div class="col-lg-12">
                                        <label for=""><strong>4. Acceso a servicios básicos</strong></label>
                                        <br><br>
                                        <div class="row">
                                            <div class="col-lg-2" style="margin-bottom: 20px; height: 100px">
                                                <input @click="viviendaHogar.electricidad = viviendaHogar.electricidad === 'Si' ? 'No' : 'Si'" type="checkbox" id="control_011" name="select2">
                                                <label class="lradio" for="control_011">
                                                    <p>Electricidad</p>
                                                </label>
                                            </div>
                                            <div class="col-lg-2" style="margin-bottom: 20px; height: 100px">
                                                <input @click="viviendaHogar.agua_potable = viviendaHogar.agua_potable === 'Si' ? 'No' : 'Si'" type="checkbox" id="control_012" name="select3">
                                                <label class="lradio" for="control_012">
                                                    <p>Agua potable</p>
                                                </label>
                                            </div>
                                            <div class="col-lg-2" style="margin-bottom: 20px; height: 100px">
                                                <input @click="viviendaHogar.alcantarillado = viviendaHogar.alcantarillado === 'Si' ? 'No' : 'Si'" type="checkbox" id="control_013" name="select4">
                                                <label class="lradio" for="control_013">
                                                    <p>Alcantarillado</p>
                                                </label>
                                            </div>
                                            <div class="col-lg-2" style="margin-bottom: 20px; height: 100px">
                                                <input @click="viviendaHogar.gas_natural = viviendaHogar.gas_natural === 'Si' ? 'No' : 'Si'" type="checkbox" id="control_014" name="select5">
                                                <label class="lradio" for="control_014">
                                                    <p>Gas natural</p>
                                                </label>
                                            </div>
                                            <div class="col-lg-2" style="margin-bottom: 20px; height: 100px">
                                                <input @click="viviendaHogar.aseo = viviendaHogar.aseo === 'Si' ? 'No' : 'Si'" type="checkbox" id="control_015" name="select5">
                                                <label class="lradio" for="control_015">
                                                    <p>Aseo</p>
                                                </label>
                                            </div>
                                            <div class="col-lg-2" style="margin-bottom: 20px; height: 100px">
                                                <input @click="viviendaHogar.otro = viviendaHogar.otro === 'Si' ? 'No' : 'Si'" type="checkbox" id="control_016" name="select5">
                                                <label class="lradio" for="control_016">
                                                    <p>Otro</p>
                                                </label>
                                            </div>
                                        </div>
                                    </div> 
                                    <div class="col-lg-6">
                                        <br>
                                        <label for=""><strong>5. ¿Cuantas personas trabajan en su hogar?</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <input v-model="viviendaHogar.numero_personas_trabajan" class="form-control form-control-lg mb-1" placeholder="# de personas" type="number">                                           
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-user-shield"></i>
                                            </div>
                                        </fieldset>
                                    </div> 
                                    <div class="col-lg-6">
                                        <br>
                                        <label for=""><strong>6. Ingresos mensuales por hogar</strong></label>
                                        <fieldset class="form-group position-relative has-icon-left">
                                            <select v-model="viviendaHogar.ingresos_mensuales_hogar" class="form-control form-control-lg mb-1" name="" id="">
                                                <option value="">Seleccione una opción</option>
                                                <option value="Menos de un salario mínimo">Menos de un salario mínimo </option>
                                                <option value="Entre 1 y 2 SMLV">Entre 1 y 2 SMLV</option>
                                                <option value="Entre 3 y 4 SMLV">Entre 3 y 4 SMLV</option>
                                                <option value="Mas de 5 SMLV">Mas de 5 SMLV</option>
                                            </select>                                            
                                            <div class="form-control-position">
                                                <i style="font-size: 19px" class="fas fa-dollar-sign"></i>
                                            </div>
                                        </fieldset>
                                    </div>
                                    <div class="col-lg-12" v-if="viviendaHogar.tipo_vivienda == 'Finca' || viviendaHogar.tipo_vivienda == 'Parcela'">
                                        <hr>
                                        <h2 style="color: #2f8a96"> TENENCIA DE TIERRAS</h2>
                                        <br>
                                        <div class="row">
                                            <div class="col-lg-6">
                                                <label for=""><strong>Posesión baldía</strong></label>
                                                    <fieldset class="form-group position-relative has-icon-left">
                                                        <select v-model="viviendaHogar.posecion_baldia" class="form-control form-control-lg mb-1" name="" id="">
                                                            <option value="">Seleccione una opción</option>
                                                            <option value="Si">Si</option>
                                                            <option value="No">No</option>
                                                        </select>                                            
                                                        <div class="form-control-position">
                                                            <i style="font-size: 19px" class="fas fa-question"></i>
                                                        </div>
                                                    </fieldset>
                                            </div>   
                                            <div class="col-lg-6">
                                                <label for=""><strong>Propiedad con título</strong></label>
                                                    <fieldset class="form-group position-relative has-icon-left">
                                                        <select v-model="viviendaHogar.propiedad_titulo" class="form-control form-control-lg mb-1" name="" id="">
                                                            <option value="">Seleccione una opción</option>
                                                            <option value="Si">Si</option>
                                                            <option value="No">No</option>
                                                        </select>                                            
                                                        <div class="form-control-position">
                                                            <i style="font-size: 19px" class="fas fa-question"></i>
                                                        </div>
                                                    </fieldset>
                                            </div> 
                                            <div class="col-lg-6">
                                                <label for=""><strong>Área Total (Metros cuadrados)</strong></label>
                                                    <fieldset class="form-group position-relative has-icon-left">
                                                        <input placeholder="ejemplo: 2000" v-model="viviendaHogar.area_total" class="form-control form-control-lg mb-1" type="number">                                                                                                    
                                                        <div class="form-control-position">
                                                            <i style="font-size: 19px" class="far fa-square"></i>
                                                        </div>
                                                    </fieldset>
                                            </div>   
                                            <div class="col-lg-12">
                                                <h3><strong>Actividades económicas</strong></h3>
                                            </div>
                                            <div class="col-lg-4">
                                                <label for="input-datalist"><strong>Linea economica</strong></label>
                                                <fieldset class="form-group position-relative has-icon-left">
                                                    <div class="form-group">
                                                        <select v-model="lineaSeleccionada" @change="seleccionarLinea()" class="form-control form-control-lg mb-1" name="" id="">
                                                            <option value="">Seleccione una opción</option>
                                                            <option v-for="(item, index) in lineas_economicas" :key="index" :value="index+1">{{ item.texto }}</option>
                                                        </select>  
                                                    </div>
                                                    <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                        <i style ="font-size: 19px" class="fab fa-pagelines"></i>
                                                    </div>
                                                </fieldset>
                                            </div> 
                                            <div class="col-lg-4">
                                                <label for="input-datalist"><strong>Actividad economica</strong></label>
                                                <fieldset class="form-group position-relative has-icon-left">
                                                    <div class="form-group">
                                                        <select v-model="actividadSeleccionada" class="form-control form-control-lg mb-1" name="" id="">
                                                            <option value="">Seleccione una opción</option>
                                                            <option v-for="(item, index) in datoslineaSeleccionada" :key="index" :value="item.texto">{{ item.texto }}</option>
                                                        </select>  
                                                    </div>
                                                    <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                        <i style ="font-size: 19px" class="fas fa-tree"></i>
                                                    </div>
                                                </fieldset>
                                            </div> 
                                            <div class="col-lg-3">
                                                <label for="input-datalist"><strong>Área destinada a la actividad (m<sup>2</sup>)</strong></label>
                                                <fieldset class="form-group position-relative has-icon-left">
                                                    <div class="form-group">
                                                        <input v-model="areaDestinada" type="number" class="form-control form-control-lg mb-1">
                                                    </div>
                                                    <div class="form-control-position" style="top: 9px; left: 4px !important;">
                                                        <i style ="font-size: 19px" class="fas fa-superscript"></i>
                                                    </div>
                                                </fieldset>
                                            </div> 
                                            <div class="col-lg-1" style="display: flex; justify-content: center; align-items: center">
                                                <button class="btn btn-success" @click="agregarActividad()">
                                                    <i class="fas fa-plus"></i>
                                                </button>
                                            </div> 
                                            <div class="col-lg-12" v-if="viviendaHogar.actividades_economicas.length > 0">
                                                <br>
                                                <table style="width: 100%">
                                                    <thead>
                                                        <tr>
                                                            <th>Linea</th>
                                                            <th>Actividad</th>
                                                            <th>Área destinada</th>
                                                            <th>Opciones</th>
                                                        </tr>
                                                    </thead>
                                                    <tbody>
                                                        <tr v-for="(item, index) in viviendaHogar.actividades_economicas" :key="index">
                                                            <td>{{ item.linea }}</td>
                                                            <td>{{ item.actividad }}</td>
                                                            <td>{{ item.area_destinada }}</td>
                                                            <td><button class="btn btn-danger" @click="eliminarActividad(index)"><i class="fas fa-trash-alt"></i></button></td>
                                                        </tr>
                                                    </tbody>
                                                </table>
                                            </div>
                                        </div>
                                    </div>
                                    <div class="col-lg-12 text-right" style="padding-top: 20px">
                                        <hr>
                                        <button @click="eliminarActive('baseIcon-tab26', 'tabIcon26')" class="btn btn-danger" style="margin-right: 20px"><i style="font-size: 13px" class="fas fa-arrow-left"></i> Anterior</button>
                                        <button @click="validarCamposViviendaHogar()" class="btn btn-success">Guardar y finalizar <i style="font-size: 13px" class="fas fa-arrow-right"></i></button>
                                    </div>    
                                </div>
                            </div>
                            <div class="tab-pane" id="tabIcon28" role="tabpanel" aria-labelledby="baseIcon-tab28">
                                <div style="padding: 95px" class="alert alert-primary text-center" role="alert">
                                    <h1 style="font-size: 3rem; font-weight: 500;">¡Proceso de caracterización finalizado con éxito!</h1>
                                    <hr>
                                    <button onclick="location.reload()" class="btn btn-success">Caracterizar Nuevo Individuo <i class="fas fa-user-plus"></i></button>
                                    <router-link to="/lista-caracterizados"><button style="margin-left: 20px" class="btn btn-warning">Ver Lista de Caracterizados <i class="fas fa-users"></i></button></router-link>
                                </div>
                            </div>
                        </div>
                    </div>
                </div>
            </div>
        </div>

        <!-- Modal -->
        <div class="modal fade" id="exampleModalCenter" tabindex="-1" role="dialog" aria-labelledby="exampleModalCenterTitle" aria-hidden="true">
            <div class="modal-dialog modal-dialog-centered modal-lg" role="document">
                <div class="modal-content">
                    <div class="modal-header">
                        <h2 class="modal-title" id="exampleModalCenterTitle">Jefes de Hogar Registrados</h2>
                        <button type="button" class="close" data-dismiss="modal" aria-label="Close">
                            <span aria-hidden="true">&times;</span>
                        </button>
                    </div>
                    <div class="modal-body">
                        <h4>Seleccione su jefe de hogar</h4>
                        <hr>
                        <table id="miTabla" style="width: 100%">
                            <thead>
                                <tr>
                                    <th>Tipo Identificación</th>
                                    <th>Identificación</th>
                                    <th>Nombre Completo</th>
                                </tr>
                            </thead>
                            <tbody>
                                <tr v-for="(item, index) in jefes_hogar" :key="index" @click="seleccionarJefe(item)">
                                    <td>{{ item.tipo_identificacion }}</td>
                                    <td>{{ item.identificacion }}</td>
                                    <td>{{ item.nombre_completo }}</td>
                                </tr>
                            </tbody>
                        </table>
                    </div>
                    <div class="modal-footer">
                        <button class="btn btn-success" data-dismiss="modal">Seleccionar</button>
                        <button type="button" class="btn btn-danger" data-dismiss="modal">Cerrar</button>
                    </div>
                </div>
            </div>
        </div>
    </div>
</template>
<script>
import Loading from 'vue3-loading-overlay';
import 'vue3-loading-overlay/dist/vue3-loading-overlay.css';
import * as usuarioService from "../../services/usuario_service";
import * as caracterizacionService from "../../services/caracterizacion_service";
import * as corregimientosService from "../../services/corregimientos";

import Swal from 'sweetalert2';

export default {
    props: ['identificacion_editar'],
    components: {
        Loading
    },
    mounted() {
       this.consultarValoresDefecto()
    },
    data() {
        return {
            informacion_personal: {
                rol: "",
                id_jefe: "",
                nombre_jefe: "",
                parentesco: "",
                direccion: "",
                nombre_completo: "",
                fecha_nacimiento: "",
                tipo_identificacion: "",
                numero_identificacion: "",
                direccion_residencia: "",
                sexo: "",
                cual_sexo: "",
                orientacion_sexual: "",
                estado_civil: "",
                cual_estado_civil: "",
                tiempo_municipo: "",
                desplazado: "",
                departamento: "",
                municipio: "",
                corregimiento: ""
            },
            errores: "",
            loading: false,
            departamentos: [],
            municipios: [],
            origen_identidad: {
                identificacion_individuo: null,
                pais: "COLOMBIA",
                departamento: "",
                municipio: "",
                etnia: ""
            },
            escolaridades: [],
            educacion: {
                identificacion_individuo: null,
                nivel_educativo: "",
                cual_nivel_educativo: "",
                educacion_especifica: "",
                barreras_educacion: "",
                descriminacion_racial: "",
                apoyo_educativo: ""
            },
            ocupaciones: [],
            situacion_laboral: {
                identificacion_individuo: null,
                ocupacion: "",
                situacion_laboral: "",
                discriminacion_laboral: "",
                ingreso_mensual: "",
            },
            salud: {
                identificacion_individuo: null,
                estado_salud: "",
                condicion_discapacidad: "",
                acceso_salud: "",
                regimen: "",
                discriminacion_salud: ""
            },
            culturaTradiciones: {
                identificacion_individuo: null,
                practica_actividades: "",
                habla_lengua: "",
                cual_lengua: "",
                practicas_religiosas: "",
                concejo: ""
            },
            viviendaHogar: {
                identificacion_jefe: null,
                tipo_vivienda: "",
                cual_tipo_vivienda: "",
                tenencia: "",
                numero_personas_hogar: "",
                electricidad: "No",
                agua_potable: "No",
                alcantarillado: "No",
                gas_natural: "No",
                aseo: "No",
                otro: "No",
                numero_personas_trabajan: "",
                ingresos_mensuales_hogar: "",
                posecion_baldia: "",
                propiedad_titulo: "",
                area_total: "",
                actividades_economicas: []
            },
            jefes_hogar: [],
            lineas_economicas: [
                { value: 1, texto: "Cultivos agrícolas" },
                { value: 2, texto: "Pecuarias - especie animal con perspectiva comercial" },
                { value: 3, texto: "Cultivos forestales con perspectiva comercial" }
            ],  
            linea1: [
                {texto: "Maíz"},
                {texto: "Trigo"},
                {texto: "Arroz"},
                {texto: "Cebada"},
                {texto: "Centeno"},
                {texto: "Alpiste"},
                {texto: "Sorgo"},
                {texto: "Mijo"},
                {texto: "Avena"},
                {texto: "Trigo Espelta"},
                {texto: "Trigo kamut"},
                {texto: "Frijoles rojo"},
                {texto: "Frijol negro"},
                {texto: "Frijol cabecita negra"},
                {texto: "Frijol"},
                {texto: "Arvejas"},
                {texto: "Habas"},
                {texto: "Lenteja"},
                {texto: "Garbanzo"},
                {texto: "Gandul"},
                {texto: "Alfalfa"},
                {texto: "Almorta"},
                {texto: "Soya"},
                {texto: "Canola"},
                {texto: "Girasol"},
                {texto: "Olivo"},
                {texto: "Maíz"},
                {texto: "Ajonjolí"},
                {texto: "Palma de aceite"},
                {texto: "Acelga"},
                {texto: "Ajo"},
                {texto: "Albahaca"},
                {texto: "Apio"},
                {texto: "Alcachofa"},
                {texto: "Brócoli"},
                {texto: "Calabacín"},
                {texto: "Calabaza"},
                {texto: "Cebolla"},
                {texto: "Col"},
                {texto: "Coliflor"},
                {texto: "Espinacas"},
                {texto: "Pepino"},
                {texto: "Perejil"},
                {texto: "Pimiento"},
                {texto: "Puerro"},
                {texto: "Rábano"},
                {texto: "Remolacha"},
                {texto: "Tomate"},
                {texto: "Zanahoria"},
                {texto: "Auyama"},
                {texto: "Plátano"},
                {texto: "Cana"},
                {texto: "coco"},
                {texto: "Aguacate"},
                {texto: "Naranja"},
                {texto: "Mango"},
                {texto: "Limón"},
                {texto: "Banano"},
                {texto: "Mandarina"},
                {texto: "Sandia"},
                {texto: "Papaya"},
                {texto: "Marañón"},
                {texto: "Guanábana"},
                {texto: "Guayaba"},
                {texto: "Piña"},
                {texto: "Lulo"},
                {texto: "Maracuyá"},
                {texto: "Granadilla"},
                {texto: "Curuba"},
                {texto: "Níspero"},
                {texto: "Borojó"},
                {texto: "Zapote"},
                {texto: "Carambola"},
                {texto: "Tomate de árbol"},
                {texto: "Melón"},
                {texto: "Nueces"},
                {texto: "Arracacha"},
                {texto: "Papa amarilla"},
                {texto: "Batata"},
                {texto: "Yuca"},
                {texto: "Papa criolla"},
                {texto: "Ñame"},
                {texto: "Papa sabanera"},
                {texto: "Papa"},
                {texto: "Jengibre"},
                {texto: "Malanga"},
                {texto: "Yacón"},
                {texto: "Ornamentales de exteriores"},
                {texto: "Ornamentales de Interiores"},
                {texto: "Otros floricultivos"},
                {texto: "Rosas"},
                {texto: "Girasoles"},
                {texto: "Crisantemo"},
                {texto: "Tulipan"},
                {texto: "Clavel"},
                {texto: "Gerbera"},
                {texto: "Gladiolo"},
                {texto: "Acroclinio"},
                {texto: "Aster"},
                {texto: "Caléndula"},
                {texto: "Lavanda"},
                {texto: "Aloe vera"},
                {texto: "Romero"},
                {texto: "Manzanilla"},
                {texto: "Menta"},
                {texto: "Calendula"},
                {texto: "Diente de leon"},
                {texto: "Albahaca"},
                {texto: "Algarrobo"},
                {texto: "Boldo"},
                {texto: "Acacia"},
                {texto: "Toronjil"},
                {texto: "Paja limón"},
                {texto: "Cidrón"},
                {texto: "Achiote"},
                {texto: "Arnica"},
                {texto: "Canela"},
                {texto: "Tomillo"},
                {texto: "Caña de azucar"},
                {texto: "Tabaco"},
                {texto: "Pasto"},
                {texto: "Café"},
                {texto: "Cacao"},
                {texto: "Otro"}
            ],
            linea2: [
                { texto: "Ganado vacuno o bovino" },
                { texto: "Ganado ovino" },
                { texto: "Ganado porcino" },
                { texto: "Ganado caprino" },
                { texto: "Ganado equino" },
                { texto: "Avicultura" },
                { texto: "Apicultura" },
                { texto: "Acuicultura" },
                { texto: "Helicicultura" },
                { texto: "Cunicultura" },
                { texto: "Sericicultura" },
                { texto: "Otros" }
            ],
            linea3: [
                {texto: "Chiche"},
                {texto: "Palo Maria"},
                {texto: "Guayabo Volador"},
                {texto: "Roble"},
                {texto: "Caoba"},
                {texto: "Pino"},
                {texto: "Cedro"},
                {texto: "Árbol de Caucho"},
                {texto: "Álamo"},
                {texto: "Primavera"},
                {texto: "Hormiguillo colorado"},
                {texto: "Palo mulato"},
                {texto: "Bambú"},
                {texto: "Eucalipto"},
                {texto: "Canela"},
                {texto: "Neem"},
                {texto: "Ciprés"},
                {texto: "Paulownia"},
                {texto: "Ceiba"},
                {texto: "Guariuba"},
                {texto: "Figueira"},
                {texto: "Pashaco"},
                {texto: "Chicalá"},
                {texto: "Nogal"},
                {texto: "Abarco"},
                {texto: "Palo rosa"},
                {texto: "Campanos"},
                {texto: "Melina"},
                {texto: "Igua"},
                {texto: "Caracolí"},
                {texto: "Chanul"},
                {texto: "Dinde"}
            ],
            lineaSeleccionada: "",
            actividadSeleccionada: "",
            areaDestinada: "",
            datoslineaSeleccionada: [],
            corregimientos: [],
            municipiosIP: [],
            paises_latinoamericanos: [
                "COLOMBIA",
                "VENEZUELA",
                "MÉXICO",
                "GUATEMALA",
                "BELICE",
                "EL SALVADOR",
                "HONDURAS",
                "NICARAGUA",
                "COSTA RICA",
                "PANAMÁ",
                "CUBA",
                "REPÚBLICA DOMINICANA",
                "PUERTO RICO",
                "HAITÍ",
                "JAMAICA",
                "TRINIDAD Y TOBAGO",
                "BARBADOS",
                "GUYANA",
                "SURINAM",
                "GUYANA FRANCESA",
                "ECUADOR",
                "PERÚ",
                "BOLIVIA",
                "PARAGUAY",
                "BRASIL",
                "CHILE",
                "ARGENTINA",
                "URUGUAY"
            ]
        }
    },
    methods: {
        async consultarValoresDefecto(){
            this.loading = true;
            await this.misDatos();
            await this.verificarLogin();
            await this.consultarJefes();
            await this.consultarDatosIndividuo();
            await this.consultarDepartamentos();
            await this.consultarEscolaridades();
            await this.consultarOcupaciones();
            this.loading = false;
        },
        async consultarDatosIndividuo(){
            await caracterizacionService.consultarDatosIndividuo(this.identificacion_editar).then(async respuesta => {
                var data = respuesta.data;
                this.informacion_personal.rol = data.informacion_personal.rol;
                this.informacion_personal.id_jefe = data.informacion_personal.id_jefe;
                this.informacion_personal.parentesco = data.informacion_personal.parentesco;
                this.informacion_personal.direccion = data.informacion_personal.direccion;
                this.informacion_personal.nombre_completo = data.informacion_personal.nombre_completo;
                this.informacion_personal.fecha_nacimiento = data.informacion_personal.fecha_nacimiento;
                this.informacion_personal.tipo_identificacion = data.informacion_personal.tipo_identificacion;
                this.informacion_personal.numero_identificacion = data.informacion_personal.identificacion;
                this.informacion_personal.direccion_residencia = data.informacion_personal.direccion_residencia;
                this.informacion_personal.sexo = data.informacion_personal.sexo;
                this.informacion_personal.orientacion_sexual = data.informacion_personal.orientacion_sexual;
                this.informacion_personal.estado_civil = data.informacion_personal.estado_civil;
                this.informacion_personal.tiempo_municipo = data.informacion_personal.tiempo_municipio;
                this.informacion_personal.desplazado = data.informacion_personal.desplazado;
                this.informacion_personal.cual_estado_civil = data.informacion_personal.cual_estado_civil;
                this.informacion_personal.cual_sexo = data.informacion_personal.cual_sexo;
                this.informacion_personal.departamento = data.informacion_personal.departamento;
                await this.consultarMunicipiosIP(this.informacion_personal.departamento);
                this.informacion_personal.municipio = data.informacion_personal.municipio;
                await this.listarCorregimientosPorMuni(this.informacion_personal.municipio);
                this.informacion_personal.corregimiento = data.informacion_personal.corregimiento;

                if(this.informacion_personal.rol != "Jefe de hogar"){
                    var item_jefe = this.jefes_hogar.filter(item => item.identificacion == this.informacion_personal.id_jefe)[0];
                    this.seleccionarJefe(item_jefe);
                }

                if(data.origen_identidad != null){
                    this.origen_identidad.identificacion_individuo = data.origen_identidad.identificacion_individuo;
                    this.origen_identidad.pais = data.origen_identidad.pais_nacimiento;
                    this.origen_identidad.departamento = data.origen_identidad.departamento_nacimiento;
                    await this.consultarMunicipios(this.origen_identidad.departamento);
                    this.origen_identidad.municipio = data.origen_identidad.municipio_nacimiento;
                    this.origen_identidad.etnia = data.origen_identidad.etnia;
                }
                
                if(data.educacion != null){
                    this.educacion = data.educacion;
                }

                if(data.situacion_laboral != null){
                    this.situacion_laboral = data.situacion_laboral;
                }

                if(data.salud != null){
                    this.salud = data.salud;
                }
                
                if(data.cultura_tradiciones != null){
                    this.culturaTradiciones = data.cultura_tradiciones;
                }

                if(this.informacion_personal.rol == "Jefe de hogar"){
                    if(data.vivienda_hogar != null){
                        this.viviendaHogar = data.vivienda_hogar;
                        if(this.viviendaHogar.electricidad == "Si"){
                            $("#control_011").prop( "checked", true );
                        }
                        if(this.viviendaHogar.agua_potable == "Si"){
                            $("#control_012").prop( "checked", true );
                        }
                        if(this.viviendaHogar.alcantarillado == "Si"){
                            $("#control_013").prop( "checked", true );
                        }
                        if(this.viviendaHogar.gas_natural == "Si"){
                            $("#control_014").prop( "checked", true );
                        }
                        if(this.viviendaHogar.aseo == "Si"){
                            $("#control_015").prop( "checked", true );
                        }
                        if(this.viviendaHogar.otro == "Si"){
                            $("#control_016").prop( "checked", true );
                        }
                    }
                }              
            });
        },
        clickTabla(){
            var table = document.getElementById('miTabla');
            var rows = table.getElementsByTagName('tbody')[0].getElementsByTagName('tr');
            for (var i = 0; i < rows.length; i++) {
                rows[i].onclick = function() {
                    var isSelected = this.style.backgroundColor === 'rgb(255, 255, 153)';
                    
                    for (var j = 0; j < rows.length; j++) {
                        rows[j].style.backgroundColor = '';
                    }

                    if (!isSelected) {
                        this.style.backgroundColor = '#ffff99';
                    }
                };
            }

            setTimeout(()=>{
                this.dataTables();
            }, 200)
        },
        dataTables() {
            $("#miTabla").DataTable({
                ordering: false,
                language: {
                    "decimal": "",
                    "emptyTable": "No hay información",
                    "info": "Mostrando _START_ a _END_ de _TOTAL_ Entradas",
                    "infoEmpty": "Mostrando 0 to 0 of 0 Entradas",
                    "infoFiltered": "(Filtrado de _MAX_ total entradas)",
                    "infoPostFix": "",
                    "thousands": ",",
                    "lengthMenu": "Mostrar _MENU_ Entradas",
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
        async verificarLogin(){
            var navigate = this.$router;
            await usuarioService.verificarLogin().then(respuesta => {
                if(respuesta.data == 0){
                    navigate.push({ name: 'loginPage'})
                }
            });
        },
        misDatos: async function () {
            try {
                await usuarioService.misDatos().then(respuesta => {
                    this.datos = respuesta.data;
                });
            } catch (error) {
                console.log(error);
            }
        },
        async consultarJefes(){
            try {
                await caracterizacionService.jefesHogar().then(respuesta => {
                    this.jefes_hogar = respuesta.data;
                    setTimeout(()=>{
                        this.clickTabla();
                    }, 200)
                });
            } catch (error) {
                console.log(error);
            }
        },
        async seleccionarJefe(item){
            this.informacion_personal.id_jefe = item.identificacion;
            this.informacion_personal.nombre_jefe = item.tipo_identificacion+" - "+item.identificacion+" - "+item.nombre_completo;
            this.informacion_personal.direccion = item.direccion;
            this.informacion_personal.direccion_residencia = item.direccion_residencia;
            this.informacion_personal.departamento = item.departamento;
            await this.consultarMunicipiosIP(this.informacion_personal.departamento);
            this.informacion_personal.municipio = item.municipio;
            await this.listarCorregimientosPorMuni(this.informacion_personal.municipio);
            this.informacion_personal.corregimiento = item.corregimiento;
        },
        eliminarActive(id, id2){
            document.getElementById("baseIcon-tab21").classList.remove("active");
            document.getElementById("baseIcon-tab22").classList.remove("active");
            document.getElementById("baseIcon-tab23").classList.remove("active");
            document.getElementById("baseIcon-tab24").classList.remove("active");
            document.getElementById("baseIcon-tab25").classList.remove("active");
            document.getElementById("baseIcon-tab26").classList.remove("active");
            document.getElementById("baseIcon-tab27").classList.remove("active");


            document.getElementById("tabIcon21").classList.remove("active");
            document.getElementById("tabIcon22").classList.remove("active");
            document.getElementById("tabIcon23").classList.remove("active");
            document.getElementById("tabIcon24").classList.remove("active");
            document.getElementById("tabIcon25").classList.remove("active");
            document.getElementById("tabIcon26").classList.remove("active");
            document.getElementById("tabIcon27").classList.remove("active");
            document.getElementById("tabIcon28").classList.remove("active");

            if(id != "baseIcon-tab28"){
                document.getElementById(id).classList.add("active");
            }

            document.getElementById(id2).classList.add("active");

        },
        validarCamposInformacionPersonal() {
            this.errores = "";

            if (!this.informacion_personal.rol) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  rol </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.parentesco) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  parentesco </strong> no puede estar vacío <br>`;
            }
            
            if (!this.informacion_personal.direccion) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  dirección </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.nombre_completo) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  nombre </strong> no puede estar vacío <br>`;
            }

            if (this.informacion_personal.fecha_nacimiento == "") {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  fecha de nacimiento </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.tipo_identificacion) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  tipo de identificación </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.numero_identificacion) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  número de identificación </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.direccion_residencia) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  dirección de residencia </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.sexo) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  sexo </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.orientacion_sexual) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  orientación sexual </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.estado_civil) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  estado civil </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.tiempo_municipo) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  tiempo de residencia </strong> no puede estar vacío <br>`;
            }

            if (!this.informacion_personal.desplazado) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  desplazado </strong> no puede estar vacío <br>`;
            }


            if (this.informacion_personal.rol != "Jefe de hogar" && this.informacion_personal.id_jefe == "") {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  jefe de hogar </strong> no puede estar vacío <br>`;
            }



            if (this.errores.length === 0) {
                this.guardarInformacionPersonal();
            } else {
                Swal.fire({
                    width: 600,
                    icon: "error",
                    title: "Información Importante",
                    html: this.errores,
                });
            }
        },
        async guardarInformacionPersonal(){
            this.loading = true;
            try {
                await caracterizacionService.guardarInformacionTemporal(this.informacion_personal).then(respuesta => {
                    this.loading = false;
                    toastr.success("Datos guardados correctamente");
                    this.eliminarActive("baseIcon-tab22", "tabIcon22");
                });
            } catch (error) {
                console.log(error);
            }
        },
        async consultarDepartamentos(){
            await caracterizacionService.departamentos().then(respuesta => {
                this.departamentos = respuesta.data;
            });
        },
        async consultarMunicipios(){
            this.loading = true;
            this.origen_identidad.municipio = "";
            await caracterizacionService.municipios(this.origen_identidad.departamento).then(respuesta => {
                this.municipios = respuesta.data;
                this.loading = false;
            });
        },
        async consultarMunicipiosIP(){
            this.loading = true;
            this.informacion_personal.municipio = "";
            await caracterizacionService.municipios(this.informacion_personal.departamento).then(respuesta => {
                this.municipiosIP = respuesta.data;
                this.loading = false;
            });
        },
        async listarCorregimientosPorMuni(){
            this.loading = true;
            await corregimientosService.listarCorregimientosPorMuni( this.informacion_personal.municipio).then(respuesta => {
                this.corregimientos = respuesta.data;
                this.loading = false;
            });
        },
        validarCamposOrigenIdentidad(){
            this.errores = "";

            if (this.origen_identidad.pais == "COLOMBIA") {
                if (!this.origen_identidad.departamento) {
                    this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  departamento </strong> no puede estar vacío <br>`;
                }

                if (!this.origen_identidad.municipio) {
                    this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  municipio </strong> no puede estar vacío <br>`;
                }
            }

            if (!this.origen_identidad.etnia) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  etnia </strong> no puede estar vacío <br>`;
            }

            if (this.errores.length === 0) {
                this.origen_identidad.identificacion_individuo = this.informacion_personal.numero_identificacion;
                this.guardarOrigenIdentidad();
            } else {
                Swal.fire({
                    width: 600,
                    icon: "error",
                    title: "Información Importante",
                    html: this.errores,
                });
            }
        },
        async guardarOrigenIdentidad(){
            this.loading = true;
            try {
                await caracterizacionService.guardarOrigenIdentidad(this.origen_identidad).then(respuesta => {
                    this.loading = false;
                    toastr.success("Datos guardados correctamente");
                    this.eliminarActive("baseIcon-tab23", "tabIcon23");
                });
            } catch (error) {
                console.log(error);
            }
        },
        async consultarEscolaridades(){
            await caracterizacionService.escolaridad().then(respuesta => {
                this.escolaridades = respuesta.data;
            });
        },
        validarCamposOrigenEducacion(){
            this.errores = "";

            if (!this.educacion.apoyo_educativo) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  apoyo educativo </strong> no puede estar vacío <br>`;
            }

            if (!this.educacion.barreras_educacion) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   barreras para acceder a la educación </strong> no puede estar vacío <br>`;
            }

            if (!this.educacion.descriminacion_racial) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  discriminación racial  </strong> no puede estar vacío <br>`;
            }

            if (!this.educacion.educacion_especifica) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  educación específica sobre la cultura  </strong> no puede estar vacío <br>`;
            }

            if (!this.educacion.nivel_educativo) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  nivel educativo </strong> no puede estar vacío <br>`;
            }

            if (this.errores.length === 0) {
                this.educacion.identificacion_individuo = this.informacion_personal.numero_identificacion;
                this.guardarEducacion();
            } else {
                Swal.fire({
                    width: 600,
                    icon: "error",
                    title: "Información Importante",
                    html: this.errores,
                });
            }

        },
        async guardarEducacion(){
            this.loading = true;
            try {
                await caracterizacionService.guardarEducacion(this.educacion).then(respuesta => {
                    this.loading = false;
                    toastr.success("Datos guardados correctamente");
                    this.eliminarActive("baseIcon-tab24", "tabIcon24");
                });
            } catch (error) {
                console.log(error);
            }
        },
        async consultarOcupaciones(){
            await caracterizacionService.ocupaciones().then(respuesta => {
                this.ocupaciones = respuesta.data;
            });
        },
        validarCamposSituacionLaboral(){
            this.errores = "";

            if (!this.situacion_laboral.discriminacion_laboral) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  discriminación  laboral </strong> no puede estar vacío <br>`;
            }

            if (!this.situacion_laboral.ocupacion) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   Ocupación </strong> no puede estar vacío <br>`;
            }

            if (!this.situacion_laboral.ingreso_mensual) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  ingreso mensual </strong> no puede estar vacío <br>`;
            }

            if (!this.situacion_laboral.situacion_laboral) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  situación laboral  </strong> no puede estar vacío <br>`;
            }

            if (this.errores.length === 0) {
                this.situacion_laboral.identificacion_individuo = this.informacion_personal.numero_identificacion;
                this.guardarSituacionLaboral();
            } else {
                Swal.fire({
                    width: 600,
                    icon: "error",
                    title: "Información Importante",
                    html: this.errores,
                });
            }
        },
        async guardarSituacionLaboral(){
            this.loading = true;
            try {
                await caracterizacionService.guardarSituacionLaboral(this.situacion_laboral).then(respuesta => {
                    this.loading = false;
                    toastr.success("Datos guardados correctamente");
                    this.eliminarActive("baseIcon-tab25", "tabIcon25");
                });
            } catch (error) {
                console.log(error);
            }
        },
        validarCamposSalud(){
            this.errores = "";

            if (!this.salud.acceso_salud) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  Acceso a servicios de salud  </strong> no puede estar vacío <br>`;
            }

            if (!this.salud.condicion_discapacidad) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  condición de discapacidad  </strong> no puede estar vacío <br>`;
            }

            if (!this.salud.discriminacion_salud) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  discriminación en el ámbito de la salud  </strong> no puede estar vacío <br>`;
            }

            if (!this.salud.estado_salud) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  Estado de salud general  </strong> no puede estar vacío <br>`;
            }

            if (!this.salud.regimen) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  Régimen de salud al que pertenece  </strong> no puede estar vacío <br>`;
            }

            if (this.errores.length === 0) {
                this.salud.identificacion_individuo = this.informacion_personal.numero_identificacion;
                this.guardarSalud();
            } else {
                Swal.fire({
                    width: 650,
                    icon: "error",
                    title: "Información Importante",
                    html: this.errores,
                });
            }
        },
        async guardarSalud(){
            this.loading = true;
            try {
                await caracterizacionService.guardarSalud(this.salud).then(respuesta => {
                    this.loading = false;
                    toastr.success("Datos guardados correctamente");
                    this.eliminarActive("baseIcon-tab26", "tabIcon26");
                });
            } catch (error) {
                console.log(error);
            }
        },
        validarCamposCulturaTradiciones(){
            this.errores = "";

            if (!this.culturaTradiciones.habla_lengua) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  habla alguna lengua afrocolombiana  </strong> no puede estar vacío <br>`;
            }

            if (!this.culturaTradiciones.practica_actividades) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">  Participa en actividades culturales afrocolombianas  </strong> no puede estar vacío <br>`;
            }

            if (!this.culturaTradiciones.practicas_religiosas) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   Prácticas culturales y religiosas  </strong> no puede estar vacío <br>`;
            }

            if(this.culturaTradiciones.habla_lengua == "Si"){
                if (!this.culturaTradiciones.cual_lengua) {
                    this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   ¿Cual lengua habla?  </strong> no puede estar vacío <br>`;
                }
            }
            
            if (this.errores.length === 0) {
                this.culturaTradiciones.identificacion_individuo = this.informacion_personal.numero_identificacion;
                this.guardarCulturaTradiciones();
            } else {
                Swal.fire({
                    width: 710,
                    icon: "error",
                    title: "Información Importante",
                    html: this.errores,
                });
            }
        },
        async guardarCulturaTradiciones(){
            this.loading = true;
            try {
                await caracterizacionService.guardarCulturaTradiciones(this.culturaTradiciones).then(respuesta => {
                    this.loading = false;
                    if(this.informacion_personal.rol == "Jefe de hogar"){
                        toastr.success("Datos guardados correctamente");
                        this.eliminarActive("baseIcon-tab27", "tabIcon27");
                    }else{
                        toastr.success("Datos guardados correctamente");
                        this.eliminarActive("baseIcon-tab28", "tabIcon28");
                    }
                });
            } catch (error) {
                console.log(error);
            }
        },
        validarCamposViviendaHogar(){
            this.errores = "";

            if (!this.viviendaHogar.tipo_vivienda) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2"> Tipo de vivienda </strong> no puede estar vacío <br>`;
            }

            if (!this.viviendaHogar.tenencia) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2"> Tenencia de la vivienda </strong> no puede estar vacío <br>`;
            }

            if (!this.viviendaHogar.numero_personas_hogar) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2"> Número de personas que conforman el hogar </strong> no puede estar vacío <br>`;
            }

            if (!this.viviendaHogar.numero_personas_trabajan) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2"> Personas trabajan en su hogar </strong> no puede estar vacío <br>`;
            }

            if (!this.viviendaHogar.ingresos_mensuales_hogar) {
                this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2"> Ingresos mensuales por hogar </strong> no puede estar vacío <br>`;
            }

            if(this.viviendaHogar.tipo_vivienda == "Otro"){
                if (!this.viviendaHogar.cual_tipo_vivienda) {
                    this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   ¿Cual tipo de vivienda?  </strong> no puede estar vacío <br>`;
                }
            }
            
            if(this.viviendaHogar.tipo_vivienda == 'Finca' || this.viviendaHogar.tipo_vivienda == 'Parcela'){
                if (!this.viviendaHogar.posecion_baldia) {
                    this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   Posesión baldía  </strong> no puede estar vacío <br>`;
                }

                if (!this.viviendaHogar.propiedad_titulo) {
                    this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   Propiedad con título  </strong> no puede estar vacío <br>`;
                }

                if (!this.viviendaHogar.area_total) {
                    this.errores += `<i style ="font-size: 8px" class="fas fa-circle"></i> El campo <strong style="color: #2f95a2">   Área Total  </strong> no puede estar vacío <br>`;
                }
            }

            if (this.errores.length === 0) {
                this.viviendaHogar.identificacion_jefe = this.informacion_personal.numero_identificacion;
                this.guardarViviendaHogar();
            } else {
                Swal.fire({
                    width: 710,
                    icon: "error",
                    title: "Información Importante",
                    html: this.errores,
                });
            }
        },
        async guardarViviendaHogar(){
            this.loading = true;
            try {
                await caracterizacionService.guardarViviendaHogar(this.viviendaHogar).then(respuesta => {
                    this.loading = false;
                    toastr.success("Datos guardados correctamente");
                    this.eliminarActive("baseIcon-tab28", "tabIcon28");
                });
            } catch (error) {
                console.log(error);
            }
        },
        seleccionarLinea(){
            switch (this.lineaSeleccionada) {
                case 1:
                    this.datoslineaSeleccionada = this.linea1
                    break;
                case 2:
                    this.datoslineaSeleccionada = this.linea2
                    break;
                case 3:
                    this.datoslineaSeleccionada = this.linea3
                    break;
            }
        },
        agregarActividad(){
            var suma_tem = 0;

            this.viviendaHogar.actividades_economicas.forEach(element => {
                suma_tem += element.area_destinada;
            });

            suma_tem += parseFloat(this.areaDestinada);

            if(suma_tem <= this.viviendaHogar.area_total){
                var object = {
                    linea: this.lineas_economicas[this.lineaSeleccionada-1].texto,
                    actividad: this.actividadSeleccionada,
                    area_destinada: parseFloat(this.areaDestinada)
                }

                this.viviendaHogar.actividades_economicas.push(object);

                this.lineaSeleccionada = "";
                this.actividadSeleccionada = "";
                this.areaDestinada = "";
            }else{
                toastr.error("El total del area destinada de las actividades no puede ser mayor a "+this.viviendaHogar.area_total+" (metros cuadrados)");
            }           
        },
        eliminarActividad(index){
            Swal.fire({
                title: "¿Desea eliminar esta actividad?",
                showCancelButton: true,
                confirmButtonText: "Si",
                cancelButtonText: "No",
                cancelButtonColor: "#D70040"
            }).then((result) => {
                if (result.isConfirmed) {
                    this.viviendaHogar.actividades_economicas.splice(index, 1);
                } 
            });
           
        }
    },
}
</script>
<style scoped>
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
        font-size: 18px !important;
        cursor: pointer;
    }
</style>