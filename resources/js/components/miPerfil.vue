<template lang="">
    <div>
        <div class="app-content" style = "background-color: #f5f7fa;">
            <div class="content-overlay"></div>
            <div class="content-wrapper">
                <div class="content-header row">
                    <div class="content-header-left col-md-6 col-12 mb-2">
                        <h3 class="content-header-title mb-0">CONFIGURACIÓN DE CUENTA</h3>
                    </div>
                </div>
                <div class="content-body">
                    <!-- account setting page start -->
                    <section id="page-account-settings">
                        <div class="row">
                            <!-- left menu section -->
                            <div class="col-md-3 mb-2 mb-md-0">
                                <ul class="nav nav-pills flex-column mt-md-0 mt-1">
                                    <li class="nav-item">
                                        <a class="nav-link d-flex active" id="account-pill-general" data-toggle="pill" href="#account-vertical-general" aria-expanded="true">
                                            <i class="feather icon-globe"></i>
                                            General
                                        </a>
                                    </li>
                                    <li class="nav-item">
                                        <a class="nav-link d-flex" id="account-pill-password" data-toggle="pill" href="#account-vertical-password" aria-expanded="false">
                                            <i class="feather icon-lock"></i>
                                            Cambiar la contraseña
                                        </a>
                                    </li>
                                </ul>
                            </div>
                            <div class="col-md-9">
                                <div class="card">
                                    <div class="card-content">
                                        <div class="card-body" style="min-height: 300pt">
                                            <Skeleton v-if="loading == true"></Skeleton>
                                            <div class="tab-content" v-if="loading == false">
                                                <div role="tabpanel" class="tab-pane active" id="account-vertical-general" aria-labelledby="account-pill-general" aria-expanded="true">
                                                    <div class="media">
                                                        <a href="javascript: void(0);">
                                                            <img id="imagen_avatar2" :src="datos.imagen" class="rounded mr-75" alt="profile image" height="64" width="64">
                                                        </a>
                                                        <div class="media-body mt-75">
                                                            <div class="col-12 px-0 d-flex flex-sm-row flex-column justify-content-start">
                                                                <label class="btn btn-sm btn-primary ml-50 mb-50 mb-sm-0 cursor-pointer" for="foto_perfil">Cambiar imagen de perfil</label>
                                                                <input @change="handleFileChange" id="foto_perfil" type="file">
                                                            </div>
                                                            <p class="text-muted ml-75 mt-50"><small>Permitido JPG, GIF o PNG. máx.
                                                                    tamaño de
                                                                    800kB</small></p>
                                                        </div>
                                                    </div>
                                                    <hr>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datos.identificacion" name="nombre" placeholder="Ingrese su identificación" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-id-card"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datos.nombre" name="nombre" placeholder="Ingrese su nombre" required>
                                                                <div class="form-control-position">
                                                                    <i class="feather icon-user"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datos.rol" disabled name="nombre" placeholder="Ingrese su nombre" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-user-lock"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datos.usuario" name="nombre" placeholder="Ingrese su usuario" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-at"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datos.email" name="nombre" placeholder="Ingrese su email" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-envelope"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                        <div class="col-6">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datos.celular" name="nombre" placeholder="Ingrese su número celular" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-phone"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>
                                                    </div>
                                                    <div class="row">
                                                        <div class="col-12">
                                                            <fieldset class="form-group position-relative has-icon-left">
                                                                <input type="text" class="form-control" v-model="datos.direccion" name="nombre" placeholder="Ingrese su dirección" required>
                                                                <div class="form-control-position">
                                                                    <i class="fas fa-route"></i>
                                                                </div>
                                                            </fieldset>
                                                        </div>    
                                                    </div>
                                                    <hr>
                                                    <div class="col-12 d-flex flex-sm-row flex-column justify-content-end" style="padding: 0">
                                                        <button @click="editarUsuario" type="button" class="btn btn-primary mr-sm-1 mb-1 mb-sm-0"><i class="fas fa-save"></i> Guardar Cambios</button>
                                                        <button type="reset" class="btn btn-danger"><i class="fas fa-times"></i> Cancelar</button>
                                                    </div>
                                                </div>
                                                <div class="tab-pane fade " id="account-vertical-password" role="tabpanel" aria-labelledby="account-pill-password" aria-expanded="false">
                                                    <form novalidate>
                                                        <div class="row">
                                                            <div class="col-12">
                                                                <div class="form-group">
                                                                    <div class="controls">
                                                                        <label for="account-old-password">Contraseña anterior</label>
                                                                        <input v-model="passwordOld" type="password" class="form-control" id="account-old-password" required placeholder="Contraseña anterior" data-validation-required-message="This old password field is required">
                                                                    </div>
                                                                </div>
                                                            </div>
                                                            <div class="col-12">
                                                                <div class="form-group">
                                                                    <div class="controls">
                                                                        <label for="account-new-password">Nueva contraseña</label>
                                                                        <input v-model="passwordNew" type="password" name="password" id="account-new-password" class="form-control" placeholder="Nueva contraseña" required data-validation-required-message="The password field is required" minlength="6">
                                                                    </div>
                                                                </div>
                                                            </div>
                                                            <div class="col-12">
                                                                <div class="form-group">
                                                                    <div class="controls">
                                                                        <label for="account-retype-new-password">Reescriba la nueva contraseña</label>
                                                                        <input v-model="passwordNew2" type="password" name="con-password" class="form-control" required id="account-retype-new-password" data-validation-match-match="password" placeholder="Reescriba la nueva contraseña" data-validation-required-message="The Confirm password field is required" minlength="6">
                                                                    </div>
                                                                </div>
                                                            </div>
                                                            <div class="col-12 d-flex flex-sm-row flex-column justify-content-end">
                                                                <button @click="cambiarPassword" type="button" class="btn btn-primary mr-sm-1 mb-1 mb-sm-0"><i class="fas fa-save"></i> Guardar Cambios</button>
                                                                <button type="reset" class="btn btn-danger"><i class="fas fa-times"></i> Cancelar</button>
                                                            </div>
                                                        </div>
                                                    </form>
                                                </div>
                                            
                                            </div>
                                        </div>
                                    </div>
                                </div>
                                
                            </div>
                        </div>
                    </section>
                    <!-- account setting page end -->
                </div>
            </div>
        </div>
    </div>
</template>
<script>
import * as usuarioService from "../services/usuario_service.js";
import Skeleton from './skeleton/skeletonSmall.vue';

export default {
    components: {
        Skeleton
    },
    data() {
        return {
            datos: {
                tio_registro: null,
            },
            loading: true,
            passwordOld: "",
            passwordNew: "",
            passwordNew2: "",
        };
    },
    mounted() {
        this.misDatos();
        this.verificarLogin();
        document.title = 'Mi Perfil';
    },
    methods: {
        async verificarLogin(){
            var navigate = this.$router;
            await usuarioService.verificarLogin().then(respuesta => {
                if(respuesta.data == 0){
                    navigate.push({ name: 'loginPage'})
                }
            });
        },
        misDatos: async function () {
            this.loading = true;
            try {
                await usuarioService.misDatos().then(respuesta => {
                    this.datos = respuesta.data;
                    this.datos.imagen_anterior = this.datos.imagen;
                    this.datos.imagen = '/imagenes_usuarios/'+this.datos.imagen;
                    this.imagen_cambiada = 0;
                    this.loading = false;
                });
            } catch (error) {
                console.log(error);
            }
        },
        handleFileChange(event) {
            const selectedFile = event.target.files[0];
            this.datos.imagen = URL.createObjectURL(selectedFile);
            this.datos.imagen_nueva = event.target.files[0];
            this.datos.imagen_cambiada = 1;
        },
        async editarUsuario(){
            this.loading = true;
            try {
                await usuarioService.editar_usuario(this.datos).then(respuesta => {
                    toastr.success(respuesta.data[0]);
                    this.misDatos();
                });
            } catch (error) {
                console.log(error);
            }
        },
        async cambiarPassword(){
            this.loading = true;
            const datos = {
                password_old: this.passwordOld,
                password_new: this.passwordNew,
            };

            if(this.passwordNew != ""  && this.passwordNew2 != ""){
                if(this.passwordNew == this.passwordNew2){
                    await usuarioService.cambiar_password(datos).then(respuesta => {
                        var respuesta_ok = respuesta.data;
                        if(respuesta_ok[1] == 1){
                            toastr.success(respuesta_ok[0]);
                            this.passwordOld = "";
                            this.passwordNew = "";
                            this.passwordNew2 = "";
                        }else{
                            toastr.error(respuesta_ok[0]);
                        }

                        this.loading = false;
                    });
                }else{
                    toastr.error("Las contraseñas no coinciden.");
                }
            }else{
                toastr.error("Todos los campos son obligatorios.");
            }
        }
    }
}
</script>
<style scoped>
    #foto_perfil {
        display: none;
    }
</style>