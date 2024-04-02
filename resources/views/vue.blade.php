<!DOCTYPE html>


@include('plantilla.head')

<html lang="en">

<body style="background-image: url('/imagenes/fondo.png'); background-size: 100% 100%">
    <div id="loginDiv" style="height: 100vh">
        <section class="row flexbox-container" style="height: 100vh; align-items: center">
            <div style="height: 93vh;" class="col-6 d-flex align-items-center justify-content-center">
                <div class="col-lg-10 box-shadow-2 p-0">
                    <div class="card border-grey border-lighten-3 px-1 py-1 m-0">
                        <div class="card-header border-0">
                            <div class="card-title text-center">
                                <img style="height: 200px" src="imagenes/elPasoEscudo.png" alt="branding logo">
                            </div>
                            
                        </div>
                        <div class="card-content">
                            <p class="card-subtitle line-on-side text-muted text-center font-small-3 mx-2 my-1"><span>Ingrese los siguientes datos</span></p>
                            <div class="card-body">
                                <form class="form-horizontal">
                                    <fieldset class="form-group position-relative has-icon-left">
                                        <input id="username" type="text" class="form-control" id="user-name" placeholder="usuario / correo" required>
                                        <div class="form-control-position">
                                            <i class="feather icon-user"></i>
                                        </div>
                                    </fieldset>
                                    <fieldset class="form-group position-relative has-icon-left">
                                        <input id="password" type="password" class="form-control" id="user-password" placeholder="Contraseña" required>
                                        <div class="form-control-position">
                                            <i class="fa fa-key"></i>
                                        </div>
                                    </fieldset>
                                    <div class="form-group row">
                                        <div class="col-sm-6 col-12 text-center text-sm-left pr-0">
                                            <fieldset>
                                                <input type="checkbox" id="remember-me" class="chk-remember">
                                                <label for="remember-me"> Recordar la contraseña</label>
                                            </fieldset>
                                        </div>
                                        <div class="col-sm-6 col-12 float-sm-left text-center text-sm-right"><a href="recover-password.html" class="card-link">¿Olvidaste tu contraseña?</a></div>
                                    </div>
                                    <a onclick="login()" type="submit" class="btn btn-outline-primary btn-block"><i class="feather icon-unlock"></i> Iniciar sesión</a>
                                </form>
                            </div>
                            
                        </div>
                    </div>
                </div>
            </div>
            <div class="col-6"></div>
        </section>
    </div>

    <script src="{{ asset('js/jquery.3.6.min.js') }}"></script>
    <script src="{{ asset('app-assets/vendors/js/extensions/toastr.min.js') }}"></script>
    <script src="{{ asset('app-assets/js/scripts/extensions/toastr.js') }}"></script>

    <script>
         toastr.options = {
            "closeButton": true,
            "newestOnTop": false,
            "progressBar": false,
            "positionClass": "toast-top-right",
            "preventDuplicates": false,
            "onclick": null,
            "showDuration": "500",
            "hideDuration": "1000",
            "timeOut": "5000",
            "extendedTimeOut": "1000",
            "showEasing": "swing",
            "hideEasing": "linear",
            "showMethod": "fadeIn",
            "hideMethod": "fadeOut"
        }

        document.addEventListener("keydown", function(event) {
            if (event.key === "Enter") {
                login();
            }
        });
        
        function login(){
            var texto = document.getElementById("username").value;
            var password = document.getElementById("password").value;

            if(texto == "" || password == ""){
                toastr.error("Todos los campos son obligatorios");
            }else{
                $.ajax({
                    url: '/api/iniciar-sesion?usuario='+texto+"&password="+password,
                    type: 'GET',
                    success: function(response) {
                        if (response[1] == 1) {
                            toastr.success(response[0]);
                            setTimeout(() => {
                                location.href = "/lista-caracterizados";
                            }, 2000);
                        } else {
                            toastr.error(response[0]);
                        }
                    },
                    error: function(error) {
                        toastr.warning(error);
                    }
                });
            }
        }
    </script>
</body>
</html>