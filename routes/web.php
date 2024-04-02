<?php

use Illuminate\Support\Facades\Route;
use App\Http\Controllers\UsuarioController;
use App\Http\Controllers\CaracterizacionController;
use App\Http\Controllers\CorregimeintoController;
use App\Http\Controllers\ConsolidadoController;
use App\Http\Controllers\DashboardController;
use App\Http\Controllers\ExportarExcelController;

use Illuminate\Http\Request;


/*
|--------------------------------------------------------------------------
| Web Routes
|--------------------------------------------------------------------------
|
| Here is where you can register web routes for your application. These
| routes are loaded by the RouteServiceProvider within a group which
| contains the "web" middleware group. Now create something great!
|
*/

// routes/web.php
// Rutas de la API
Route::prefix('api')->group(function () {
    Route::get('/iniciar-sesion', [UsuarioController::class, 'loginUsuario'])->name('loginUsuario');
    Route::get('/mis-datos', [UsuarioController::class, 'misDatos'])->name('misDatos');
    Route::get('/verificarLogin', [UsuarioController::class, 'verificarLogin'])->name('verificarLogin');
    Route::post('/editar-usuario', [UsuarioController::class, 'editarUsuario'])->name('editarUsuario');
    Route::post('/cambiar-password', [UsuarioController::class, 'cambiarPassword'])->name('cambiarPassword');
    Route::get('/consultar-usuarios', [UsuarioController::class, 'consultarUsuarios'])->name('consultarUsuarios');
    Route::post('/registrar-usuario', [UsuarioController::class, 'registrarUsuario'])->name('registrarUsuario');
    Route::post('/editar-usuario-admin', [UsuarioController::class, 'editarUsuarioAdmin'])->name('editarUsuarioAdmin');
    Route::get('/cambiar-estado-usuario', [UsuarioController::class, 'cambiarEstadoUsuario'])->name('cambiarEstadoUsuario');
    Route::get('/cerrar-sesion', [UsuarioController::class, 'cerrarSesion'])->name('cerrarSesion');
    Route::get('/listar-usuarios', [UsuarioController::class, 'listarUsuarios'])->name('listarUsuarios');

    Route::get('/caracterizados', [CaracterizacionController::class, 'listarCaracterizados'])->name('listarCaracterizados');
    Route::get('/caracterizados-digitador', [CaracterizacionController::class, 'consultarCaracterizadosDigitador'])->name('consultarCaracterizadosDigitador');
    Route::get('/datos-individuo', [CaracterizacionController::class, 'consultarDatosIndividuo'])->name('consultarDatosIndividuo');

    Route::post('/guardar-informacion-personal', [CaracterizacionController::class, 'guardarInformacionPersonal'])->name('guardarInformacionPersonal');
    Route::post('/guardar-origen-entidad', [CaracterizacionController::class, 'guardarOrigenEntidad'])->name('guardarOrigenEntidad');
    Route::post('/guardar-educacion', [CaracterizacionController::class, 'guardarEducacion'])->name('guardarEducacion');
    Route::post('/guardar-situacion-laboral', [CaracterizacionController::class, 'guardarSituacionLaboral'])->name('guardarSituacionLaboral');
    Route::post('/guardar-salud', [CaracterizacionController::class, 'guardarSalud'])->name('guardarSalud');
    Route::post('/guardar-cultura-tradiciones', [CaracterizacionController::class, 'guardarCulturaTradiciones'])->name('guardarCulturaTradiciones');
    Route::post('/guardar-vivienda-hogar', [CaracterizacionController::class, 'guardarViviendaHogar'])->name('guardarViviendaHogar');
    Route::get('/eliminar-caracterizacion', [CaracterizacionController::class, 'eliminarCaracterizacion'])->name('eliminarCaracterizacion');

    Route::get('/jefes-hogar', [CaracterizacionController::class, 'consultarJefesHogar'])->name('consultarJefesHogar');
    Route::get('/departamentos', [CaracterizacionController::class, 'consultarDepartamentos'])->name('consultarDepartamentos');
    Route::get('/municipios', [CaracterizacionController::class, 'consultarMunicipios'])->name('consultarMunicipios');
    Route::get('/escolaridad', [CaracterizacionController::class, 'consultarEscolaridad'])->name('consultarEscolaridad');
    Route::get('/ocupaciones', [CaracterizacionController::class, 'consultarOcupaciones'])->name('consultarOcupaciones');
    
    Route::get('/consolidado', [ConsolidadoController::class, 'consolidadoCensados'])->name('consolidadoCensados');
    Route::get('/municipios-consolidado', [ConsolidadoController::class, 'municipiosConsolidado'])->name('municipiosConsolidado');
    Route::get('/paginacion-consolidado', [ConsolidadoController::class, 'paginacionConsolidado'])->name('paginacionConsolidado');
    Route::get('/exportar-consolidado', [ConsolidadoController::class, 'exportarConsolidadoExcel'])->name('exportarConsolidadoExcel');
    Route::get('/consolidado-concejo', [ConsolidadoController::class, 'consolidadoCensadosConcejo'])->name('consolidadoCensadosConcejo');
    Route::get('/paginacion-consolidado-concejo', [ConsolidadoController::class, 'paginacionConsolidadoConcejo'])->name('paginacionConsolidadoConcejo');
    Route::get('/exportar-consolidado-concejo', [ConsolidadoController::class, 'exportarConsolidadoExcelConcejo'])->name('exportarConsolidadoExcelConcejo');

    Route::get('/lista-corregimientos', [CorregimeintoController::class, 'listarCorregimientos'])->name('listarCorregimientos');
    Route::post('/guardar-corregimiento', [CorregimeintoController::class, 'guardarCorregimiento'])->name('guardarCorregimiento');
    Route::get('/lista-corregimientos-id', [CorregimeintoController::class, 'listarCorregimientosMunicipio'])->name('listarCorregimientosMunicipio');
    Route::get('/eliminar-corregimiento', [CorregimeintoController::class, 'elimianrCorregimiento'])->name('elimianrCorregimiento');

    Route::get('/datos-dashboard', [DashboardController::class, 'datosDashboard'])->name('datosDashboard');
    Route::get('/datos-informe', [DashboardController::class, 'datosInforme'])->name('datosInforme');


    Route::get('/exportar-excel', [ExportarExcelController::class, 'exportarDatosExcel'])->name('exportarDatosExcel');

});

Route::get('/', function () {
    return view('vue');
});

Route::get('/{any}', function () {
    return view('principal');
})->where('any', '.*');