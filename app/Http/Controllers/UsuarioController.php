<?php

namespace App\Http\Controllers;
use Illuminate\Http\Request;
use DB;
use Illuminate\Support\Facades\File;

use Illuminate\Support\Facades\Session;

class UsuarioController extends Controller
{

    public function loginUsuario(Request $request){
        $username = $request->input('usuario');
        $password = $request->input('password');

        $usuario = DB::connection("mysql")->table("users")
        ->where("usuario", $username)
        ->orWhere("email", $username)
        ->first();

        if (!$usuario) {
            return response()->json(["No encontramos ningún usuario registrado con el usuario/correo ingresado", 0], 200);
        } else {
            if($usuario->password == md5($password)){
                Session::flush();
                Session::start();
                Session::put('id', $usuario->id);
                Session::put('nombre', $usuario->nombre);
                Session::put('email', $usuario->email);
                Session::put('imagen', $usuario->imagen);
                Session::put('rol', $usuario->rol);
                Session::put('logueado', true);
                
                return response()->json(["Bienvenido ".$usuario->nombre, 1], 200);
            }else{
                return response()->json(["Contraseña incorrecta", 0], 200);
            }
        }
    }

    public function misDatos(){
    
        $idUsuario = Session::get('id');

        $usuario = DB::connection("mysql")->table("users")
        ->where("id", $idUsuario)
        ->first();
        
        return response()->json($usuario, 200);
        
    }


    public function verificarLogin(){
        if(Session::has('logueado')){
            return response()->json(1, 200);
        }else{
            return response()->json(0, 200);
        }
    }

    public function editarUsuario(Request $request){

        $datos = [
            'id' => $request->input('id'),
            'identificacion' => $request->input('identificacion'),
            'nombre' => $request->input('nombre'),
            'usuario' => $request->input('usuario'),
            'email' => $request->input('email'),
            'celular' => $request->input('celular'),
            'direccion' => $request->input('direccion'),
        ];

        $imagen_cambiada = $request->input('imagen_cambiada');
       
        if($imagen_cambiada == 1){
            $imagen_nueva = $request->file('imagen_nueva');
            $nombreArchivo = time() . '_' . $imagen_nueva->getClientOriginalName();
            $imagen_nueva->move(public_path('imagenes_usuarios'), $nombreArchivo);
            $datos["imagen"] = $nombreArchivo;

            $eliminarImagen = public_path('imagenes_usuarios/'.$request->input('imagen_anterior'));
            File::delete($eliminarImagen);
        }

        DB::connection('mysql')->table('users')->updateOrInsert(
            ['id' => $datos['id']],
            $datos 
        );
     
        return response()->json(["Datos actualizados correctamente", 1], 200);
    }


    public function cambiarPassword(Request $request){

        $idUsuario = Session::get('id');
        $password_old = $request->input('password_old');
        $password_new = $request->input('password_new');

        $usuario = DB::connection("mysql")->table("users")
        ->where("id", $idUsuario)
        ->where("password", md5($password_old))
        ->first();

        if ($usuario) {
            
            $datos = [
                'id' => $idUsuario,
                'password' =>  md5($password_new)
            ];

            DB::connection('mysql')->table('users')->updateOrInsert(
                ['id' => $datos['id']],
                $datos 
            );
    
            return response()->json(["¡Contraseña actualizada correctamente!", 1], 200);
        }else{
            return response()->json(["¡La contraseña anterior no es correcta!", 0], 200);
        }

    }

    public function consultarUsuarios(){
        $usuario = DB::connection("mysql")->table("users")
        ->select('id', 'identificacion', 'nombre', 'usuario', 'email', 'celular', 'direccion', 'estado', 'password')
        ->get();
        
        return response()->json($usuario, 200);
    }

    public function registrarUsuario(Request $request){

        $datos = [
            'identificacion' => $request->input('identificacion'),
            'nombre' => $request->input('nombre'),
            'usuario' => $request->input('usuario'),
            'email' => $request->input('email'),
            'password' => md5($request->input('password')),
            'celular' => $request->input('celular'),
            'direccion' => $request->input('direccion'),
            'rol' => 'encuestador',
            'imagen' => 'default.png',
        ];

        $insertado = DB::connection('mysql')->table('users')->Insert(
            $datos 
        );

        if ($insertado) {
            return response()->json(["¡Datos actualizados correctamente!", 1], 200);
        }else{
            return response()->json(["¡Ocurrió un error, intente nuevamente!", 1], 200);
        }
     
    }

    public function editarUsuarioAdmin(Request $request){

        $datos = [
            'identificacion' => $request->input('identificacion'),
            'nombre' => $request->input('nombre'),
            'usuario' => $request->input('usuario'),
            'email' => $request->input('email'),
            'password' => md5($request->input('password')),
            'celular' => $request->input('celular'),
            'direccion' => $request->input('direccion'),
            'rol' => 'encuestador',
            'imagen' => 'default.png',
        ];

        $insertado = DB::connection('mysql')->table('users')
        ->where('id', $request->input('id'))
        ->update(
            $datos 
        );

        if ($insertado) {
            return response()->json(["¡Datos actualizados correctamente!", 1], 200);
        }else{
            return response()->json(["¡Ocurrió un error, intente nuevamente!", 1], 200);
        }
     
    }

    public function cambiarEstadoUsuario(Request $request){
        $id = $request->input('id');
        $estado_actual = $request->input('estado_actual');
    
        $datos = [
            'estado' => $estado_actual,
        ];

        $insertado = DB::connection('mysql')->table('users')
        ->where('id', $request->input('id'))
        ->update(
            $datos 
        );

        if ($insertado) {
            return response()->json(["¡Se cambio el estado del usuario correctamente!", 1], 200);
        }else{
            return response()->json(["¡Ocurrió un error, intente nuevamente!", 1], 200);
        }
    }

    public function cerrarSesion(){
        Session::flush();
        return 1;
    }

    public function listarUsuarios(){
        $usuarios = DB::connection("mysql")->table("users")
        ->where("estado", 1)
        ->get();

        return response()->json($usuarios);
    }
}
