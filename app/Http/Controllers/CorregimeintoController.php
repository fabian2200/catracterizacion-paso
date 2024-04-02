<?php

namespace App\Http\Controllers;
use Illuminate\Http\Request;
use DB;
use Illuminate\Support\Facades\File;

use Illuminate\Support\Facades\Session;

class CorregimeintoController extends Controller
{
    public function listarCorregimientos(){
        $corregimientos = DB::connection('mysql')->table('corregimientos')
        ->join("dptos", "dptos.codigo", "corregimientos.departamento")
        ->join("muni", "muni.id", "corregimientos.municipio")
        ->select("corregimientos.id", "corregimientos.nombre as nombre", "dptos.descripcion as nombre_departamento",  "muni.descripcion as nombre_municipio", "corregimientos.departamento", "corregimientos.municipio", "corregimientos.estado")
        ->get();
        
        return response()->json($corregimientos);
    }

    public function guardarCorregimiento(Request $request){
        $datos = [
            'id' => $request->input('id'),
            'departamento' => $request->input('departamento'),
            'municipio' => $request->input('municipio'),
            'nombre' => $request->input('nombre'),
            'estado' => "Activo"
        ];

        $insertado = DB::connection('mysql')->table('corregimientos')->updateOrInsert(
            ['id' => $datos['id']],
            $datos 
        );

        if($insertado){
            return response()->json(['mensaje' => 'Datos guardados correctamente', 'code' => 1]);
        }else{
            return response()->json(['mensaje' => 'Ocurrió un error, intente nuevamente', 'code' => 0]);
        }
    }

    public function listarCorregimientosMunicipio(Request $request){
        $municipio = $request->input('municipio');

        $corregimientos = DB::connection('mysql')->table('corregimientos')
        ->join("dptos", "dptos.codigo", "corregimientos.departamento")
        ->join("muni", "muni.id", "corregimientos.municipio")
        ->select("corregimientos.id", "corregimientos.nombre as nombre", "dptos.descripcion as departamento",  "muni.descripcion as municipio")
        ->where("corregimientos.municipio", $municipio)
        ->where("corregimientos.estado", "Activo")
        ->get();
        
        return response()->json($corregimientos);
    }

    public function elimianrCorregimiento(Request $request){
        $id = $request->input('id');
        $estado = $request->input('estado');

        DB::connection('mysql')->table('corregimientos')
        ->where("id", $id)
        ->update(['estado' => $estado]);

        return response()->json(['mensaje' => '¡Se cambio de estado al corregimiento correctamente!', 'code' => 1]);
    }

}
