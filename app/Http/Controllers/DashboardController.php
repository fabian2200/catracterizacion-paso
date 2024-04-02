<?php

namespace App\Http\Controllers;
use Illuminate\Http\Request;
use DB;
use Illuminate\Support\Facades\File;

use Illuminate\Support\Facades\Session;
use Carbon\Carbon;

class DashboardController extends Controller
{
    public function datosDashboard(){
        $numero_personas = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->count();

        $numero_hogares = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->where("rol", "Jefe de hogar")
        ->count();

        $tenencia_tierras = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", 1)
        ->where("rol", "Jefe de hogar")
        ->count();

        $integrantes = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->get();

        $edad0a4 = [0,0, "0-4"];
        $edad5a9 = [0,0, "5-9"];
        $edad10a14 = [0,0, "10-14"];
        $edad15a19 = [0,0, "15-19"];
        $edad20a24 = [0,0, "20-24"];
        $edad25a29 = [0,0, "25-29"];
        $edad30a34 = [0,0, "30-34"];
        $edad35a39 = [0,0, "35-39"];
        $edad40a44 = [0,0, "40-44"];
        $edad45a49 = [0,0, "45-49"];
        $edad50a54 = [0,0, "50-54"];
        $edad55a59 = [0,0, "55-59"];
        $edad60a64 = [0,0, "60-64"];
        $edad65a69 = [0,0, "65-69"];
        $edad70a74 = [0,0, "70-74"];
        $edad75mas = [0,0, "75+"];

        $masculino = 0;
        $femenino = 0;
        $numero_hogares = 0;
        $desplazados = 0;

        foreach ($integrantes as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
            switch (true) {
                case ($item->edad >= 0 && $item->edad <= 4):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad0a4[0] += 1;
                    }else{
                        $femenino++;
                        $edad0a4[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 5 && $item->edad <= 9):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad5a9[0] += 1;
                    }else{
                        $femenino++;
                        $edad5a9[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 10 && $item->edad <= 14):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad10a14[0] += 1;
                    }else{
                        $femenino++;
                        $edad10a14[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 15 && $item->edad <= 19):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad15a19[0] += 1;
                    }else{
                        $femenino++;
                        $edad15a19[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 20 && $item->edad <= 24):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad20a24[0] += 1;
                    }else{
                        $femenino++;
                        $edad20a24[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 25 && $item->edad <= 29):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad25a29[0] += 1;
                    }else{
                        $femenino++;
                        $edad25a29[1] += 1;
                    }
                    break;
                case ($item->edad >= 30 && $item->edad <= 34):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad30a34[0] += 1;
                    }else{
                        $femenino++;
                        $edad30a34[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 35 && $item->edad <= 39):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad35a39[0] += 1;
                    }else{
                        $femenino++;
                        $edad35a39[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 40 && $item->edad <= 44):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad40a44[0] += 1;
                    }else{
                        $femenino++;
                        $edad40a44[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 45 && $item->edad <= 49):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad45a49[0] += 1;
                    }else{
                        $femenino++;
                        $edad45a49[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 50 && $item->edad <= 54):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad50a54[0] += 1;
                    }else{
                        $femenino++;
                        $edad50a54[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 55 && $item->edad <= 59):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad55a59[0] += 1;
                    }else{
                        $femenino++;
                        $edad55a59[1] += 1;
                    }
                    break;
                case ($item->edad >= 60 && $item->edad <= 64):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad60a64[0] += 1;
                    }else{
                        $femenino++;
                        $edad60a64[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 65 && $item->edad <= 69):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad65a69[0] += 1;
                    }else{
                        $femenino++;
                        $edad65a69[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 70 && $item->edad <= 74):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad70a74[0] += 1;
                    }else{
                        $femenino++;
                        $edad70a74[1] += 1;
                    }
                    break;
                case ($item->edad >= 75):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad75mas[0] += 1;
                    }else{
                        $femenino++;
                        $edad75mas[1] += 1;
                    }
                    break;
            }

            if($item->rol == "Jefe de hogar"){
                $numero_hogares++;
            }

            if($item->desplazado == "Si"){
                $desplazados++;
            }
        }

        $piramide_edad = [
            $edad0a4,
            $edad5a9,
            $edad10a14,
            $edad15a19,
            $edad20a24,
            $edad25a29,
            $edad30a34,
            $edad35a39,
            $edad40a44,
            $edad45a49,
            $edad50a54,
            $edad55a59,
            $edad60a64,
            $edad65a69,
            $edad70a74,
            $edad75mas,
        ];

        $corregimientos = DB::connection('mysql')->table('informacion_personal')
        ->join("corregimientos", "informacion_personal.corregimiento", "corregimientos.id")
        ->select(DB::raw("count(*) as cantidad"), "corregimientos.nombre")
        ->groupBy("corregimientos.nombre")
        ->get();

        $datos = [
            "piramide_edad" => $piramide_edad,
            "numero_masculino" => $masculino,
            "numero_femenino" => $femenino,
            "por_corregimientos" => $corregimientos,
            "numero_hogares" => $numero_hogares,
            "desplazados" => $desplazados
        ];

        return response()->json($datos);
    }

    public function calcularEdad($fechaNacimiento){
        $fechaNacimiento = Carbon::parse($fechaNacimiento);
        $fechaActual = Carbon::now();
        $edad = $fechaNacimiento->diffInYears($fechaActual);
        return $edad;
    }
    
    public function datosInforme(){
        $numero_personas = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->count();

        $numero_hogares = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->where("rol", "Jefe de hogar")
        ->count();

        $tenencia_tierras = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", 1)
        ->where("rol", "Jefe de hogar")
        ->count();

        $integrantes = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->get();

        #region pestaña_sociodemografico
        $edad0a4 = [0,0, "0-4"];
        $edad5a9 = [0,0, "5-9"];
        $edad10a14 = [0,0, "10-14"];
        $edad15a19 = [0,0, "15-19"];
        $edad20a24 = [0,0, "20-24"];
        $edad25a29 = [0,0, "25-29"];
        $edad30a34 = [0,0, "30-34"];
        $edad35a39 = [0,0, "35-39"];
        $edad40a44 = [0,0, "40-44"];
        $edad45a49 = [0,0, "45-49"];
        $edad50a54 = [0,0, "50-54"];
        $edad55a59 = [0,0, "55-59"];
        $edad60a64 = [0,0, "60-64"];
        $edad65a69 = [0,0, "65-69"];
        $edad70a74 = [0,0, "70-74"];
        $edad75mas = [0,0, "75+"];

        $menores15 = [0,0];
        $de16a64 = [0,0];
        $mayores65 = [0,0];

        $masculino = 0;
        $femenino = 0;
        $numero_hogares = 0;
        $desplazados = 0;

        foreach ($integrantes as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
            switch (true) {
                case ($item->edad >= 0 && $item->edad <= 4):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad0a4[0] += 1;
                    }else{
                        $femenino++;
                        $edad0a4[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 5 && $item->edad <= 9):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad5a9[0] += 1;
                    }else{
                        $femenino++;
                        $edad5a9[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 10 && $item->edad <= 14):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad10a14[0] += 1;
                    }else{
                        $femenino++;
                        $edad10a14[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 15 && $item->edad <= 19):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad15a19[0] += 1;
                    }else{
                        $femenino++;
                        $edad15a19[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 20 && $item->edad <= 24):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad20a24[0] += 1;
                    }else{
                        $femenino++;
                        $edad20a24[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 25 && $item->edad <= 29):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad25a29[0] += 1;
                    }else{
                        $femenino++;
                        $edad25a29[1] += 1;
                    }
                    break;
                case ($item->edad >= 30 && $item->edad <= 34):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad30a34[0] += 1;
                    }else{
                        $femenino++;
                        $edad30a34[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 35 && $item->edad <= 39):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad35a39[0] += 1;
                    }else{
                        $femenino++;
                        $edad35a39[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 40 && $item->edad <= 44):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad40a44[0] += 1;
                    }else{
                        $femenino++;
                        $edad40a44[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 45 && $item->edad <= 49):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad45a49[0] += 1;
                    }else{
                        $femenino++;
                        $edad45a49[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 50 && $item->edad <= 54):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad50a54[0] += 1;
                    }else{
                        $femenino++;
                        $edad50a54[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 55 && $item->edad <= 59):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad55a59[0] += 1;
                    }else{
                        $femenino++;
                        $edad55a59[1] += 1;
                    }
                    break;
                case ($item->edad >= 60 && $item->edad <= 64):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad60a64[0] += 1;
                    }else{
                        $femenino++;
                        $edad60a64[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 65 && $item->edad <= 69):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad65a69[0] += 1;
                    }else{
                        $femenino++;
                        $edad65a69[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 70 && $item->edad <= 74):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad70a74[0] += 1;
                    }else{
                        $femenino++;
                        $edad70a74[1] += 1;
                    }
                    break;
                case ($item->edad >= 75):
                    if($item->sexo == "Masculino"){
                        $masculino++;
                        $edad75mas[0] += 1;
                    }else{
                        $femenino++;
                        $edad75mas[1] += 1;
                    }
                    break;
            }

            if($item->rol == "Jefe de hogar"){
                $numero_hogares++;
            }

            if($item->desplazado == "Si"){
                $desplazados++;
            }

            if($item->edad < 15){
                if($item->sexo == "Masculino"){
                    $menores15[0] += 1;
                }else{
                    $menores15[1] += 1;
                }
            }

            if($item->edad >= 15 && $item->edad <= 64){
                if($item->sexo == "Masculino"){
                    $de16a64[0] += 1;
                }else{
                    $de16a64[1] += 1;
                }
            }
            
            if($item->edad >= 65){
                if($item->sexo == "Masculino"){
                    $mayores65[0] += 1;
                }else{
                    $mayores65[1] += 1;
                }
            }
        }

        $piramide_edad = [
            $edad0a4,
            $edad5a9,
            $edad10a14,
            $edad15a19,
            $edad20a24,
            $edad25a29,
            $edad30a34,
            $edad35a39,
            $edad40a44,
            $edad45a49,
            $edad50a54,
            $edad55a59,
            $edad60a64,
            $edad65a69,
            $edad70a74,
            $edad75mas,
            $menores15,
            $de16a64,
            $mayores65
        ];

        $corregimientos = DB::connection('mysql')->table('informacion_personal')
        ->join("corregimientos", "informacion_personal.corregimiento", "corregimientos.id")
        ->select(DB::raw("count(*) as cantidad"), "corregimientos.nombre")
        ->where("informacion_personal.estado", "1")
        ->groupBy("corregimientos.nombre")
        ->get();


        $estado_civil = [
            ["Soltero", 0, 0],
            ["Casado", 0, 0],
            ["Divorciado", 0, 0],
            ["Viudo(a)", 0, 0],
            ["Unión Libre", 0, 0],
            [ "Otro", 0, 0],
            [ "TOTAL PARCIAL", 0, 0],
            [ "M<12", 0, 0],
        ];

        foreach ($integrantes as $item) {
            if($item->edad >= 12){
                switch (true) {
                    case ($item->estado_civil == "Soltero"):
                        $estado_civil[0][1] += 1;
                        break;
                    case ($item->estado_civil == "Casado"):
                        $estado_civil[1][1] += 1;
                        break;
                    case ($item->estado_civil == "Divorciado"):
                        $estado_civil[2][1] += 1;
                        break;
                    case ($item->estado_civil == "Viudo(a)"):
                        $estado_civil[3][1] += 1;
                        break;
                    case ($item->estado_civil == "Unión Libre"):
                        $estado_civil[4][1] += 1;
                        break;
                    case ($item->estado_civil == "Otro"):
                        $estado_civil[5][1] += 1;
                        break;
                }

                $estado_civil[6][1] += 1;
            }else{
                $estado_civil[7][1] += 1;
            }
        }

        if($numero_personas != 0){
            $estado_civil[0][2] = round(($estado_civil[0][1] / $numero_personas) * 100, 2);
            $estado_civil[1][2] = round(($estado_civil[1][1] / $numero_personas) * 100, 2);
            $estado_civil[2][2] = round(($estado_civil[2][1] / $numero_personas) * 100, 2);
            $estado_civil[3][2] = round(($estado_civil[3][1] / $numero_personas) * 100, 2);
            $estado_civil[4][2] = round(($estado_civil[4][1] / $numero_personas) * 100, 2);
            $estado_civil[5][2] = round(($estado_civil[5][1] / $numero_personas) * 100, 2);
            $estado_civil[6][2] = round(($estado_civil[6][1] / $numero_personas) * 100, 2);
            $estado_civil[7][2] = round(($estado_civil[7][1] / $numero_personas) * 100, 2);
        }else{
            $estado_civil[0][2] = 0;
            $estado_civil[1][2] = 0;
            $estado_civil[2][2] = 0;
            $estado_civil[3][2] = 0;
            $estado_civil[4][2] = 0;
            $estado_civil[5][2] = 0;
            $estado_civil[6][2] = 0;
            $estado_civil[7][2] = 0;
        }

        $datos_desplazados = self::poblacionDesplazada();

        #endregion pestaña_sociodemografico

        #region pertaña_empleo
        $poblacion_e_activa = [0,0,0];
        $poblacion_e_inactiva = [0,0,0];
        foreach ($integrantes as $item) {
            if($item->edad >= 15){
                $situacion_laboral = DB::connection('mysql')->table('situacion_laboral')
                ->where("identificacion_individuo", $item->identificacion)
                ->first();
                
                if($situacion_laboral != null){
                   if($situacion_laboral->situacion_laboral != "Desempleado/a (sin búsqueda activa de empleo en el momento)"){
                        if($item->sexo == "Masculino"){
                            $poblacion_e_activa[0] += 1;
                        }else{
                            $poblacion_e_activa[1] += 1;
                        }
                        $poblacion_e_activa[2] += 1;
                    }else{
                        if($item->sexo == "Masculino"){
                            $poblacion_e_inactiva[0] += 1;
                        }else{
                            $poblacion_e_inactiva[1] += 1;
                        }
                        $poblacion_e_inactiva[2] += 1;
                    } 
                }
            }
        }


        $nivel_educativo = DB::connection('mysql')->table('informacion_personal')
        ->join("educacion", "educacion.identificacion_individuo", "informacion_personal.identificacion")
        ->join("escolaridad", "escolaridad.id", "educacion.nivel_educativo")
        ->where("informacion_personal.estado", "1")
        ->select(DB::raw("count(*) as cantidad"), "escolaridad.id", "escolaridad.descripcion")
        ->whereRaw("DATEDIFF(CURDATE(), informacion_personal.fecha_nacimiento) / 365 >= ?", [5])
        ->groupBy("escolaridad.id", "escolaridad.descripcion")
        ->get();
        
        $menores_5a = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", "1")
        ->whereRaw("DATEDIFF(CURDATE(), fecha_nacimiento) / 365 < ?", [5])
        ->count();


        $nivel_educativo_sexo = [];

        $hgc = 0;
        $mgc = 0;
        foreach ($nivel_educativo as $item) {
            $hombres_cant = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "educacion.identificacion_individuo", "informacion_personal.identificacion")
            ->whereRaw("DATEDIFF(CURDATE(), informacion_personal.fecha_nacimiento) / 365 >= ?", [5])
            ->where("informacion_personal.estado", "1")
            ->where("informacion_personal.sexo", "Masculino")
            ->where("educacion.nivel_educativo", $item->id)
            ->count();

            $mujeres_cant = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "educacion.identificacion_individuo", "informacion_personal.identificacion")
            ->whereRaw("DATEDIFF(CURDATE(), informacion_personal.fecha_nacimiento) / 365 >= ?", [5])
            ->where("informacion_personal.estado", "1")
            ->where("informacion_personal.sexo", "Femenino")
            ->where("educacion.nivel_educativo", $item->id)
            ->count();

            $nivel_educativo_sexo[] = [
                "nivel_educativo" => $item->descripcion,
                "mujeres" => $mujeres_cant,
                "hombres" => $hombres_cant
            ];

            $hgc += $hombres_cant;
            $mgc += $mujeres_cant;
        }

        $nivel_educativo[] = [
            "cantidad" => $menores_5a,
            "id" => null,
            "descripcion" => "M < 5 Años"
        ];

        $por_ocupacion = DB::connection('mysql')->table('informacion_personal')
        ->join("situacion_laboral", "situacion_laboral.identificacion_individuo", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", "1")
        ->select(DB::raw("count(*) as cantidad"), "situacion_laboral.ocupacion")
        ->groupBy("situacion_laboral.ocupacion")
        ->orderby("cantidad", "desc")
        ->get();

        $trabajo_formal = DB::connection('mysql')->table('informacion_personal')
        ->join("situacion_laboral", "situacion_laboral.identificacion_individuo", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", "1")
        ->where("situacion_laboral.situacion_laboral", "<>", "N.A")
        ->where("situacion_laboral.situacion_laboral", "<>", "Trabajador informal")
        ->where("situacion_laboral.situacion_laboral", "<>", "Desempleado/a (en búsqueda activa de empleo)")
        ->where("situacion_laboral.situacion_laboral", "<>", "Desempleado/a (sin búsqueda activa de empleo en el momento)")
        ->count();

        $trabajo_formal_m = DB::connection('mysql')->table('informacion_personal')
        ->join("situacion_laboral", "situacion_laboral.identificacion_individuo", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", "1")
        ->where("informacion_personal.sexo", "Masculino")
        ->where("situacion_laboral.situacion_laboral", "<>", "N.A")
        ->where("situacion_laboral.situacion_laboral", "<>", "Trabajador informal")
        ->where("situacion_laboral.situacion_laboral", "<>", "Desempleado/a (en búsqueda activa de empleo)")
        ->where("situacion_laboral.situacion_laboral", "<>", "Desempleado/a (sin búsqueda activa de empleo en el momento)")
        ->count();

        $trabajo_formal_f = DB::connection('mysql')->table('informacion_personal')
        ->join("situacion_laboral", "situacion_laboral.identificacion_individuo", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", "1")
        ->where("informacion_personal.sexo", "Femenino")
        ->where("situacion_laboral.situacion_laboral", "<>", "N.A")
        ->where("situacion_laboral.situacion_laboral", "<>", "Trabajador informal")
        ->where("situacion_laboral.situacion_laboral", "<>", "Desempleado/a (en búsqueda activa de empleo)")
        ->where("situacion_laboral.situacion_laboral", "<>", "Desempleado/a (sin búsqueda activa de empleo en el momento)")
        ->count();

        $trabajo_informal = DB::connection('mysql')->table('informacion_personal')
        ->join("situacion_laboral", "situacion_laboral.identificacion_individuo", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", "1")
        ->where("situacion_laboral.situacion_laboral", "=", "Trabajador informal")
        ->count();

        $trabajo_informal_m = DB::connection('mysql')->table('informacion_personal')
        ->join("situacion_laboral", "situacion_laboral.identificacion_individuo", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", "1")
        ->where("informacion_personal.sexo", "Masculino")
        ->where("situacion_laboral.situacion_laboral", "=", "Trabajador informal")
        ->count();

        $trabajo_informal_f = DB::connection('mysql')->table('informacion_personal')
        ->join("situacion_laboral", "situacion_laboral.identificacion_individuo", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", "1")
        ->where("informacion_personal.sexo", "Femenino")
        ->where("situacion_laboral.situacion_laboral", "=", "Trabajador informal")
        ->count();

        $situacion_laboral = [
            "trabajo_formal" => $trabajo_formal,
            "trabajo_formal_m" => $trabajo_formal_m,
            "trabajo_formal_f" => $trabajo_formal_f,
            "trabajo_informal" => $trabajo_informal,
            "trabajo_informal_m" => $trabajo_informal_m,
            "trabajo_informal_f" => $trabajo_informal_f
        ];

        #endregion pertaña_empleo

        #region pestaña_vivienda

        $posesion_vivienda_si = DB::connection('mysql')->table('vivienda_hogar')
        ->join("informacion_personal", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
        ->where("informacion_personal.estado", "1")
        ->where("tenencia", "<>", "Alquilada")
        ->count();

        $posesion_vivienda_no = DB::connection('mysql')->table('vivienda_hogar')
        ->join("informacion_personal", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
        ->where("informacion_personal.estado", "1")
        ->where("tenencia", "=", "Alquilada")
        ->count();

        $posesion_baldios = [];
        
        $baldios_si = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("vivienda_hogar.posecion_baldia", "Si")
        ->where("informacion_personal.estado", 1)
        ->count();

        $baldios_no = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("vivienda_hogar.posecion_baldia", "No")
        ->where("informacion_personal.estado", 1)
        ->count();

        $titulo_si = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("vivienda_hogar.propiedad_titulo", "Si")
        ->where("informacion_personal.estado", 1)
        ->count();

        $titulo_no = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("vivienda_hogar.propiedad_titulo", "No")
        ->where("informacion_personal.estado", 1)
        ->count();

        $total_area = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", 1)
        ->sum('vivienda_hogar.area_total');

        $posesion_baldios[] = [
            "baldios_si" => $baldios_si,
            "baldios_no" => $baldios_no,
            "titulo_si" => $baldios_si,
            "titulo_no" => $baldios_no,
            "total_area" => $total_area
        ];

        $cultivos = DB::connection('mysql')->table('informacion_personal')
        ->join("actividades_vivienda_hogar", "actividades_vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", 1)
        ->groupBy("actividades_vivienda_hogar.linea", "actividades_vivienda_hogar.actividad")
        ->selectRaw('SUM(actividades_vivienda_hogar.area_destinada) as area_destinada, actividades_vivienda_hogar.linea, actividades_vivienda_hogar.actividad')
        ->orderBy("area_destinada", "DESC")
        ->get();

        $electricidad = [0, 0, "Electricidad"];
        $agua_potable = [0, 0, "Agua Potable"];
        $alcantarillado = [0, 0, "Alcantarillado"];
        $gas_natural = [0, 0, "Gas Natural"];
        $aseo = [0, 0, "Aseo"];
        $otro = [0, 0, "Otro"];

        $viviendas = DB::connection('mysql')->table('vivienda_hogar')
        ->join("informacion_personal", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
        ->where("informacion_personal.estado", "1")
        ->get();

        foreach ($viviendas as $item) {
            if($item->electricidad == "Si"){
                $electricidad[0] += 1;
            }else{
                $electricidad[1] += 1;
            }

            if($item->agua_potable == "Si"){
                $agua_potable[0] += 1;
            }else{
                $agua_potable[1] += 1;
            }

            if($item->alcantarillado == "Si"){
                $alcantarillado[0] += 1;
            }else{
                $alcantarillado[1] += 1;
            }

            if($item->gas_natural == "Si"){
                $gas_natural[0] += 1;
            }else{
                $gas_natural[1] += 1;
            }

            if($item->aseo == "Si"){
                $aseo[0] += 1;
            }else{
                $aseo[1] += 1;
            }

            if($item->otro == "Si"){
                $otro[0] += 1;
            }else{
                $otro[1] += 1;
            }
        }

        $servicios = [
            $electricidad,
            $agua_potable,
            $alcantarillado,
            $gas_natural,
            $aseo,
            $otro
        ];


        $tipo_vivienda = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", 1)
        ->selectRaw('COUNT(vivienda_hogar.id) as cantidad, vivienda_hogar.tipo_vivienda')
        ->groupBy("tipo_vivienda",)
        ->orderBy("cantidad", "DESC")
        ->get();

        $cantidad_viviendas = DB::connection('mysql')->table('informacion_personal')
        ->join("vivienda_hogar", "vivienda_hogar.identificacion_jefe", "informacion_personal.identificacion")
        ->where("informacion_personal.estado", 1)
        ->count();

        #endregion pestaña_vivienda


        #region afro y negritudes
            $afrocolombianos_negritudes = self::afrocolombianosyNegritudes();
        #endregion afro


        #region salud
            $salud = self::datosSalud();
        #endregion salud


        $datos = [
            "piramide_edad" => $piramide_edad,
            "numero_masculino" => $masculino,
            "numero_femenino" => $femenino,
            "por_corregimientos" => $corregimientos,
            "numero_hogares" => $numero_hogares,
            "desplazados" => $desplazados,
            "numero_personas" => $numero_personas,
            "estado_civil" => $estado_civil,
            "poblacion_e_activa" => $poblacion_e_activa,
            "poblacion_e_inactiva" => $poblacion_e_inactiva,
            "nivel_educativo" => $nivel_educativo,
            "nivel_educativo_sexo" => $nivel_educativo_sexo,
            "personas_nivel_educativo" => [
                "total" =>  ($hgc + $mgc),
                "total_hombres" => $hgc,
                "total_mujeres" => $mgc,
            ],
            "por_ocupacion" => $por_ocupacion,
            "situacion_laboral" => $situacion_laboral,
            "posesion_vivienda" => [
                "poseen" => $posesion_vivienda_si,
                "no_poseen" => $posesion_vivienda_no
            ],
            "posesion_baldios" => $posesion_baldios,
            "cultivos" => $cultivos,
            "datos_desplazados" => $datos_desplazados,
            "servicios" => $servicios,
            "tipo_vivienda" => $tipo_vivienda,
            "cantidad_viviendas" => $cantidad_viviendas,
            "afrocolombianos_negritudes" => $afrocolombianos_negritudes,
            "salud" => $salud
        ];

        return response()->json($datos);
    }

    public function poblacionDesplazada(){

        $cantidad_poblacion_no_desplazada = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->where("informacion_personal.desplazado", "No")
        ->count();

        $poblacion_desplazada = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->where("informacion_personal.desplazado", "Si")
        ->get();

        $cantidad_poblacion_desplazada = count($poblacion_desplazada);

        $masculinos_desplazados = 0;
        $femeninos_desplazados = 0;
        $desplazados_m15 = [0,0, "< 15 Años"];
        $desplazados_15a64 = [0,0, "15 a 64 Años"];
        $desplazados_mas65 = [0,0, "> 65 Años"];

        foreach ($poblacion_desplazada as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
            switch (true) {
                case ($item->edad < 15):
                    if($item->sexo == "Masculino"){
                        $masculinos_desplazados++;
                        $desplazados_m15[0] += 1;
                    }else{
                        $femeninos_desplazados++;
                        $desplazados_m15[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 15 && $item->edad <= 64):
                    if($item->sexo == "Masculino"){
                        $masculinos_desplazados++;
                        $desplazados_15a64[0] += 1;
                    }else{
                        $femeninos_desplazados++;
                        $desplazados_15a64[1] += 1;
                    }
                    break;
            
                case ($item->edad >= 65):
                    if($item->sexo == "Masculino"){
                        $masculinos_desplazados++;
                        $desplazados_mas65[0] += 1;
                    }else{
                        $femeninos_desplazados++;
                        $desplazados_mas65[1] += 1;
                    }
                    break;
            }           
        }

        $desplazados = [
            "cantidad_poblacion_no_desplazada" => $cantidad_poblacion_no_desplazada,
            "cantidad_poblacion_desplazada" => $cantidad_poblacion_desplazada,
            "masculinos_desplazados" => $masculinos_desplazados,
            "femeninos_desplazados" => $femeninos_desplazados,
            "desplazados_m15" => $desplazados_m15,
            "desplazados_15a64" => $desplazados_15a64,
            "desplazados_mas65" => $desplazados_mas65,
        ];

        return $desplazados;
    }

    public function afrocolombianosyNegritudes(){

        $cantidad_poblacion = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->count();

        $por_concejo = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(cultura_tradiciones.identificacion_individuo) as cantidad, cultura_tradiciones.concejo')
        ->groupBy("cultura_tradiciones.concejo")
        ->get();

        $por_etnia = DB::connection('mysql')->table('origen_etnia')
        ->join("informacion_personal", "informacion_personal.identificacion", "origen_etnia.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(origen_etnia.identificacion_individuo) as cantidad, origen_etnia.etnia')
        ->groupBy("origen_etnia.etnia")
        ->get();

        $por_practicas_religiosas = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(cultura_tradiciones.identificacion_individuo) as cantidad, cultura_tradiciones.practicas_religiosas')
        ->groupBy("cultura_tradiciones.practicas_religiosas")
        ->get();

        $por_practica_culturales = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(cultura_tradiciones.identificacion_individuo) as cantidad, cultura_tradiciones.practica_actividades')
        ->groupBy("cultura_tradiciones.practica_actividades")
        ->get();

        $habla_lenguas = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(cultura_tradiciones.identificacion_individuo) as cantidad, cultura_tradiciones.habla_lengua')
        ->groupBy("cultura_tradiciones.habla_lengua")
        ->get();

        $cual_lengua_habla = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->where("cultura_tradiciones.cual_lengua", "<>", null)
        ->selectRaw('COUNT(cultura_tradiciones.identificacion_individuo) as cantidad, cultura_tradiciones.cual_lengua')
        ->groupBy("cultura_tradiciones.cual_lengua")
        ->get();


        $afrocolombianos_negritudes = [
            "por_concejo" => $por_concejo,
            "por_etnia" => $por_etnia,
            "por_practicas_religiosas" => $por_practicas_religiosas,
            "por_practica_culturales" => $por_practica_culturales,
            "habla_lenguas" => $habla_lenguas,
            "cual_lengua_habla" => $cual_lengua_habla,
            "cantidad_poblacion" => $cantidad_poblacion
        ];

        return $afrocolombianos_negritudes;
    }

    public function datosSalud(){
        $cantidad_poblacion = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", 1)
        ->count();

        $estado_salud = DB::connection('mysql')->table('salud')
        ->join("informacion_personal", "informacion_personal.identificacion", "salud.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(salud.identificacion_individuo) as cantidad, salud.estado_salud')
        ->groupBy("salud.estado_salud")
        ->orderBy("cantidad", "DESC")
        ->get();

        $condicion_discapacidad = DB::connection('mysql')->table('salud')
        ->join("informacion_personal", "informacion_personal.identificacion", "salud.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(salud.identificacion_individuo) as cantidad, salud.condicion_discapacidad')
        ->groupBy("salud.condicion_discapacidad")
        ->orderBy("cantidad", "DESC")
        ->get();

        $acceso_salud = DB::connection('mysql')->table('salud')
        ->join("informacion_personal", "informacion_personal.identificacion", "salud.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(salud.identificacion_individuo) as cantidad, salud.acceso_salud')
        ->groupBy("salud.acceso_salud")
        ->orderBy("cantidad", "DESC")
        ->get();

        $regimen = DB::connection('mysql')->table('salud')
        ->join("informacion_personal", "informacion_personal.identificacion", "salud.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(salud.identificacion_individuo) as cantidad, salud.regimen')
        ->groupBy("salud.regimen")
        ->orderBy("cantidad", "DESC")
        ->get();

        $discriminacion_salud = DB::connection('mysql')->table('salud')
        ->join("informacion_personal", "informacion_personal.identificacion", "salud.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->selectRaw('COUNT(salud.identificacion_individuo) as cantidad, salud.discriminacion_salud')
        ->groupBy("salud.discriminacion_salud")
        ->orderBy("cantidad", "DESC")
        ->get();

        $salud = [
            "estado_salud" => $estado_salud,
            "condicion_discapacidad" => $condicion_discapacidad,
            "acceso_salud" => $acceso_salud,
            "regimen" => $regimen,
            "discriminacion_salud" => $discriminacion_salud,
            "cantidad_poblacion" => $cantidad_poblacion
        ];

        return $salud;
    }
}