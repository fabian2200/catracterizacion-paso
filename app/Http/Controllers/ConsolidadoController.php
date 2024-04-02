<?php

namespace App\Http\Controllers;
use Illuminate\Http\Request;
use DB;
use File;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use Carbon\Carbon;

use Illuminate\Support\Facades\Session;

class ConsolidadoController extends Controller
{
    public function consolidadoCensados(Request $request){

        $individuos = [];
        $corregimiento = $request->input('corregimiento');

        $pagina = $request->input('pagina');
        $limit = 10;

        if ($pagina == null || $pagina == "") {
            $pagina = 1;
        }

        if ($pagina == 1) {
            $offset = 0;
            $pagina--;
        } else {
            $pagina--;
            $offset = $pagina * $limit;
        }

        if($corregimiento == "todo"){
            $jefes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->join("vivienda_hogar", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
            ->where("rol", "Jefe de hogar")
            ->offset($offset)
            ->limit($limit) 
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
    
            $integrantes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->where("rol", "<>", "Jefe de hogar")
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
        }else{
            $jefes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->join("vivienda_hogar", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
            ->where("rol", "Jefe de hogar")
            ->where("informacion_personal.corregimiento", $corregimiento)
            ->offset($offset)
            ->limit($limit) 
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
    
            $integrantes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->where("rol", "<>", "Jefe de hogar")
            ->where("informacion_personal.corregimiento", $corregimiento)
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
        }

        
        foreach ($integrantes as $integrante) {
            $integrante->escolaridad = DB::connection('mysql')->table('escolaridad')
            ->where("id", $integrante->nivel_educativo)
            ->first()->descripcion;
        }

        $numero_familia = 1;
        foreach ($jefes as $jefe) {
            
            $rowspan = 1;
            $numero_orden = 1;
          
            $jefe->numero_orden = $numero_orden;
            $individuos[] = $jefe;

            $numero_orden++;

            if($jefe->tipo_vivienda == "Finca" || $jefe->tipo_vivienda == "Parcela"){
                $jefe->area_destinada_ocupada =  DB::connection('mysql')->table('actividades_vivienda_hogar')
                ->where("identificacion_jefe", $jefe->identificacion)
                ->sum("area_destinada");
            }

            $jefe->escolaridad = DB::connection('mysql')->table('escolaridad')
            ->where("id", $jefe->nivel_educativo)
            ->first()->descripcion;

            foreach ($integrantes as $key => $integrante) {
                if ($integrante->id_jefe == $jefe->identificacion) {
                    $integrante->numero_familia = $numero_familia;
                    $integrante->numero_orden = $numero_orden;
                    $numero_orden++;
                    $individuos[] = $integrante;
                    unset($integrantes[$key]);
                    $rowspan++;
                }
            }

            $jefe->numero_familia = $numero_familia;
            $jefe->rowspan = $rowspan;

            $numero_familia += 1;
        }

        return response()->json($individuos);
    }

    public function paginacionConsolidado(Request $request){
       
        $corregimiento = $request->input('corregimiento');
        $pagina = $request->input('pagina');

        if($corregimiento == "todo"){
            $jefes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->join("vivienda_hogar", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
            ->where("rol", "Jefe de hogar")
            ->count();
        }else{
            $jefes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->join("vivienda_hogar", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
            ->where("rol", "Jefe de hogar")
            ->where("informacion_personal.corregimiento", $corregimiento)
            ->count();
        }


        $paginas = ceil($jefes / 10);
        
        $respuesta = [
            'paginas_consulta' => $paginas,
            'pagina_actual' => (int)$pagina,
            'numero_registros' => $jefes,
        ];
        
        return response()->json($respuesta, 200);
    }

    public function municipiosConsolidado(){
        $municipios = [];

        $municipios = DB::connection('mysql')->table('informacion_personal')
        ->join("dptos", "informacion_personal.departamento", "dptos.codigo")
        ->join("muni", "informacion_personal.municipio", "muni.id")
        ->join("corregimientos", "informacion_personal.corregimiento", "corregimientos.id")
        ->select(DB::raw("count(*) as cantidad"), "corregimientos.nombre", "dptos.descripcion as departamento", "muni.descripcion as municipio", "corregimientos.id as id_corregimiento")
        ->groupBy("corregimientos.nombre", "departamento", "municipio", "id_corregimiento")
        ->get();

        return response()->json($municipios);
    }

    public function exportarConsolidadoExcel(Request $request) {
        $idCorregimiento = $request->input('corregimiento');
        $tipo_reporte = $request->input('tipo_reporte');

        $datos = self::consultarIntegrantesJefes($idCorregimiento);

        $styleArray = [
            'font' => [
                'bold' => true,
                'size' => 10
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'], // Black color
                ],
            ],
        ];

        $styleArray2 = [
            'font' => [
                'bold' => true,
                'size' => 10,
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'indent' => 1,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'], // Black color
                ],
            ],
            'padding' => [
                'left' => 20, // Set your desired left padding value
            ],
        ];
    
    
        $titulo = "reporte_caracterizacion";

        $nombre = $titulo.".xlsx";

        if($idCorregimiento == "todo"){
            $corregimiento = "Consolidado General";   
        }else{
            $corregimiento = DB::connection('mysql')->table('corregimientos')
            ->where("id", $idCorregimiento)
            ->first()->nombre;
        }
       
        $spreadsheet = new Spreadsheet();
            
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Consolidado General');

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('Z')->setAutoSize(true);
       

        $sheet->mergeCells('A1:C2');
        $sheet->getRowDimension(1)->setRowHeight(60); 
        $sheet->getRowDimension(2)->setRowHeight(30);
        $sheet->getRowDimension(3)->setRowHeight(30);
        $sheet->getRowDimension(4)->setRowHeight(30);
        $sheet->getRowDimension(5)->setRowHeight(30);
        $sheet->getRowDimension(6)->setRowHeight(20);

        $spreadsheet->getActiveSheet()->getStyle('A1:C2')->applyFromArray($styleArray);

        $sheet->setCellValue('D1', 'MACROPROCESO: AGENCIA NACIONAL DE TIERRAS '. PHP_EOL . 'PROCESO: DIRECCIÓN DE ASUNTOS ÉTNICOS');
        $sheet->mergeCells('D1:O1');
        $spreadsheet->getActiveSheet()->getStyle('D1:O1')->applyFromArray($styleArray);
        $style = $sheet->getStyle('D1:O1');
        $alignment = $style->getAlignment();
        $alignment->setWrapText(true);

        $sheet->setCellValue('D2', 'FORMATO: CENSO POBLACIÓN COMUNIDADES NEGRAS');
        $sheet->mergeCells('D2:O2');
        $spreadsheet->getActiveSheet()->getStyle('D2:O2')->applyFromArray($styleArray);
       
        $sheet->setCellValue('P1', 'CÓDIGO: '. PHP_EOL . 'FECHA:');
        $sheet->mergeCells('P1:S1');
        $spreadsheet->getActiveSheet()->getStyle('P1:S1')->applyFromArray($styleArray);
        $style = $sheet->getStyle('P1:S1');
        $alignment = $style->getAlignment();
        $alignment->setWrapText(true);

        $sheet->setCellValue('P2', 'PAGINA: 1/1');
        $sheet->mergeCells('P2:S2');
        $spreadsheet->getActiveSheet()->getStyle('P2:S2')->applyFromArray($styleArray);

        $sheet->setCellValue('A3', 'DATOS DE IDENTIFICACIÓN');
        $sheet->mergeCells('A3:S3');
        $spreadsheet->getActiveSheet()->getStyle('A3:S3')->applyFromArray($styleArray);

        $sheet->setCellValue('A4', '1. Departamento: CESAR');
        $sheet->mergeCells('A4:C4');
        $spreadsheet->getActiveSheet()->getStyle('A4:C4')->applyFromArray($styleArray2);

        $sheet->setCellValue('D4', '2. Municipio: EL PASO');
        $sheet->mergeCells('D4:F4');
        $spreadsheet->getActiveSheet()->getStyle('D4:F4')->applyFromArray($styleArray2);

        $sheet->setCellValue('G4', '3. Corregimiento o vereda:   '.$corregimiento);
        $sheet->mergeCells('G4:O4');
        $spreadsheet->getActiveSheet()->getStyle('G4:O4')->applyFromArray($styleArray2);

        $sheet->setCellValue('P4', '4. Fecha: 14 de Noviembre del 2023');
        $sheet->mergeCells('P4:S4');
        $spreadsheet->getActiveSheet()->getStyle('P4:S4')->applyFromArray($styleArray2);

        $sheet->setCellValue('A5', '5. Representante Legal: ');
        $sheet->mergeCells('A5:F5');
        $spreadsheet->getActiveSheet()->getStyle('A5:F5')->applyFromArray($styleArray2);

        $sheet->setCellValue('G5', '6. Consejo Comunitario: Todos');
        $sheet->mergeCells('G5:S5');
        $spreadsheet->getActiveSheet()->getStyle('G5:S5')->applyFromArray($styleArray2);
        

        $sheet->setCellValue('A6', 'Listado de Personas');
        $sheet->mergeCells('A6:B6');
        $spreadsheet->getActiveSheet()->getStyle('A6:B6')->applyFromArray($styleArray);
        $sheet->getStyle('A6:B6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A6:B6')->getFill()->getStartColor()->setARGB('8cee8c'); 
        $sheet->getStyle('A6')->getAlignment()->setWrapText(true);

        $sheet->setCellValue('C6', 'Datos');
        $sheet->mergeCells('C6:F6');
        $spreadsheet->getActiveSheet()->getStyle('C6:F6')->applyFromArray($styleArray);
        $sheet->getStyle('C6:F6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('C6:F6')->getFill()->getStartColor()->setARGB('43d7f0'); 

        $sheet->setCellValue('G6', '');
        $sheet->mergeCells('G6:L6');
        $spreadsheet->getActiveSheet()->getStyle('G6:L6')->applyFromArray($styleArray);
        $sheet->getStyle('G6:L6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('G6:L6')->getFill()->getStartColor()->setARGB('cdf043'); 

        $sheet->setCellValue('M6', '');
        $sheet->mergeCells('M6:N6');
        $spreadsheet->getActiveSheet()->getStyle('M6:N6')->applyFromArray($styleArray);
        $sheet->getStyle('M6:N6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('M6:N6')->getFill()->getStartColor()->setARGB('43d7f0'); 

        $sheet->setCellValue('O6', 'Tenencia de Tierras');
        $sheet->mergeCells('O6:S6');
        $spreadsheet->getActiveSheet()->getStyle('O6:S6')->applyFromArray($styleArray);
        $sheet->getStyle('O6:S6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('O6:S6')->getFill()->getStartColor()->setARGB('8cee8c'); 

        $sheet->setCellValue('A7', 'Nº Familia');
        $sheet->setCellValue('B7', 'Nº Orden');
        $spreadsheet->getActiveSheet()->getStyle('A7:B7')->applyFromArray($styleArray);
        $sheet->getStyle('A7:B7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A7:B7')->getFill()->getStartColor()->setARGB('8cee8c'); 


        $sheet->setCellValue('C7', 'Nombre Completo');
        $sheet->setCellValue('D7', 'Tipo');
        $sheet->setCellValue('E7', 'Nº Documento');
        $sheet->setCellValue('F7', 'Dirección');
        $spreadsheet->getActiveSheet()->getStyle('C7:F7')->applyFromArray($styleArray);
        $sheet->getStyle('C6:F7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('C6:F7')->getFill()->getStartColor()->setARGB('43d7f0');

        $sheet->setCellValue('G7', 'Sexo');
        $sheet->setCellValue('H7', 'Edad');
        $sheet->setCellValue('I7', 'Estado Civil');
        $sheet->setCellValue('J7', 'Parentesco');
        $sheet->setCellValue('K7', 'Ocupación');
        $sheet->setCellValue('L7', 'Escolaridad');
        $spreadsheet->getActiveSheet()->getStyle('G7:L7')->applyFromArray($styleArray);
        $sheet->getStyle('G7:L7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('G7:L7')->getFill()->getStartColor()->setARGB('cdf043'); 


        $sheet->setCellValue('M7', 'Tipo '. PHP_EOL . ' Vivienda');
        $sheet->setCellValue('N7', 'Tenencia '. PHP_EOL . ' Vivienda');
        $spreadsheet->getActiveSheet()->getStyle('M7:N7')->applyFromArray($styleArray);
        $sheet->getStyle('M6:N7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('M6:N7')->getFill()->getStartColor()->setARGB('43d7f0');

        $sheet->setCellValue('O7', 'Posesión '. PHP_EOL . ' Baldíos');
        $sheet->setCellValue('P7', 'Propiedad '. PHP_EOL . ' con Título');
        $sheet->setCellValue('Q7', 'Area M2');
        $sheet->setCellValue('R7', 'Area Ha');
        $sheet->setCellValue('S7', 'Area Utilizada');
        $spreadsheet->getActiveSheet()->getStyle('O7:S7')->applyFromArray($styleArray);
        $sheet->getStyle('O7:S7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('O7:S7')->getFill()->getStartColor()->setARGB('8cee8c'); 


        $sheet->getStyle('O7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('O7')->getAlignment()->setTextRotation(90);
        $sheet->getStyle('M7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('M7')->getAlignment()->setTextRotation(90);
        $sheet->getStyle('N7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('N7')->getAlignment()->setTextRotation(90);
        $sheet->getStyle('P7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('P7')->getAlignment()->setTextRotation(90);
        

        $sheet->getColumnDimension('A')->setWidth(200);
        $sheet->getColumnDimension('B')->setWidth(200);
        $sheet->getColumnDimension('K')->setWidth(20);
        $sheet->getColumnDimension('L')->setWidth(25);

        // Add an image to cell A1
        $imagePath = public_path('imagenes/censo.png'); // Use public_path to get the correct path
        $drawing = new Drawing();
        $drawing->setPath($imagePath);
        $drawing->setCoordinates('A1');
        $drawing->setOffsetX(80);
        $drawing->setOffsetY(13);
        $drawing->setWidth(180); 
        $drawing->setHeight(100);
        $drawing->setWorksheet($sheet);
        
        if($tipo_reporte == "general"){
            $row = 8;
            $index = 0;
            $orden = 1;

            foreach ($datos as $dato) {
                $sheet->getRowDimension($row)->setRowHeight(30);

                $sheet->setCellValue('A'.$row, $dato->numero_familia);
                $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
            
                if($dato->rol == "Jefe de hogar"){
                    $orden = 1;
                }else{
                    $orden++;
                }

                $sheet->setCellValue('B'.$row, $orden);
                $spreadsheet->getActiveSheet()->getStyle('B'.$row)->applyFromArray($styleArray);

                $sheet->setCellValue('C'.$row, $dato->nombre_completo);
                $spreadsheet->getActiveSheet()->getStyle('C'.$row)->applyFromArray($styleArray);
                $sheet->getStyle('C'.$row)->getAlignment()->setWrapText(true);

                $sheet->setCellValue('D'.$row, $dato->tipo_identificacion);
                $spreadsheet->getActiveSheet()->getStyle('D'.$row)->applyFromArray($styleArray);

                $sheet->setCellValue('E'.$row, $dato->identificacion);
                $spreadsheet->getActiveSheet()->getStyle('E'.$row)->applyFromArray($styleArray);
                $spreadsheet->getActiveSheet()->getStyle('E'.$row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
            
                $sheet->setCellValue('F'.$row, $dato->direccion);
                $spreadsheet->getActiveSheet()->getStyle('F'.$row)->applyFromArray($styleArray);

                $sheet->setCellValue('G'.$row, $dato->sexo);
                $spreadsheet->getActiveSheet()->getStyle('G'.$row)->applyFromArray($styleArray);

                $sheet->setCellValue('H'.$row, self::calcularEdad($dato->fecha_nacimiento)." Años");
                $spreadsheet->getActiveSheet()->getStyle('H'.$row)->applyFromArray($styleArray);

                $sheet->setCellValue('I'.$row, $dato->estado_civil);
                $spreadsheet->getActiveSheet()->getStyle('I'.$row)->applyFromArray($styleArray);

                $sheet->setCellValue('J'.$row, $dato->rol);
                $spreadsheet->getActiveSheet()->getStyle('J'.$row)->applyFromArray($styleArray);

                $sheet->setCellValue('K'.$row, $dato->ocupacion);
                $spreadsheet->getActiveSheet()->getStyle('K'.$row)->applyFromArray($styleArray);
                $sheet->getStyle('K'.$row)->getAlignment()->setWrapText(true);

                $sheet->setCellValue('L'.$row, $dato->escolaridad);
                $spreadsheet->getActiveSheet()->getStyle('L'.$row)->applyFromArray($styleArray);
                $sheet->getStyle('L'.$row)->getAlignment()->setWrapText(true);

                if($dato->rol == "Jefe de hogar"){
                    $sheet->setCellValue('M'.$row, $dato->tipo_vivienda);
                    $spreadsheet->getActiveSheet()->getStyle('M'.$row)->applyFromArray($styleArray);
            
                    $sheet->setCellValue('N'.$row, $dato->tenencia);
                    $spreadsheet->getActiveSheet()->getStyle('N'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('O'.$row, $dato->posecion_baldia == "" ? "N.A" : $dato->posecion_baldia );
                    $spreadsheet->getActiveSheet()->getStyle('O'.$row)->applyFromArray($styleArray);
                    
                    $sheet->setCellValue('P'.$row, $dato->propiedad_titulo == "" ? "N.A" : $dato->propiedad_titulo );
                    $spreadsheet->getActiveSheet()->getStyle('P'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('Q'.$row, $dato->area_total == "" ? "N.A" : $dato->area_total);
                    $spreadsheet->getActiveSheet()->getStyle('Q'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('R'.$row, $dato->area_total == "" ? "N.A" : ((double)$dato->area_total / 10000));
                    $spreadsheet->getActiveSheet()->getStyle('R'.$row)->applyFromArray($styleArray);

                    if($dato->tipo_vivienda == "Finca" || $dato->tipo_vivienda == "Parcela"){
                        $sheet->setCellValue('S'.$row, $dato->area_destinada_ocupada);
                        $spreadsheet->getActiveSheet()->getStyle('S'.$row)->applyFromArray($styleArray);
                    }else{
                        $sheet->setCellValue('S'.$row, "N.A");
                        $spreadsheet->getActiveSheet()->getStyle('S'.$row)->applyFromArray($styleArray);
                    }
                    
                }

                $row++; 
            }

            $row = 8;
            foreach ($datos as $dato) {
                if($dato->rol == "Jefe de hogar"){
                    $sheet->mergeCells('A'.$row.':A'.($row + ($dato->rowspan - 1)));
                    $sheet->mergeCells('M'.$row.':M'.($row + ($dato->rowspan - 1)));
                    $spreadsheet->getActiveSheet()->getStyle('M'.$row.':M'.($row + ($dato->rowspan - 1)))->applyFromArray($styleArray);
                    $sheet->mergeCells('N'.$row.':N'.($row + ($dato->rowspan - 1)));
                    $spreadsheet->getActiveSheet()->getStyle('N'.$row.':N'.($row + ($dato->rowspan - 1)))->applyFromArray($styleArray);
                    $sheet->mergeCells('O'.$row.':O'.($row + ($dato->rowspan - 1)));
                    $spreadsheet->getActiveSheet()->getStyle('O'.$row.':O'.($row + ($dato->rowspan - 1)))->applyFromArray($styleArray);
                    $sheet->mergeCells('P'.$row.':P'.($row + ($dato->rowspan - 1)));
                    $spreadsheet->getActiveSheet()->getStyle('P'.$row.':P'.($row + ($dato->rowspan - 1)))->applyFromArray($styleArray);
                    $sheet->mergeCells('Q'.$row.':Q'.($row + ($dato->rowspan - 1)));
                    $spreadsheet->getActiveSheet()->getStyle('Q'.$row.':R'.($row + ($dato->rowspan - 1)))->applyFromArray($styleArray);
                    $sheet->mergeCells('R'.$row.':R'.($row + ($dato->rowspan - 1)));
                    $spreadsheet->getActiveSheet()->getStyle('R'.$row.':Q'.($row + ($dato->rowspan - 1)))->applyFromArray($styleArray);
                    $sheet->mergeCells('S'.$row.':S'.($row + ($dato->rowspan - 1)));
                    $spreadsheet->getActiveSheet()->getStyle('S'.$row.':S'.($row + ($dato->rowspan - 1)))->applyFromArray($styleArray);            
                }
                $row++;
            }
        }else{
            $row = 8;
            $index = 0;
            $orden = 1;

            foreach ($datos as $dato) {
                if($dato->rol == "Jefe de hogar"){
                    $sheet->getRowDimension($row)->setRowHeight(30);
                    $sheet->setCellValue('A'.$row, $dato->numero_familia);
                    $spreadsheet->getActiveSheet()->getStyle('A'.$row)->applyFromArray($styleArray);
                
                    if($dato->rol == "Jefe de hogar"){
                        $orden = 1;
                    }else{
                        $orden++;
                    }

                    $sheet->setCellValue('B'.$row, $orden);
                    $spreadsheet->getActiveSheet()->getStyle('B'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('C'.$row, $dato->nombre_completo);
                    $spreadsheet->getActiveSheet()->getStyle('C'.$row)->applyFromArray($styleArray);
                    $sheet->getStyle('C'.$row)->getAlignment()->setWrapText(true);
    
                    $sheet->setCellValue('D'.$row, $dato->tipo_identificacion);
                    $spreadsheet->getActiveSheet()->getStyle('D'.$row)->applyFromArray($styleArray);
    
                    $sheet->setCellValue('E'.$row, $dato->identificacion);
                    $spreadsheet->getActiveSheet()->getStyle('E'.$row)->applyFromArray($styleArray);
                    $spreadsheet->getActiveSheet()->getStyle('E'.$row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
                
                    $sheet->setCellValue('F'.$row, $dato->direccion);
                    $spreadsheet->getActiveSheet()->getStyle('F'.$row)->applyFromArray($styleArray);    

                    $sheet->setCellValue('G'.$row, $dato->sexo);
                    $spreadsheet->getActiveSheet()->getStyle('G'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('H'.$row, self::calcularEdad($dato->fecha_nacimiento)." Años");
                    $spreadsheet->getActiveSheet()->getStyle('H'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('I'.$row, $dato->estado_civil);
                    $spreadsheet->getActiveSheet()->getStyle('I'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('J'.$row, $dato->rol);
                    $spreadsheet->getActiveSheet()->getStyle('J'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('K'.$row, $dato->ocupacion);
                    $spreadsheet->getActiveSheet()->getStyle('K'.$row)->applyFromArray($styleArray);
                    $sheet->getStyle('K'.$row)->getAlignment()->setWrapText(true);

                    $sheet->setCellValue('L'.$row, $dato->escolaridad);
                    $spreadsheet->getActiveSheet()->getStyle('L'.$row)->applyFromArray($styleArray);
                    $sheet->getStyle('L'.$row)->getAlignment()->setWrapText(true);

                    $sheet->setCellValue('M'.$row, $dato->tipo_vivienda);
                    $spreadsheet->getActiveSheet()->getStyle('M'.$row)->applyFromArray($styleArray);
            
                    $sheet->setCellValue('N'.$row, $dato->tenencia);
                    $spreadsheet->getActiveSheet()->getStyle('N'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('O'.$row, $dato->posecion_baldia == "" ? "N.A" : $dato->posecion_baldia );
                    $spreadsheet->getActiveSheet()->getStyle('O'.$row)->applyFromArray($styleArray);
                    
                    $sheet->setCellValue('P'.$row, $dato->propiedad_titulo == "" ? "N.A" : $dato->propiedad_titulo );
                    $spreadsheet->getActiveSheet()->getStyle('P'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('Q'.$row, $dato->area_total == "" ? "N.A" : $dato->area_total);
                    $spreadsheet->getActiveSheet()->getStyle('Q'.$row)->applyFromArray($styleArray);

                    $sheet->setCellValue('R'.$row, $dato->area_total == "" ? "N.A" : ((double)$dato->area_total / 10000));
                    $spreadsheet->getActiveSheet()->getStyle('R'.$row)->applyFromArray($styleArray);

                    if($dato->tipo_vivienda == "Finca" || $dato->tipo_vivienda == "Parcela"){
                        $sheet->setCellValue('S'.$row, $dato->area_destinada_ocupada);
                        $spreadsheet->getActiveSheet()->getStyle('S'.$row)->applyFromArray($styleArray);
                    }else{
                        $sheet->setCellValue('S'.$row, "N.A");
                        $spreadsheet->getActiveSheet()->getStyle('S'.$row)->applyFromArray($styleArray);
                    }
                    
                    $row++; 
                }
            }
        }
        
        $writer = new Xlsx($spreadsheet);
        $filePath = 'reportes/' . $nombre;
        $writer->save($filePath);

        $respuesta = [
            'nombre' => '/reportes/' . $nombre,
        ];

        return response()->json($respuesta, 200);
    }

    public function calcularEdad($fechaNacimiento){
        $fechaNacimiento = Carbon::parse($fechaNacimiento);
        $fechaActual = Carbon::now();
        $edad = $fechaNacimiento->diffInYears($fechaActual);
        return $edad;
    }

    public function consultarIntegrantesJefes($idCorregimiento){
        if($idCorregimiento == "todo"){
            $jefes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->join("vivienda_hogar", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
            ->where("rol", "Jefe de hogar")
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
    
            $integrantes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->where("rol", "<>", "Jefe de hogar")
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
        }else{
            $jefes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->join("vivienda_hogar", "informacion_personal.identificacion", "vivienda_hogar.identificacion_jefe")
            ->where("rol", "Jefe de hogar")
            ->where("informacion_personal.corregimiento", $idCorregimiento)
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
    
            $integrantes = DB::connection('mysql')->table('informacion_personal')
            ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
            ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
            ->where("rol", "<>", "Jefe de hogar")
            ->where("informacion_personal.corregimiento", $idCorregimiento)
            ->orderBy("informacion_personal.numero_caracterizacion")
            ->get();
        }

        $individuos = [];

        $numero_familia = 1;
        foreach ($jefes as $jefe) {
            
            $rowspan = 1;
            $numero_orden = 1;

            foreach ($integrantes as $key => $integrante) {
                if ($integrante->id_jefe == $jefe->identificacion) {
                    $rowspan++;
                }
            }

            $jefe->numero_familia = $numero_familia;
            $jefe->rowspan = $rowspan;
            $jefe->numero_orden = $numero_orden;
            $individuos[] = $jefe;

            $numero_orden++;

            if($jefe->tipo_vivienda == "Finca" || $jefe->tipo_vivienda == "Parcela"){
                $jefe->area_destinada_ocupada =  DB::connection('mysql')->table('actividades_vivienda_hogar')
                ->where("identificacion_jefe", $jefe->identificacion)
                ->sum("area_destinada");
            }

            $jefe->escolaridad = DB::connection('mysql')->table('escolaridad')
            ->where("id", $jefe->nivel_educativo)
            ->first()->descripcion;


            foreach ($integrantes as $key => $integrante) {
                if ($integrante->id_jefe == $jefe->identificacion) {
                    $integrante->escolaridad = DB::connection('mysql')->table('escolaridad')
                    ->where("id", $integrante->nivel_educativo)
                    ->first()->descripcion;
                    $integrante->numero_familia = $numero_familia;
                    $integrante->numero_orden = $numero_orden;
                    $numero_orden++;
                    $individuos[] = $integrante;
                    unset($integrantes[$key]);
                    $rowspan++;
                }
            }

            $numero_familia += 1;
        }

        return $individuos;
    }

    public function consolidadoCensadosConcejo(Request $request){
        $individuos = [];
        $individuos2 = [];

        $concejo = $request->input('concejo');

        $pagina = $request->input('pagina');
        $limit = 10;

        if ($pagina == null || $pagina == "") {
            $pagina = 1;
        }

        if ($pagina == 1) {
            $offset = 0;
            $pagina--;
        } else {
            $pagina--;
            $offset = $pagina * $limit;
        }

     
        $individuos = DB::connection('mysql')->table('informacion_personal')
        ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
        ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
        ->join("cultura_tradiciones", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("cultura_tradiciones.concejo", $concejo)
        ->offset($offset)
        ->limit($limit) 
        ->orderBy("informacion_personal.numero_caracterizacion")
        ->get();

        
        foreach ($individuos as $individuo) {
            if($individuo->rol == "Jefe de hogar"){
                $individuo->vivienda_hogar = DB::connection('mysql')->table('vivienda_hogar')
                ->where("identificacion_jefe", $individuo->identificacion)
                ->first();

                if($individuo->vivienda_hogar->tipo_vivienda == "Finca" || $individuo->vivienda_hogar->tipo_vivienda == "Parcela"){
                    $individuo->area_destinada_ocupada =  DB::connection('mysql')->table('actividades_vivienda_hogar')
                    ->where("identificacion_jefe", $individuo->identificacion)
                    ->sum("area_destinada");
                }
            }
           
            $individuo->escolaridad = DB::connection('mysql')->table('escolaridad')
            ->where("id", $individuo->nivel_educativo)
            ->first()->descripcion;

            $individuo->edad = self::calcularEdad($individuo->fecha_nacimiento);
        }

        $numero_familia = 1;
        foreach ($individuos as $key) {
            $bandera = false;
            foreach($individuos2 as $key2){
                if($key->id_jefe == $key2->identificacion || ($key->id_jefe == $key2->id_jefe && $key->id_jefe != null)){
                    $key->numero_familia = $key2->numero_familia;
                    $bandera = true;
                    break;
                }
            }

            if($bandera == false){
                $key->numero_familia = $numero_familia;
                $numero_familia++;
            }

            $individuos2[] = $key;
        }

        usort($individuos2, function ($a, $b) {
            return $a->numero_familia - $b->numero_familia;
        });

        return response()->json($individuos2);
    }

    function compareByNumeroFamilia($a, $b) {
        return $a->numero_familia - $b->numero_familia;
    }

    public function paginacionConsolidadoConcejo(Request $request){
       
        $concejo = $request->input('concejo');
        $pagina = $request->input('pagina');

        $individuos = DB::connection('mysql')->table('informacion_personal')
        ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
        ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
        ->join("cultura_tradiciones", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("cultura_tradiciones.concejo", $concejo)
        ->count();

        $paginas = ceil($individuos / 10);
        
        $respuesta = [
            'paginas_consulta' => $paginas,
            'pagina_actual' => (int)$pagina,
            'numero_registros' => $individuos,
        ];
        
        return response()->json($respuesta, 200);
    }

    public function consultarIntegrantesJefesConcejo($concejo){
        $individuos = [];
        $individuos2 = [];
     
        $individuos = DB::connection('mysql')->table('informacion_personal')
        ->join("educacion", "informacion_personal.identificacion", "educacion.identificacion_individuo")
        ->join("situacion_laboral", "informacion_personal.identificacion", "situacion_laboral.identificacion_individuo")
        ->join("cultura_tradiciones", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("cultura_tradiciones.concejo", $concejo)
        ->orderBy("informacion_personal.numero_caracterizacion")
        ->get();


        foreach ($individuos as $individuo) {
            if($individuo->rol == "Jefe de hogar"){
                $individuo->vivienda_hogar = DB::connection('mysql')->table('vivienda_hogar')
                ->where("identificacion_jefe", $individuo->identificacion)
                ->first();

                if($individuo->vivienda_hogar->tipo_vivienda == "Finca" || $individuo->vivienda_hogar->tipo_vivienda == "Parcela"){
                    $individuo->area_destinada_ocupada =  DB::connection('mysql')->table('actividades_vivienda_hogar')
                    ->where("identificacion_jefe", $individuo->identificacion)
                    ->sum("area_destinada");
                }
            }
           
            $individuo->escolaridad = DB::connection('mysql')->table('escolaridad')
            ->where("id", $individuo->nivel_educativo)
            ->first()->descripcion;

            $individuo->edad = self::calcularEdad($individuo->fecha_nacimiento);
        }

        $numero_familia = 1;
        foreach ($individuos as $key) {
            $bandera = false;
            foreach($individuos2 as $key2){
                if($key->id_jefe == $key2->identificacion || ($key->id_jefe == $key2->id_jefe && $key->id_jefe != null)){
                    $key->numero_familia = $key2->numero_familia;
                    $bandera = true;
                    break;
                }
            }

            if($bandera == false){
                $key->numero_familia = $numero_familia;
                $numero_familia++;
            }

            $individuos2[] = $key;
        }

        usort($individuos2, function ($a, $b) {
            return $a->numero_familia - $b->numero_familia;
        });

        return $individuos2;
    }

    public function exportarConsolidadoExcelConcejo(Request $request) {
        $concejo = $request->input('concejo');
        $tipo_reporte = $request->input('tipo_reporte');

        $datos = self::consultarIntegrantesJefesConcejo($concejo);

        $styleArray = [
            'font' => [
                'bold' => true,
                'size' => 10
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
        ];

        $styleArray2 = [
            'font' => [
                'bold' => true,
                'size' => 10,
            ],
            'alignment' => [
                'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
                'vertical' => \PhpOffice\PhpSpreadsheet\Style\Alignment::VERTICAL_CENTER,
                'indent' => 1,
            ],
            'borders' => [
                'allBorders' => [
                    'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_THIN,
                    'color' => ['rgb' => '000000'],
                ],
            ],
            'padding' => [
                'left' => 20,
            ],
        ];
    
    
        $titulo = "reporte_caracterizacion_concejo";

        $nombre = $titulo.".xlsx";
       
        $spreadsheet = new Spreadsheet();
            
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Consolidado General');

        $spreadsheet->getActiveSheet()->getColumnDimension('A')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('B')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('C')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('D')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('E')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('F')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('G')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('H')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('I')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('J')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('M')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('N')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('O')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('P')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('Q')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('R')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('S')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('T')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('U')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('V')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('W')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('X')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('Y')->setAutoSize(true);
        $spreadsheet->getActiveSheet()->getColumnDimension('Z')->setAutoSize(true);
       

        $sheet->mergeCells('A1:C2');
        $sheet->getRowDimension(1)->setRowHeight(60); 
        $sheet->getRowDimension(2)->setRowHeight(30);
        $sheet->getRowDimension(3)->setRowHeight(30);
        $sheet->getRowDimension(4)->setRowHeight(30);
        $sheet->getRowDimension(5)->setRowHeight(30);
        $sheet->getRowDimension(6)->setRowHeight(20);

        $spreadsheet->getActiveSheet()->getStyle('A1:C2')->applyFromArray($styleArray);

        $sheet->setCellValue('D1', 'MACROPROCESO: AGENCIA NACIONAL DE TIERRAS '. PHP_EOL . 'PROCESO: DIRECCIÓN DE ASUNTOS ÉTNICOS');
        $sheet->mergeCells('D1:O1');
        $spreadsheet->getActiveSheet()->getStyle('D1:O1')->applyFromArray($styleArray);
        $style = $sheet->getStyle('D1:O1');
        $alignment = $style->getAlignment();
        $alignment->setWrapText(true);

        $sheet->setCellValue('D2', 'FORMATO: CENSO POBLACIÓN COMUNIDADES NEGRAS');
        $sheet->mergeCells('D2:O2');
        $spreadsheet->getActiveSheet()->getStyle('D2:O2')->applyFromArray($styleArray);
       
        $sheet->setCellValue('P1', 'CÓDIGO: '. PHP_EOL . 'FECHA:');
        $sheet->mergeCells('P1:S1');
        $spreadsheet->getActiveSheet()->getStyle('P1:S1')->applyFromArray($styleArray);
        $style = $sheet->getStyle('P1:S1');
        $alignment = $style->getAlignment();
        $alignment->setWrapText(true);

        $sheet->setCellValue('P2', 'PAGINA: 1/1');
        $sheet->mergeCells('P2:S2');
        $spreadsheet->getActiveSheet()->getStyle('P2:S2')->applyFromArray($styleArray);

        $sheet->setCellValue('A3', 'DATOS DE IDENTIFICACIÓN');
        $sheet->mergeCells('A3:S3');
        $spreadsheet->getActiveSheet()->getStyle('A3:S3')->applyFromArray($styleArray);

        $sheet->setCellValue('A4', '1. Departamento: CESAR');
        $sheet->mergeCells('A4:C4');
        $spreadsheet->getActiveSheet()->getStyle('A4:C4')->applyFromArray($styleArray2);

        $sheet->setCellValue('D4', '2. Municipio: EL PASO');
        $sheet->mergeCells('D4:F4');
        $spreadsheet->getActiveSheet()->getStyle('D4:F4')->applyFromArray($styleArray2);

        $sheet->setCellValue('G4', '3. Consejo Comunitario:   '.$concejo);
        $sheet->mergeCells('G4:O4');
        $spreadsheet->getActiveSheet()->getStyle('G4:O4')->applyFromArray($styleArray2);

        $sheet->setCellValue('P4', '4. Fecha: 14 de Noviembre del 2023');
        $sheet->mergeCells('P4:S4');
        $spreadsheet->getActiveSheet()->getStyle('P4:S4')->applyFromArray($styleArray2);

        $sheet->setCellValue('A5', '5. Representante Legal: EUFROSINA VEGA MIELES');
        $sheet->mergeCells('A5:F5');
        $spreadsheet->getActiveSheet()->getStyle('A5:F5')->applyFromArray($styleArray2);

        $sheet->setCellValue('G5', '6. Consejo Comunitario: '. $concejo);
        $sheet->mergeCells('G5:S5');
        $spreadsheet->getActiveSheet()->getStyle('G5:S5')->applyFromArray($styleArray2);
        

        $sheet->setCellValue('A6', 'Listado de Personas');
        $sheet->mergeCells('A6:B6');
        $spreadsheet->getActiveSheet()->getStyle('A6:B6')->applyFromArray($styleArray);
        $sheet->getStyle('A6:B6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A6:B6')->getFill()->getStartColor()->setARGB('8cee8c'); 
        $sheet->getStyle('A6')->getAlignment()->setWrapText(true);

        $sheet->setCellValue('C6', 'Datos');
        $sheet->mergeCells('C6:F6');
        $spreadsheet->getActiveSheet()->getStyle('C6:F6')->applyFromArray($styleArray);
        $sheet->getStyle('C6:F6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('C6:F6')->getFill()->getStartColor()->setARGB('43d7f0'); 

        $sheet->setCellValue('G6', '');
        $sheet->mergeCells('G6:L6');
        $spreadsheet->getActiveSheet()->getStyle('G6:L6')->applyFromArray($styleArray);
        $sheet->getStyle('G6:L6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('G6:L6')->getFill()->getStartColor()->setARGB('cdf043'); 

        $sheet->setCellValue('M6', '');
        $sheet->mergeCells('M6:N6');
        $spreadsheet->getActiveSheet()->getStyle('M6:N6')->applyFromArray($styleArray);
        $sheet->getStyle('M6:N6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('M6:N6')->getFill()->getStartColor()->setARGB('43d7f0'); 

        $sheet->setCellValue('O6', 'Tenencia de Tierras');
        $sheet->mergeCells('O6:S6');
        $spreadsheet->getActiveSheet()->getStyle('O6:S6')->applyFromArray($styleArray);
        $sheet->getStyle('O6:S6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('O6:S6')->getFill()->getStartColor()->setARGB('8cee8c'); 

        $sheet->setCellValue('A7', 'Nº Familia');
        $sheet->mergeCells('A7:B7');
        $spreadsheet->getActiveSheet()->getStyle('A7:B7')->applyFromArray($styleArray);
        $sheet->getStyle('A7:B7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A7:B7')->getFill()->getStartColor()->setARGB('8cee8c'); 


        $sheet->setCellValue('C7', 'Nombre Completo');
        $sheet->setCellValue('D7', 'Tipo');
        $sheet->setCellValue('E7', 'Nº Documento');
        $sheet->setCellValue('F7', 'Dirección');
        $spreadsheet->getActiveSheet()->getStyle('C7:F7')->applyFromArray($styleArray);
        $sheet->getStyle('C6:F7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('C6:F7')->getFill()->getStartColor()->setARGB('43d7f0');

        $sheet->setCellValue('G7', 'Sexo');
        $sheet->setCellValue('H7', 'Edad');
        $sheet->setCellValue('I7', 'Estado Civil');
        $sheet->setCellValue('J7', 'Parentesco');
        $sheet->setCellValue('K7', 'Ocupación');
        $sheet->setCellValue('L7', 'Escolaridad');
        $spreadsheet->getActiveSheet()->getStyle('G7:L7')->applyFromArray($styleArray);
        $sheet->getStyle('G7:L7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('G7:L7')->getFill()->getStartColor()->setARGB('cdf043'); 


        $sheet->setCellValue('M7', 'Tipo '. PHP_EOL . ' Vivienda');
        $sheet->setCellValue('N7', 'Tenencia '. PHP_EOL . ' Vivienda');
        $spreadsheet->getActiveSheet()->getStyle('M7:N7')->applyFromArray($styleArray);
        $sheet->getStyle('M6:N7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('M6:N7')->getFill()->getStartColor()->setARGB('43d7f0');

        $sheet->setCellValue('O7', 'Posesión '. PHP_EOL . ' Baldíos');
        $sheet->setCellValue('P7', 'Propiedad '. PHP_EOL . ' con Título');
        $sheet->setCellValue('Q7', 'Area M2');
        $sheet->setCellValue('R7', 'Area Ha');
        $sheet->setCellValue('S7', 'Area Utilizada');
        $spreadsheet->getActiveSheet()->getStyle('O7:S7')->applyFromArray($styleArray);
        $sheet->getStyle('O7:S7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('O7:S7')->getFill()->getStartColor()->setARGB('8cee8c'); 


        $sheet->getStyle('O7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('O7')->getAlignment()->setTextRotation(90);
        $sheet->getStyle('M7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('M7')->getAlignment()->setTextRotation(90);
        $sheet->getStyle('N7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('N7')->getAlignment()->setTextRotation(90);
        $sheet->getStyle('P7')->getAlignment()->setWrapText(true);
        $sheet->getStyle('P7')->getAlignment()->setTextRotation(90);
        

        $sheet->getColumnDimension('A')->setWidth(200);
        $sheet->getColumnDimension('B')->setWidth(200);
        $sheet->getColumnDimension('K')->setWidth(20);
        $sheet->getColumnDimension('L')->setWidth(25);

        // Add an image to cell A1
        $imagePath = public_path('imagenes/censo.png'); // Use public_path to get the correct path
        $drawing = new Drawing();
        $drawing->setPath($imagePath);
        $drawing->setCoordinates('A1');
        $drawing->setOffsetX(10);
        $drawing->setOffsetY(33);
        $drawing->setWidth(100); 
        $drawing->setHeight(60);
        $drawing->setWorksheet($sheet);

        $row = 8;
        $orden = 1;

        foreach ($datos as $dato) {
            $sheet->getRowDimension($row)->setRowHeight(30);

            $sheet->setCellValue('A'.$row, $dato->numero_familia);
            $sheet->mergeCells('A'.$row.':B'.$row);
            $spreadsheet->getActiveSheet()->getStyle('A'.$row.':'.'B'.$row)->applyFromArray($styleArray);
           
            $sheet->setCellValue('C'.$row, $dato->nombre_completo);
            $spreadsheet->getActiveSheet()->getStyle('C'.$row)->applyFromArray($styleArray);
            $sheet->getStyle('C'.$row)->getAlignment()->setWrapText(true);

            $sheet->setCellValue('D'.$row, $dato->tipo_identificacion);
            $spreadsheet->getActiveSheet()->getStyle('D'.$row)->applyFromArray($styleArray);

            $sheet->setCellValue('E'.$row, $dato->identificacion);
            $spreadsheet->getActiveSheet()->getStyle('E'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('E'.$row)->getNumberFormat()->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_NUMBER);
           
            $sheet->setCellValue('F'.$row, $dato->direccion);
            $spreadsheet->getActiveSheet()->getStyle('F'.$row)->applyFromArray($styleArray);

            $sheet->setCellValue('G'.$row, $dato->sexo);
            $spreadsheet->getActiveSheet()->getStyle('G'.$row)->applyFromArray($styleArray);

            $sheet->setCellValue('H'.$row, self::calcularEdad($dato->fecha_nacimiento)." Años");
            $spreadsheet->getActiveSheet()->getStyle('H'.$row)->applyFromArray($styleArray);

            $sheet->setCellValue('I'.$row, $dato->estado_civil);
            $spreadsheet->getActiveSheet()->getStyle('I'.$row)->applyFromArray($styleArray);

            $sheet->setCellValue('J'.$row, $dato->rol);
            $spreadsheet->getActiveSheet()->getStyle('J'.$row)->applyFromArray($styleArray);

            $sheet->setCellValue('K'.$row, $dato->ocupacion);
            $spreadsheet->getActiveSheet()->getStyle('K'.$row)->applyFromArray($styleArray);
            $sheet->getStyle('K'.$row)->getAlignment()->setWrapText(true);

            $sheet->setCellValue('L'.$row, $dato->escolaridad);
            $spreadsheet->getActiveSheet()->getStyle('L'.$row)->applyFromArray($styleArray);
            $sheet->getStyle('L'.$row)->getAlignment()->setWrapText(true);

            if($dato->rol == "Jefe de hogar"){
                $sheet->setCellValue('M'.$row, $dato->vivienda_hogar->tipo_vivienda);
                $sheet->setCellValue('N'.$row, $dato->vivienda_hogar->tenencia);
                $sheet->setCellValue('O'.$row, $dato->vivienda_hogar->posecion_baldia == "" ? "N.A" : $dato->vivienda_hogar->posecion_baldia );
                $sheet->setCellValue('P'.$row, $dato->vivienda_hogar->propiedad_titulo == "" ? "N.A" : $dato->vivienda_hogar->propiedad_titulo );
                $sheet->setCellValue('Q'.$row, $dato->vivienda_hogar->area_total == "" ? "N.A" : $dato->vivienda_hogar->area_total);
                $sheet->setCellValue('R'.$row, $dato->vivienda_hogar->area_total == "" ? "N.A" : ((double)$dato->vivienda_hogar->area_total / 10000));
                if($dato->vivienda_hogar->tipo_vivienda == "Finca" || $dato->vivienda_hogar->tipo_vivienda == "Parcela"){
                    $sheet->setCellValue('S'.$row, $dato->area_destinada_ocupada);
                }else{
                    $sheet->setCellValue('S'.$row, "N.A");
                }
                
            }

            $spreadsheet->getActiveSheet()->getStyle('M'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('N'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('O'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('P'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('Q'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('R'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('S'.$row)->applyFromArray($styleArray);
            $spreadsheet->getActiveSheet()->getStyle('S'.$row)->applyFromArray($styleArray);


            $row++; 
        }

        
        $writer = new Xlsx($spreadsheet);
        $filePath = 'reportes/' . $nombre;
        $writer->save($filePath);

        $respuesta = [
            'nombre' => '/reportes/' . $nombre,
        ];

        return response()->json($respuesta, 200);
    }
}
