<?php

namespace App\Http\Controllers;

use Illuminate\Http\Request;

namespace App\Http\Controllers;
use Illuminate\Http\Request;
use DB;
use File;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Worksheet\Drawing;
use Carbon\Carbon;

use Illuminate\Support\Facades\Session;


class ExportarExcelController extends Controller
{
    public function exportarDatosExcel(Request $request) {

        $tipoExcel = $request->input('tipo');

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
    
    
        $nombre = "informe_caracterizacion_afro.xlsx";

        
        $spreadsheet = new Spreadsheet();
            
        $sheet = $spreadsheet->getActiveSheet();
        $sheet->setTitle('Listado de personas');

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
       
        $sheet->setCellValue('P1', 'CÓDIGO: ');
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

        $sheet->setCellValue('G4', '3. Corregimiento o vereda: Todos');
        $sheet->mergeCells('G4:O4');
        $spreadsheet->getActiveSheet()->getStyle('G4:O4')->applyFromArray($styleArray2);

        $sheet->setCellValue('P4', '4. Fecha: '.date("d/m/Y - H:m:s"));
        $sheet->mergeCells('P4:S4');
        $spreadsheet->getActiveSheet()->getStyle('P4:S4')->applyFromArray($styleArray2);

        $sheet->setCellValue('A5', '5. Representante Legal: ');
        $sheet->mergeCells('A5:F5');
        $spreadsheet->getActiveSheet()->getStyle('A5:F5')->applyFromArray($styleArray2);

        $sheet->setCellValue('G5', '6. Consejo Comunitario: Todos');
        $sheet->mergeCells('G5:S5');
        $spreadsheet->getActiveSheet()->getStyle('G5:S5')->applyFromArray($styleArray2);
        

        $sheet->setCellValue('A6', 'Listado de Personas');
        $sheet->mergeCells('A6:S6');
        $spreadsheet->getActiveSheet()->getStyle('A6:S6')->applyFromArray($styleArray);
        $sheet->getStyle('A6:S6')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A6:S6')->getFill()->getStartColor()->setARGB('8cee8c'); 
        $sheet->getStyle('A6')->getAlignment()->setWrapText(true);

        $sheet->setCellValue('A7', 'Tipo de Documento');
        $sheet->setCellValue('B7', '# Documento');
        $sheet->setCellValue('C7', 'Nombre');
        $sheet->setCellValue('I7', 'Dirección');
        $sheet->setCellValue('N7', 'Sexo');
        $sheet->setCellValue('O7', 'Edad');
      
        $spreadsheet->getActiveSheet()->getStyle('A7')->applyFromArray($styleArray);
        $sheet->getStyle('A7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('A7')->getFill()->getStartColor()->setARGB('8cee8c');

        $spreadsheet->getActiveSheet()->getStyle('B7')->applyFromArray($styleArray);
        $sheet->getStyle('B7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('B7')->getFill()->getStartColor()->setARGB('8cee8c');
        
        $sheet->mergeCells('C7:H7');
        $spreadsheet->getActiveSheet()->getStyle('C7:H7')->applyFromArray($styleArray);
        $sheet->getStyle('C7:H7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('C7:H7')->getFill()->getStartColor()->setARGB('8cee8c');

        $sheet->mergeCells('I7:M7');
        $spreadsheet->getActiveSheet()->getStyle('I7:M7')->applyFromArray($styleArray);
        $sheet->getStyle('I7:M7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('I7:M7')->getFill()->getStartColor()->setARGB('8cee8c');

        $spreadsheet->getActiveSheet()->getStyle('N7')->applyFromArray($styleArray);
        $sheet->getStyle('N7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('N7')->getFill()->getStartColor()->setARGB('8cee8c');
       
        $spreadsheet->getActiveSheet()->getStyle('O7')->applyFromArray($styleArray);
        $sheet->getStyle('O7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
        $sheet->getStyle('O7')->getFill()->getStartColor()->setARGB('8cee8c');


        $datos = [];
        switch ($tipoExcel) {
            case 'e11':
                $sheet->setCellValue('P7', 'Concejo Comunitario');
                $datos = self::PoblacionPorConcejo();
                break;
            case 'e12':
                $sheet->setCellValue('P7', 'Etnia');
                $datos = self::PoblacionPorEtnia();
                break;
            case 'e13':
                $sheet->setCellValue('P7', 'Practica Cultural y religiosa');
                $datos = self::PoblacionPorPracticaCultural();
                break;
            case 'e14':
                $sheet->setCellValue('P7', '¿Habla lengua Afro?');
                $datos = self::PoblacionHablaAfro();
                break;
            case 'e21':
                $sheet->setCellValue('P7', 'GRUPO DE EDAD');
                $datos = self::PoblacionPorGrupoEdad();
                break;
            case 'e22':
                $sheet->setCellValue('N7', 'Edad');
                $sheet->setCellValue('O7', 'Sexo');

                $datos = self::PoblacionPorConcejo();
                break;
            case 'e23':
                $sheet->setCellValue('P7', 'Desplazado');
                $datos = self::PoblacionDesplazados();
                break;
            case 'e24':
                $sheet->setCellValue('P7', 'Estado Civil');
                $datos = self::PoblacionPorEstadoCivil();
                break;
            default:
                # code...
                break;
        }
       
       
        
        $i = 8;
        if($tipoExcel == "e22"){

            $sheet->mergeCells('O7:S7');
            $spreadsheet->getActiveSheet()->getStyle('O7:S7')->applyFromArray($styleArray);
            $sheet->getStyle('O7:S7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
            $sheet->getStyle('O7:S7')->getFill()->getStartColor()->setARGB('8cee8c');

            foreach ($datos as $item) {
                $sheet->setCellValue('A'.$i, $item->tipo_identificacion);
                $sheet->setCellValue('B'.$i, $item->identificacion);
                $sheet->setCellValue('C'.$i, $item->nombre_completo);
                $sheet->setCellValue('I'.$i, $item->direccion);
                $sheet->setCellValue('N'.$i, $item->edad." Años");
                $sheet->setCellValue('O'.$i, $item->sexo);
              
                $spreadsheet->getActiveSheet()->getStyle('A'.$i)->applyFromArray($styleArray);
        
                $spreadsheet->getActiveSheet()->getStyle('B'.$i)->applyFromArray($styleArray);
                $spreadsheet->getActiveSheet()->getStyle('B'.$i)->getNumberFormat()->setFormatCode('0');
    
                $sheet->mergeCells('C'.$i.':H'.$i);
                $spreadsheet->getActiveSheet()->getStyle('C'.$i.':H'.$i)->applyFromArray($styleArray);
        
                $sheet->mergeCells('I'.$i.':M'.$i);
                $spreadsheet->getActiveSheet()->getStyle('I'.$i.':M'.$i)->applyFromArray($styleArray);
        
                $spreadsheet->getActiveSheet()->getStyle('N'.$i)->applyFromArray($styleArray);
                               
    
                $sheet->mergeCells('O'.$i.':S'.$i);
                $spreadsheet->getActiveSheet()->getStyle('O'.$i.':S'.$i)->applyFromArray($styleArray);
                $style = $sheet->getStyle('O'.$i.':S'.$i);
                $alignment = $style->getAlignment();
                $alignment->setWrapText(true);
    
                $sheet->getStyle('B'.$i)->getNumberFormat()->setFormatCode('0');
                $sheet->getRowDimension($i)->setRowHeight(35);
    
                $i++;
            }
        }else{

            $sheet->mergeCells('P7:S7');
            $spreadsheet->getActiveSheet()->getStyle('P7:S7')->applyFromArray($styleArray);
            $sheet->getStyle('P7:S7')->getFill()->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID);
            $sheet->getStyle('P7:S7')->getFill()->getStartColor()->setARGB('8cee8c');

            foreach ($datos as $item) {
                $sheet->setCellValue('A'.$i, $item->tipo_identificacion);
                $sheet->setCellValue('B'.$i, $item->identificacion);
                $sheet->setCellValue('C'.$i, $item->nombre_completo);
                $sheet->setCellValue('I'.$i, $item->direccion);
                $sheet->setCellValue('N'.$i, $item->sexo);
                $sheet->setCellValue('O'.$i, $item->edad. " Años");
                $sheet->setCellValue('P'.$i, $item->variable);
              
                $spreadsheet->getActiveSheet()->getStyle('A'.$i)->applyFromArray($styleArray);
        
                $spreadsheet->getActiveSheet()->getStyle('B'.$i)->applyFromArray($styleArray);
                $spreadsheet->getActiveSheet()->getStyle('B'.$i)->getNumberFormat()->setFormatCode('0');
    
                $sheet->mergeCells('C'.$i.':H'.$i);
                $spreadsheet->getActiveSheet()->getStyle('C'.$i.':H'.$i)->applyFromArray($styleArray);
        
                $sheet->mergeCells('I'.$i.':M'.$i);
                $spreadsheet->getActiveSheet()->getStyle('I'.$i.':M'.$i)->applyFromArray($styleArray);
        
                $spreadsheet->getActiveSheet()->getStyle('N'.$i)->applyFromArray($styleArray);
               
                $spreadsheet->getActiveSheet()->getStyle('O'.$i)->applyFromArray($styleArray);
                
    
                $sheet->mergeCells('P'.$i.':S'.$i);
                $spreadsheet->getActiveSheet()->getStyle('P'.$i.':S'.$i)->applyFromArray($styleArray);
                $style = $sheet->getStyle('P'.$i.':S'.$i);
                $alignment = $style->getAlignment();
                $alignment->setWrapText(true);
    
                $sheet->getStyle('B'.$i)->getNumberFormat()->setFormatCode('0');
                $sheet->getRowDimension($i)->setRowHeight(35);
    
                $i++;
            }
        }
        
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

        $writer = new Xlsx($spreadsheet);
        $filePath = 'reportes-excel/' . $nombre;
        $writer->save($filePath);

        $respuesta = [
            'nombre' => '/reportes-excel/' . $nombre,
        ];

        return response()->json($respuesta, 200);
    }

    public function calcularEdad($fechaNacimiento){
        $fechaNacimiento = Carbon::parse($fechaNacimiento);
        $fechaActual = Carbon::now();
        $edad = $fechaNacimiento->diffInYears($fechaActual);
        return $edad;
    }

    public function PoblacionPorConcejo(){
        $por_concejo = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->select("informacion_personal.*", "cultura_tradiciones.concejo as variable")
        ->orderBy("informacion_personal.direccion")
        ->get();

        foreach ($por_concejo as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
        }

        return $por_concejo;
    }

    public function PoblacionPorEtnia(){
        $por_etnia = DB::connection('mysql')->table('origen_etnia')
        ->join("informacion_personal", "informacion_personal.identificacion", "origen_etnia.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->select("informacion_personal.*", "origen_etnia.etnia as variable")
        ->orderBy("informacion_personal.direccion")
        ->get();

        foreach ($por_etnia as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
        }

        return $por_etnia;
    }

    public function PoblacionPorPracticaCultural(){
        $por_practicas = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->select("informacion_personal.*", "cultura_tradiciones.practicas_religiosas as variable")
        ->orderBy("informacion_personal.direccion")
        ->get();

        foreach ($por_practicas as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
        }

        return $por_practicas;
    }

    public function PoblacionHablaAfro(){
        $habla_lenguas = DB::connection('mysql')->table('cultura_tradiciones')
        ->join("informacion_personal", "informacion_personal.identificacion", "cultura_tradiciones.identificacion_individuo")
        ->where("informacion_personal.estado", "1")
        ->select("informacion_personal.*", "cultura_tradiciones.habla_lengua as variable")
        ->orderBy("informacion_personal.direccion")
        ->get();

        foreach ($habla_lenguas as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
        }

        return $habla_lenguas;
    }

    public function PoblacionPorGrupoEdad(){
        $por_edad = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", "1")
        ->select("informacion_personal.*")
        ->get();

        foreach ($por_edad as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);

            switch (true) {
                case ($item->edad >= 0 && $item->edad <= 4):
                    $item->variable = "De 0 a 4 Años";
                    break;
                case ($item->edad >= 5 && $item->edad <= 9):
                    $item->variable = "De 5 a 9 Años";
                    break;
                case ($item->edad >= 10 && $item->edad <= 14):
                    $item->variable = "De 10 a 14 Años";
                    break;
                case ($item->edad >= 15 && $item->edad <= 19):
                    $item->variable = "De 15 a 19 Años";
                    break;
                case ($item->edad >= 20 && $item->edad <= 24):
                    $item->variable = "De 20 a 24 Años";
                    break;
                case ($item->edad >= 25 && $item->edad <= 29):
                    $item->variable = "De 25 a 29 Años";
                    break;
                case ($item->edad >= 30 && $item->edad <= 34):
                    $item->variable = "De 30 a 34 Años";
                    break;
                case ($item->edad >= 35 && $item->edad <= 39):
                    $item->variable = "De 35 a 39 Años";
                    break;
                case ($item->edad >= 40 && $item->edad <= 44):
                    $item->variable = "De 40 a 44 Años";
                    break;
                case ($item->edad >= 45 && $item->edad <= 49):
                    $item->variable = "De 45 a 49 Años";
                    break;
                case ($item->edad >= 50 && $item->edad <= 54):
                    $item->variable = "De 50 a 54 Años";
                    break;
                case ($item->edad >= 55 && $item->edad <= 59):
                    $item->variable = "De 55 a 59 Años";
                    break;
                case ($item->edad >= 60 && $item->edad <= 64):
                    $item->variable = "De 60 a 64 Años";
                    break;
                case ($item->edad >= 65 && $item->edad <= 69):
                    $item->variable = "De 65 a 69 Años";
                    break;
                case ($item->edad >= 70 && $item->edad <= 74):
                    $item->variable = "De 70 a 74 Años";
                    break;
                case ($item->edad >= 75):
                    $item->variable = "Mayores de 75 Años";
                    break;
            }
        }

        $por_edad = $por_edad->sortBy("edad");
        return $por_edad;
    }

    public function PoblacionDesplazados(){
        $desplazados = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", "1")
        ->where("informacion_personal.desplazado", "Si")
        ->select("informacion_personal.*", "informacion_personal.desplazado as variable")
        ->get();

        foreach ($desplazados as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
        }

        return $desplazados; 
    }

    public function PoblacionPorEstadoCivil(){
        $estado_civil = DB::connection('mysql')->table('informacion_personal')
        ->where("informacion_personal.estado", "1")
        ->select("informacion_personal.*", "informacion_personal.estado_civil as variable")
        ->get();

        foreach ($estado_civil as $item) {
            $item->edad = self::calcularEdad($item->fecha_nacimiento);
        }

        return $estado_civil; 
    }
}
