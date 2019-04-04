<?php

require './vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

$spreadsheet = new Spreadsheet();

$conexio = new mysqli("localhost", "root","root", "empresa");
$conexio->query("SET NAMES 'utf80");

$numDep = $_GET['numDep'];

$query = "SELECT d.codi, e.nom, d.ciutat, e.sou, e.comisio FROM departaments AS d INNER JOIN empleats AS e ON d.codi = e.ndepartament WHERE d.codi = ?";

$consulta = $conexio->prepare($query);
$consulta->bind_param("i", $numDep);
if($consulta->execute()){
    echo "OK";
    $result = $consulta->get_result();
} else {
    echo "ERROR SQL";
}

// TITLES
$spreadsheet->setActiveSheetIndex(0)
    ->setCellValue('B2', 'DADES DEPARTAMENT PER CODI - GitHub')
    ->setCellValue('B3', 'Codi')
    ->setCellValue('C3', 'Nom')
    ->setCellValue('D3', 'Ciutat')
    ->setCellValue('E3', 'Sou')
    ->setCellValue('F3', 'Comissió')
    ->setCellValue('G3', 'Suma Total');
    

// INFO RESULTS
$index = 4;
while ($dades = $result->fetch_array()) {
    if($dades['comisio'] == ""){
        $dades['comisio'] = 0;
    }
    $spreadsheet->setActiveSheetIndex(0)
        ->setCellValue('B'.$index, $dades['codi'])
        ->setCellValue('C'.$index, $dades['nom'])
        ->setCellValue('D'.$index, $dades['ciutat'])
        ->setCellValue('E'.$index, $dades['sou'])
        ->setCellValue('F'.$index, $dades['comisio'])
        ->setCellValue('G'.$index,"=SUM(E".$index.":F".$index.")");
    
    $spreadsheet->getActiveSheet()
        ->getStyle('E'.$index.':F'.$index)	
        ->getNumberFormat()
        //->setFormatCode(\PhpOffice\PhpSpreadsheet\Style\NumberFormat::FORMAT_CURRENCY_EUR_SIMPLE);
        ->setFormatCode('#,##0.00 €');

        $index++;
}

/* STYLES */
$tableFormat = array(
    'font' => [
        'bold' => true,
        'name' => 'Arial',
        'color' => ['argb' => '0000'],
        'size' => 10
    ],
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
            'color' => ['argb' => '3232ff'],
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID,
        'rotation' => 90,
        'color' => ['argb' => '55ffee'],
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_LEFT,
    ],
);    

$styleArray = array(
    'font' => [
        'bold' => true,
        'name' => 'Verdana',
        'color' => ['argb' => '0000ff'],
        'size' => 12
    ],
    'borders' => [
        'allBorders' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
            'color' => ['argb' => '0000'],
        ],
    ],
    'fill' => [
        'fillType' => \PhpOffice\PhpSpreadsheet\Style\Fill::FILL_GRADIENT_LINEAR,
        'rotation' => 90,
        'startColor' => [
            'argb' => 'c5ffbc',
        ],
        'endColor' => [
            'argb' => 'e5e5ff',
        ],
    ],
    'alignment' => [
        'horizontal' => \PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER,
    ],
);

$spreadsheet->getActiveSheet()->getColumnDimension('B')->SetAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('C')->SetAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('D')->SetAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('E')->SetAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('F')->SetAutoSize(true);
$spreadsheet->getActiveSheet()->getColumnDimension('G')->SetAutoSize(true);
$spreadsheet->getActiveSheet()->getStyle('B:G')->getAlignment()->setHorizontal('center');

$spreadsheet->getActiveSheet()->getStyle('B2:G3')->applyFromArray($styleArray);
$index2 = $index - 1;
$spreadsheet->getActiveSheet()->getStyle('B4:G'.$index2)->applyFromArray($tableFormat);
$spreadsheet->getActiveSheet()->getStyle('B2:F3')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$spreadsheet->getActiveSheet()->mergeCells('B2:G2');

// LOGO GITHUB
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
$drawing->setName('Logo');
$drawing->setDescription('Logo');
$drawing->setPath('./images/github.png');
$drawing->setHeight(30);
$drawing->setCoordinates('H3');
$drawing->setOffsetX(20);
$drawing->setRotation(25);
$drawing->getShadow()->setVisible(true);
$drawing->getShadow()->setDirection(45);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

$spreadsheet->getActiveSheet()->setTitle('Pàgina 1');
$writer = new Xlsx($spreadsheet);
$writer->save('practica2.xlsx');

/*$filename = 'practica2.xlsx';
// Redirect output to a client's web browser (Xlsx)
header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename="'.$filename.'"');
header('Cache-Control: max-age=0');
// If you're serving to IE 9, then the following may be needed
header('Cache-Control: max-age=1');
 
// If you're serving to IE over SSL, then the following may be needed
header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
header('Pragma: public'); // HTTP/1.
$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save('php://output');*/
?>