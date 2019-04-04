<?php

require './vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\IOFactory;

$spreadsheet = new Spreadsheet();

$conexio = new mysqli("localhost", "root","root", "empresa");
$conexio->query("SET NAMES 'utf80");

$query = "select * from departaments";

$consulta = $conexio->prepare($query);
if($consulta->execute()){
    echo "OK";
    $result = $consulta->get_result();
} else {
    echo "ERROR SQL";
}

// TITLES
$spreadsheet->setActiveSheetIndex(0)
    ->setCellValue('B2', 'DADES DEPARTAMENT - GitHub')
    ->setCellValue('B3', 'Codi')
    ->setCellValue('C3', 'Nom')
    ->setCellValue('D3', 'Ciutat');

// INFO RESULTS
$index = 4;
while ($dades = $result->fetch_array()) {
    $spreadsheet->setActiveSheetIndex(0)
        ->setCellValue('B'.$index, $dades['codi'])
        ->setCellValue('C'.$index, $dades['nom'])
        ->setCellValue('D'.$index, $dades['ciutat']);
        $index++;
}

/* STYLES */
$spreadsheet->getDefaultStyle()
    ->getFont()
    ->setName('Arial')
    ->setSize(10);

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
$spreadsheet->getActiveSheet()->getStyle('B:D')->getAlignment()->setHorizontal('center');
//$spreadsheet->getActiveSheet()->getStyle('A:C')->getAllBorders()

// LOGO GITHUB
$drawing = new \PhpOffice\PhpSpreadsheet\Worksheet\Drawing();
$drawing->setName('Logo');
$drawing->setDescription('Logo');
$drawing->setPath('./images/github.png');
$drawing->setHeight(30);
$drawing->setCoordinates('E4');
$drawing->setOffsetX(20);
$drawing->setRotation(25);
$drawing->getShadow()->setVisible(true);
$drawing->getShadow()->setDirection(45);
$drawing->setWorksheet($spreadsheet->getActiveSheet());

$spreadsheet->getActiveSheet()->getStyle('B2:D3')->applyFromArray($styleArray);
$index2 = $index - 1;
$spreadsheet->getActiveSheet()->getStyle('B4:D'.$index2)->applyFromArray($tableFormat);
$spreadsheet->getActiveSheet()->getStyle('B2:D3')->getAlignment()->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);
$spreadsheet->getActiveSheet()->mergeCells('B2:D2');

$spreadsheet->getActiveSheet()->setTitle('Pàgina 1');
$writer = new Xlsx($spreadsheet);
$writer->save('practica1.xlsx');

/*$filename = 'practica1.xlsx';
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