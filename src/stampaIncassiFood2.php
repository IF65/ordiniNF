<?php
//@ini_set('memory_limit','8192M');

require '../vendor/autoload.php';
// leggo i dati da un file
$request = file_get_contents('../examples/incassi.json');
//$request = file_get_contents('php://input');
$data = json_decode($request, true);

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Shared\Date;

// verifico l'esistenza della cartella temp e se serve la creo
// con mask 777.
if (! file_exists ( '../temp' )) {
    $oldMask = umask(0);
    mkdir('../temp', 0777);
    umask($oldMask);
}

$fileName = '../temp/' . 'testReportIncassi2' . '.xlsx';

$incassi = json_decode($request, true);

$style = new Style();

// creazione del workbook
$workBook = new Spreadsheet();
$workBook->getDefaultStyle()->getFont()->setName('Arial');
$workBook->getDefaultStyle()->getFont()->setSize(12);
$workBook->getDefaultStyle()->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
$workBook->getProperties()
    ->setCreator("IF65 S.p.A. (Gruppo Italmark)")
    ->setLastModifiedBy("IF65 S.p.A.")
    ->setTitle("report incassi")
    ->setSubject("report incassi 2")
    ->setDescription("report incassi")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("IF65 Docs");

$sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
$sheet->setTitle('Periodo');
$sheet->getDefaultRowDimension()->setRowHeight(32);
$sheet->getDefaultColumnDimension()->setWidth(12);
//$sheet->getAlignment()->setWrapText(true);

$timeZone = new DateTimeZone('Europe/Rome');

$integerFormat = '###,###,##0;[Red][<0]-###,###,##0;';
$currencyFormat = '###,###,##0.00;[Red][<0]-###,###,##0.00;';
$percentageFormat = '0.00%;[Red][<0]-0.00%;';

$styleTitleRow = [
    'font' => ['bold' => true],
    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]
];

$styleBorder = [
    'borders' => [
        'outline' => [
            'borderStyle' => Border::BORDER_MEDIUM,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
];

$xOrigin = 1;
$yOrigin = 1;

// riga titoli principale
$currentRow = $yOrigin;
$currentColumn = $xOrigin;
$sheet->setCellValueExplicitByColumnAndRow($currentColumn, $currentRow, 'Totali Periodo', DataType::TYPE_STRING );
$currentColumn += 2;



$currentRow = $yOrigin + 2;
$currentColumn = $xOrigin;
$sheet->setCellValueExplicitByColumnAndRow($currentColumn, $currentRow, 'Gruppo Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, 'Tipo Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, 'Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, 'Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, 'Ore Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, 'Incasso Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, 'Proc. Rep.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, 'Inc. Rep. Anno Prec.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$currentColumn, $currentRow, '% su Anno Prec.', DataType::TYPE_STRING );



/* FORMATTAZIONI
 * ---------------------- */
$sheet->freezePane(XY($xOrigin + 4, $yOrigin + 1)); // blocca la riga ela colonna sopra e a sinistra di questa cella

$highestColumn = Coordinate::columnIndexFromString($sheet->getHighestColumn());

// Prima riga Titoli
$sheet->mergeCells(RXY($xOrigin, $yOrigin,$xOrigin + 3,$yOrigin ));
$sheet->getStyle(RXY($xOrigin, $yOrigin, $highestColumn, $yOrigin))->applyFromArray($styleTitleRow);

$sheet->mergeCells(RXY($xOrigin, $yOrigin + 1,$xOrigin + 3,$yOrigin +1));

// Seconda riga Titoli
$sheet->getStyle(RXY($xOrigin, $yOrigin + 2, $highestColumn, $yOrigin + 2))->applyFromArray($styleTitleRow);


$writer = new Xlsx($workBook);
$writer->save($fileName);∆

if (file_exists($fileName)) {
    header('Content-Description: File Transfer');
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="'.basename($fileName).'"');
    header('Expires: 0');
    header('Cache-Control: must-revalidate');
    header('Pragma: public');
    header('Content-Length: ' . filesize($fileName));
    readfile($fileName);
}

exit;

function XY(int $columnIndex, int $rowIndex): string {
    return Coordinate::stringFromColumnIndex($columnIndex) . "$rowIndex";
}

function RXY(int $column1Index, int $row1Index, int $column2Index, int $row2Index): string {
    return Coordinate::stringFromColumnIndex($column1Index) . "$row1Index" . ":" . Coordinate::stringFromColumnIndex($column2Index) . "$row2Index";
}
