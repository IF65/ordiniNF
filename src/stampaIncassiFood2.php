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

$blockedColumns = 4;

// riga titoli dei totali
$yCurrent = $yOrigin;
$xCurrent = $xOrigin;
$sheet->setCellValueExplicitByColumnAndRow($xCurrent, $yCurrent, 'Totali Periodo', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Ore Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Incasso Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Proc. Rep.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Inc. Rep. Anno Prec.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '% su Anno Prec.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Inc. totale nego.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Inc. totale nego.AP', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '%InciAPr', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '%suTot', DataType::TYPE_STRING );
$sheet->getStyle(RXY($xOrigin, $yCurrent, $xCurrent, $yCurrent))->applyFromArray($styleTitleRow); // applica lo stile alla riga
$sheet->mergeCells(RXY($xOrigin, $yCurrent,$xOrigin + $blockedColumns - 2, $yCurrent )); // unisce le prime n colonne bloccate  - 1
$sheet->getRowDimension($yCurrent)->setRowHeight(48);
$sheet->getStyle(RXY(1, $yCurrent, Coordinate::columnIndexFromString($sheet->getHighestColumn()), $yCurrent))->getAlignment()->setWrapText(true);
$sheet->getStyle(RXY(1, $yCurrent, Coordinate::columnIndexFromString($sheet->getHighestColumn()), $yCurrent))->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFDCDCDC');

// righe di totale reparti
$xCurrent = $xOrigin + $blockedColumns - 1;
foreach ($incassi['elencoReparti'] as $codice => $reparto) {
   $sheet->setCellValueExplicitByColumnAndRow($xCurrent, ++$yCurrent, $codice . ' - ' . $reparto, DataType::TYPE_STRING );
   $sheet->getRowDimension($yCurrent)->setRowHeight(24);
}

// seconda riga titoli
$xCurrent = $xOrigin;
$sheet->setCellValueExplicitByColumnAndRow($xCurrent, ++$yCurrent, 'Gruppo Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Tipo Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Ore Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Incasso Reparto', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Proc. Rep.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Inc. Rep. Anno Prec.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '% su Anno Prec.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Inc. totale nego.', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, 'Inc. totale nego.AP', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '%InciAPr', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow(++$xCurrent, $yCurrent, '%suTot', DataType::TYPE_STRING );
$sheet->getStyle(RXY($xOrigin, $yCurrent, $xCurrent, $yCurrent))->applyFromArray($styleTitleRow); // applica lo stile alla riga
$sheet->getRowDimension($yCurrent)->setRowHeight(48);
$sheet->getStyle(RXY(1, $yCurrent, Coordinate::columnIndexFromString($sheet->getHighestColumn()), $yCurrent))->getAlignment()->setWrapText(true);
$sheet->getStyle(RXY(1, $yCurrent, Coordinate::columnIndexFromString($sheet->getHighestColumn()), $yCurrent))->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFDCDCDC');

// blocca la riga e la colonna sopra e a sinistra di questa cella
$sheet->freezePane(XY($xOrigin + $blockedColumns, $yCurrent  + 1));


foreach ($incassi['dati'] as $incassoSede) {
    foreach ($incassi['elencoReparti'] as $codice => $reparto) {
        $xCurrent = $xOrigin;
        $sheet->setCellValueExplicitByColumnAndRow( $xCurrent, ++$yCurrent, $incassoSede['areaSede'], DataType::TYPE_STRING );
        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, $incassoSede['categoriaSede'], DataType::TYPE_STRING );
        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, $incassoSede['descrizioneSede'], DataType::TYPE_STRING );
        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, $codice . ' - ' . $reparto, DataType::TYPE_STRING );

        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, (key_exists( $codice, $incassoSede['reparti'] )) ? $incassoSede['reparti'][$codice]['ore'] : 0, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($currencyFormat);

        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, (key_exists( $codice, $incassoSede['reparti'] )) ? $incassoSede['reparti'][$codice]['incassato'] : 0, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

        $formula = "=if(" . XY($xCurrent - 1, $yCurrent) . "<> 0, " . XY($xCurrent , $yCurrent) . "/" . XY($xCurrent - 1, $yCurrent) . ", 0)" ;
        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($currencyFormat);

        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, (key_exists( $codice, $incassoSede['reparti'] )) ? $incassoSede['reparti'][$codice]['incassatoAP'] : 0, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

        $formula = "=if(" . XY($xCurrent , $yCurrent) . "<> 0, (" . XY($xCurrent - 2 , $yCurrent) . "-" . XY($xCurrent , $yCurrent) .  ")/" . XY($xCurrent , $yCurrent) . ", 0)" ;
        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($percentageFormat);

        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, (key_exists( 'incassato', $incassoSede )) ? $incassoSede['incassato'] : 0, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, (key_exists( 'incassatoAP', $incassoSede )) ? $incassoSede['incassatoAP'] : 0, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

        $formula = "=if(" . XY($xCurrent , $yCurrent) . "<> 0, (" . XY($xCurrent - 1 , $yCurrent) . "-" . XY($xCurrent , $yCurrent) .  ")/" . XY($xCurrent , $yCurrent) . ", 0)" ;
        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($percentageFormat);

        $formula = "=if(" . XY($xCurrent - 2, $yCurrent) . "<> 0, " . XY($xCurrent - 6, $yCurrent) . "/" . XY($xCurrent - 2, $yCurrent) . ", 0)" ;
        $sheet->setCellValueExplicitByColumnAndRow( ++$xCurrent, $yCurrent, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($xCurrent, $yCurrent)->getNumberFormat()->setFormatCode($percentageFormat);

        $sheet->getRowDimension($yCurrent)->setRowHeight(24);
    }
}

// formule riquadro totali
$yCurrent = 2;
for ($i = 0; $i < count($incassi['elencoReparti']); $i++) {

    // indici di riga
    $righe = [];
    for ($j = 1; $j <= count( $incassi['dati'] ); $j++) {
        $righe[] = $yCurrent + 1 + ($j * count( $incassi['elencoReparti'] ));
    }

    // ore
    $letteraColonna = Coordinate::stringFromColumnIndex(5);
    $sheet->setCellValueExplicitByColumnAndRow( 5, $yCurrent, "=SUBTOTAL(9,$letteraColonna" . implode(",$letteraColonna", $righe) . ')', DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(5, $yCurrent)->getNumberFormat()->setFormatCode($currencyFormat);

    // incasso reparto anno corrente
    $letteraColonna = Coordinate::stringFromColumnIndex(6);
    $sheet->setCellValueExplicitByColumnAndRow( 6, $yCurrent, "=SUBTOTAL(9,$letteraColonna" . implode(",$letteraColonna", $righe) . ')', DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(6, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

    // procapite anno corrente
    $formula = "=if(" . XY(5, $yCurrent) . "<> 0, " . XY(6 , $yCurrent) . "/" . XY(5, $yCurrent) . ", 0)" ;
    $sheet->setCellValueExplicitByColumnAndRow( 7, $yCurrent, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(7, $yCurrent)->getNumberFormat()->setFormatCode($currencyFormat);

    // incasso reparto anno precedente
    $letteraColonna = Coordinate::stringFromColumnIndex(8);
    $sheet->setCellValueExplicitByColumnAndRow( 8, $yCurrent, "=SUBTOTAL(9,$letteraColonna" . implode(",$letteraColonna", $righe) . ')', DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(8, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

    // incremento reparto su anno precedente
    $formula = "=if(" . XY(8 , $yCurrent) . "<> 0, (" . XY(6 , $yCurrent) . "-" . XY(8 , $yCurrent) .  ")/" . XY(8 , $yCurrent) . ", 0)" ;
    $sheet->setCellValueExplicitByColumnAndRow( 9, $yCurrent, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(9, $yCurrent)->getNumberFormat()->setFormatCode($percentageFormat);

    // incasso anno corrente
    $letteraColonna = Coordinate::stringFromColumnIndex(10);
    $sheet->setCellValueExplicitByColumnAndRow( 10, $yCurrent, "=SUBTOTAL(9,$letteraColonna" . implode(",$letteraColonna", $righe) . ')', DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(10, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

    // incasso anno precedente
    $letteraColonna = Coordinate::stringFromColumnIndex(11);
    $sheet->setCellValueExplicitByColumnAndRow( 11, $yCurrent, "=SUBTOTAL(9,$letteraColonna" . implode(",$letteraColonna", $righe) . ')', DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(11, $yCurrent)->getNumberFormat()->setFormatCode($integerFormat);

    // % incremento su anno precedente
    $formula = "=if(" . XY(11 , $yCurrent) . "<> 0, (" . XY(10 , $yCurrent) . "-" . XY(11 , $yCurrent) .  ")/" . XY(11 , $yCurrent) . ", 0)" ;
    $sheet->setCellValueExplicitByColumnAndRow( 12, $yCurrent, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(12, $yCurrent)->getNumberFormat()->setFormatCode($percentageFormat);

    // % incidenza reparto
    $formula = "=if(" . XY(10, $yCurrent) . "<> 0, " . XY(6, $yCurrent) . "/" . XY(10, $yCurrent) . ", 0)" ;
    $sheet->setCellValueExplicitByColumnAndRow( 13, $yCurrent, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow(13, $yCurrent)->getNumberFormat()->setFormatCode($percentageFormat);

    $yCurrent++;
}

// filtri
$sheet->setAutoFilter(RXY(1, count($incassi['elencoReparti'])+2, 4, $sheet->getHighestRow(Coordinate::stringFromColumnIndex(4))));

// larghezza colonne
$sheet->getColumnDimensionByColumn(1)->setWidth(18.0);
$sheet->getColumnDimensionByColumn(2)->setWidth(20.0);
$sheet->getColumnDimensionByColumn(3)->setWidth(25.0);
$sheet->getColumnDimensionByColumn(4)->setWidth(25.0);

$sheet->mergeCells('A2:C11');
$intervalliDate  = "Intervalli date di calcolo:\n";
$intervalliDate .= "Anno precedente dal " . date("d/m/Y",strtotime($incassi['inizioAP'])) . ' al ' . date("d/m/Y",strtotime($incassi['fineAP'])) .".\n";;
$intervalliDate .= "Anno effettivo dal " . date("d/m/Y",strtotime($incassi['inizio'])) . ' al ' . date("d/m/Y",strtotime($incassi['fine'])) .".\n";
$sheet->getStyle('A2')->getAlignment()->setWrapText(true);
$sheet->getStyle('A2')->applyFromArray(['font' => ['bold' => true], 'alignment' => ['horizontal' => Alignment::HORIZONTAL_LEFT, 'vertical' => Alignment::VERTICAL_TOP]]);
$sheet->setCellValueExplicitByColumnAndRow( 1, 2, $intervalliDate, DataType::TYPE_STRING );

$writer = new Xlsx($workBook);
$writer->save($fileName);

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
