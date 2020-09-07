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

$fileName = '../temp/' . 'testReportIncassi' . '.xlsx';

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
    ->setSubject("report incassi")
    ->setDescription("report incassi")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("IF65 Docs");

$sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
$sheet->setTitle('Periodo');
$sheet->getDefaultRowDimension()->setRowHeight(32);
$sheet->getDefaultColumnDimension()->setWidth(12);
//$sheet->getAlignment()->setWrapText(true);
$sheet->freezePane('E9'); // blocca la riga sopra questa

$timeZone = new DateTimeZone('Europe/Rome');


$integerFormat = '###,###,##0;[Red][<0]-###,###,##0;';
$currencyFormat = '###,###,##0.00;[Red][<0]-###,###,##0.00;';
$percentageFormat = '0.00%;[Red][<0]-0.00%;';

$styleBorderArray = [
    'borders' => [
        'outline' => [
            'borderStyle' => \PhpOffice\PhpSpreadsheet\Style\Border::BORDER_MEDIUM,
            'color' => ['argb' => 'FF000000'],
        ],
    ],
];

$currentRow = 1;

// riquadro subtotali
$sheet->mergeCells('A1:C1');

$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Totali Periodo', DataType::TYPE_STRING );
$currentColumn++;
$currentColumn++;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Tipo', DataType::TYPE_STRING );
foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, ucwords(strtolower(str_replace("\/", "\n", $descrizione)),"\t\r\n\f\v\/\\"), DataType::TYPE_STRING );
}
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Tot. Incasso', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Clienti', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Spesa Media', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Ore', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Procapite', DataType::TYPE_STRING );
$highestColumn = $sheet->getHighestColumn();
$sheet->getStyle('A'.$currentRow.':' . $highestColumn . $currentRow )->applyFromArray(
    ['font' => ['bold' => true], 'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]]);
$sheet->getStyle('A'.$currentRow.':'. $highestColumn . $currentRow)->applyFromArray($styleBorderArray);
$sheet->getRowDimension($currentRow)->setRowHeight(48);
$sheet->getStyle('A'.$currentRow.':'. $highestColumn . $currentRow)->getAlignment()->setWrapText(true);
$sheet->getStyle('A'.$currentRow.':'. $highestColumn . $currentRow)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFDCDCDC');
$currentRow++;

$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Anno prec.', DataType::TYPE_STRING );
$sheet->getRowDimension($currentRow)->setRowHeight(32);
$currentRow++;

$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Effettivo', DataType::TYPE_STRING );
$sheet->getRowDimension($currentRow)->setRowHeight(32);
$currentRow++;

$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'E +/- Suap', DataType::TYPE_STRING );
$sheet->getRowDimension($currentRow)->setRowHeight(32);
$currentRow++;

$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Inc.su tap', DataType::TYPE_STRING );
$sheet->getRowDimension($currentRow)->setRowHeight(32);
$currentRow++;

$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Inc.su tac', DataType::TYPE_STRING );
$sheet->getRowDimension($currentRow)->setRowHeight(32);
$currentRow++;

$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, '', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Dif su inc', DataType::TYPE_STRING );
$sheet->getRowDimension($currentRow)->setRowHeight(32);
$currentRow++;

$highestColumn = $sheet->getHighestColumn();
$sheet->getStyle('A'. ($currentRow - 6) . ':' . $highestColumn . ($currentRow - 1))->applyFromArray($styleBorderArray);
$sheet->getStyle('A'.$currentRow.':'. $highestColumn . $currentRow)->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFDCDCDC');

// dettagli
$currentColumn = 0;
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Gruppo Sedi', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Tipo Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Sede', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Tipo', DataType::TYPE_STRING );
foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, ucwords(strtolower(str_replace("\/", "\n", $descrizione)),"\t\r\n\f\v\/\\"), DataType::TYPE_STRING );
}
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Tot. Incasso', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Clienti', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Spesa Media', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Ore', DataType::TYPE_STRING );
$sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Procapite', DataType::TYPE_STRING );
$highestColumn = $sheet->getHighestColumn();
$sheet->getStyle('A'.$currentRow.':' . $highestColumn . $currentRow )->applyFromArray(
    ['font' => ['bold' => true], 'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]]);
$sheet->getStyle('A'.$currentRow.':'. $highestColumn . $currentRow)->applyFromArray($styleBorderArray);
$sheet->getRowDimension($currentRow)->setRowHeight(48);
$sheet->getStyle('A'.$currentRow.':'. $highestColumn . $currentRow)->getAlignment()->setWrapText(true);

$currentRow++;

foreach ($incassi['dati'] as $codiceSede => $datiSede) {
    $currentColumn = 0;
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['areaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['categoriaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['descrizioneSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Anno prec.', DataType::TYPE_STRING );
    foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
        $incassoAP = (key_exists($codice, $datiSede['reparti'])) ? $datiSede['reparti'][$codice]['incassatoAP'] : 0;
        $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $incassoAP, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($integerFormat);
    }
    $range = Coordinate::stringFromColumnIndex( 3 ) . $currentRow . ':' .
        Coordinate::stringFromColumnIndex( $currentColumn ) . $currentRow;
    $formula = "=sum($range)";
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($integerFormat);
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['clientiAP'], DataType::TYPE_NUMERIC );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($integerFormat);
    $XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn - 1) . $currentRow;
    $XY_clientiAP = Coordinate::stringFromColumnIndex($currentColumn) . $currentRow;
    $formula = "=if($XY_clientiAP<>0,$XY_totaleIncassatoAP/$XY_clientiAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($currencyFormat);
    $totaleOreAP = 0;
    foreach ($datiSede['reparti'] as $reparto) {
        $totaleOreAP += $reparto['oreAP'];
    }
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $totaleOreAP, DataType::TYPE_NUMERIC );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($currencyFormat);
    $XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn - 3) . $currentRow;
    $XY_oreAP = Coordinate::stringFromColumnIndex($currentColumn) . $currentRow;
    $formula = "=if($XY_oreAP<>0,$XY_totaleIncassatoAP/$XY_oreAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($currencyFormat);
    $sheet->getRowDimension($currentRow)->setRowHeight(32);
    $currentRow++;

    $currentColumn = 0;
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['areaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['categoriaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['descrizioneSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Effettivo', DataType::TYPE_STRING );
    foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
        $incasso = (key_exists($codice, $datiSede['reparti'])) ? $datiSede['reparti'][$codice]['incassato'] : 0;
        $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $incasso, DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($integerFormat);
    }
    $range = Coordinate::stringFromColumnIndex( 3 ) . $currentRow . ':' .
        Coordinate::stringFromColumnIndex( $currentColumn ) . $currentRow;
    $formula = "=sum($range)";
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($integerFormat);
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['clienti'], DataType::TYPE_NUMERIC );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($integerFormat);
    $XY_totaleIncassato = Coordinate::stringFromColumnIndex( $currentColumn - 1) . $currentRow;
    $XY_clienti = Coordinate::stringFromColumnIndex($currentColumn) . $currentRow;
    $formula = "=if($XY_clienti<>0,$XY_totaleIncassato/$XY_clienti,0)";
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($currencyFormat);
    $totaleOre = 0;
    foreach ($datiSede['reparti'] as $reparto) {
        $totaleOre += $reparto['ore'];
    }
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $totaleOre, DataType::TYPE_NUMERIC );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($currencyFormat);
    $XY_totaleIncassato = Coordinate::stringFromColumnIndex( $currentColumn - 3) . $currentRow;
    $XY_ore = Coordinate::stringFromColumnIndex($currentColumn) . $currentRow;
    $formula = "=if($XY_ore<>0,$XY_totaleIncassato/$XY_ore,0)";
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($currencyFormat);
    $sheet->getRowDimension($currentRow)->setRowHeight(32);
    $currentRow++;

    $currentColumn = 0;
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['areaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['categoriaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['descrizioneSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'E +/- Suap', DataType::TYPE_STRING );
    foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
        $currentColumn++;
        $XY_incassatoAP = Coordinate::stringFromColumnIndex( $currentColumn ) . ($currentRow - 2);
        $XY_incassato = Coordinate::stringFromColumnIndex( $currentColumn ) . ($currentRow - 1);
        $formula = "=if($XY_incassatoAP<>0,($XY_incassato - $XY_incassatoAP)/$XY_incassatoAP,0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    }
    $currentColumn++;
    $XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn) . ($currentRow - 2);
    $XY_totaleIncassato = Coordinate::stringFromColumnIndex($currentColumn) . ($currentRow - 1);
    $formula = "=if($XY_totaleIncassatoAP<>0,($XY_totaleIncassato - $XY_totaleIncassatoAP)/$XY_totaleIncassatoAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    $currentColumn++;
    $XY_clientiAP = Coordinate::stringFromColumnIndex( $currentColumn) . ($currentRow - 2);
    $XY_clienti = Coordinate::stringFromColumnIndex($currentColumn) . ($currentRow - 1);
    $formula = "=if($XY_clientiAP<>0,($XY_clienti - $XY_clientiAP)/$XY_clientiAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    $currentColumn++;
    $XY_spesaMediaAP = Coordinate::stringFromColumnIndex( $currentColumn) . ($currentRow - 2);
    $XY_spesaMedia = Coordinate::stringFromColumnIndex($currentColumn) . ($currentRow - 1);
    $formula = "=if($XY_spesaMediaAP<>0,($XY_spesaMedia - $XY_spesaMediaAP)/$XY_spesaMediaAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    $currentColumn++;
    $XY_oreAP = Coordinate::stringFromColumnIndex( $currentColumn) . ($currentRow - 2);
    $XY_ore = Coordinate::stringFromColumnIndex($currentColumn) . ($currentRow - 1);
    $formula = "=if($XY_oreAP<>0,($XY_ore - $XY_oreAP)/$XY_oreAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    $currentColumn++;
    $XY_procapiteAP = Coordinate::stringFromColumnIndex( $currentColumn) . ($currentRow - 2);
    $XY_procapite = Coordinate::stringFromColumnIndex($currentColumn) . ($currentRow - 1);
    $formula = "=if($XY_procapiteAP<>0,($XY_procapite - $XY_procapiteAP)/$XY_procapiteAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    $sheet->getRowDimension($currentRow)->setRowHeight(32);
    $currentRow++;

    $currentColumn = 0;
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['areaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['categoriaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['descrizioneSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Inc.su tap', DataType::TYPE_STRING );
    foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
        $currentColumn++;
        $XY_incassatoAP = Coordinate::stringFromColumnIndex( $currentColumn ) . ($currentRow - 3);
        $range = Coordinate::stringFromColumnIndex( 3 ) . ($currentRow - 3) . ':' .
            Coordinate::stringFromColumnIndex( 3 + count( $incassi['elencoReparti'] ) - 1 ) . ($currentRow - 3);
        $formula = "=if(sum($range)<>0,$XY_incassatoAP/sum($range),0)";
        $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    }
    $sheet->getRowDimension($currentRow)->setRowHeight(32);
    $currentRow++;

    $currentColumn = 0;
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['areaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['categoriaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['descrizioneSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Inc.su tac', DataType::TYPE_STRING );
    foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
        $currentColumn++;
        $XY_incassato = Coordinate::stringFromColumnIndex( $currentColumn ) . ($currentRow - 3);
        $range = Coordinate::stringFromColumnIndex( 3 ) . ($currentRow - 3) . ':' .
            Coordinate::stringFromColumnIndex( 3 + count( $incassi['elencoReparti'] ) - 1 ) . ($currentRow - 3);
        $formula = "=if(sum($range)<>0,$XY_incassato/sum($range),0)";
        $sheet->setCellValueExplicitByColumnAndRow($currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    }
    $sheet->getRowDimension($currentRow)->setRowHeight(32);
    $currentRow++;

    $currentColumn = 0;
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['areaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['categoriaSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, $datiSede['descrizioneSede'], DataType::TYPE_STRING );
    $sheet->setCellValueExplicitByColumnAndRow( ++$currentColumn, $currentRow, 'Dif su inc', DataType::TYPE_STRING );
    foreach ($incassi['elencoReparti'] as $codice => $descrizione) {
        $currentColumn++;
        $XY_incidenzaAP = Coordinate::stringFromColumnIndex( $currentColumn ) . ($currentRow - 2);
        $XY_incidenza = Coordinate::stringFromColumnIndex( $currentColumn ) . ($currentRow - 1);
        $formula = "=$XY_incidenza - $XY_incidenzaAP";
        $sheet->setCellValueExplicitByColumnAndRow($currentColumn, $currentRow, $formula, DataType::TYPE_FORMULA );
        $sheet->getStyleByColumnAndRow($currentColumn, $currentRow)->getNumberFormat()->setFormatCode($percentageFormat);
    }

    $sheet->getRowDimension($currentRow)->setRowHeight(32);
    $currentRow++;

    $highestColumn = $sheet->getHighestColumn();
    $sheet->getStyle('A'. ($currentRow - 6) . ':' . $highestColumn . ($currentRow - 1))->applyFromArray($styleBorderArray);

}
$sheet->setAutoFilter('A8:A' . $sheet->getHighestRow('A'));
$sheet->setAutoFilter('B8:A' . $sheet->getHighestRow('A'));
$sheet->setAutoFilter('C8:A' . $sheet->getHighestRow('A'));
$sheet->getColumnDimensionByColumn(1)->setWidth(20.0);
$sheet->getColumnDimensionByColumn(2)->setWidth(20.0);
$sheet->getColumnDimensionByColumn(3)->setWidth(25.0);
$sheet->getColumnDimensionByColumn(4)->setWidth(18.0);

// formule del riquadro totali
$highestRow = $sheet->getHighestRow('A');

// anno precedente
$currentColumn = 4;
for ($i = 0; $i < count($incassi['elencoReparti']); $i++) {
    $currentColumn++;

    $formula = "=SUBTOTAL(9,";
    for($r = 9; $r < $highestRow; $r += 6) {
        $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
    }
    if (preg_match('/^(.*).$/', $formula, $matches)) {
        $formula = $matches[1] . ')';
    }

    $sheet->setCellValueExplicitByColumnAndRow($currentColumn , 2, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn , 2)->getNumberFormat()->setFormatCode($integerFormat);

    $XY_incassatoAP = Coordinate::stringFromColumnIndex( $currentColumn ) . 2;
    $XY_incassato = Coordinate::stringFromColumnIndex( $currentColumn ) . 3;
    $formula = "=if($XY_incassatoAP<>0,($XY_incassato - $XY_incassatoAP)/$XY_incassatoAP,0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 4, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, 4)->getNumberFormat()->setFormatCode($percentageFormat);

    // incidenza su tap
    $XY_incassatoAP = Coordinate::stringFromColumnIndex( $currentColumn ) . 2;
    $range = Coordinate::stringFromColumnIndex( 3 ) . '2:' . Coordinate::stringFromColumnIndex( 3 + count( $incassi['elencoReparti'])  - 1) . 2;
    $formula = "=if(sum($range)<>0,$XY_incassatoAP/sum($range),0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 5, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, 5)->getNumberFormat()->setFormatCode($percentageFormat);

    // incidenza su tac
    $XY_incassato = Coordinate::stringFromColumnIndex( $currentColumn ) . 3;
    $range = Coordinate::stringFromColumnIndex( 3 ) . '3:' . Coordinate::stringFromColumnIndex( 3 + count( $incassi['elencoReparti'])  - 1) . 3;
    $formula = "=if(sum($range)<>0,$XY_incassato/sum($range),0)";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 6, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, 6)->getNumberFormat()->setFormatCode($percentageFormat);

    // tac - tap
    $XY_incidenza_tap = Coordinate::stringFromColumnIndex( $currentColumn ) . 5;
    $XY_incidenza_tac = Coordinate::stringFromColumnIndex( $currentColumn ) . 6;
    $formula = "=$XY_incidenza_tac - $XY_incidenza_tap";
    $sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 7, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, 7)->getNumberFormat()->setFormatCode($percentageFormat);
}

// totale AP
$currentColumn++;
$formula = "=SUBTOTAL(9,";
for($r = 9; $r < $highestRow; $r += 6) {
    $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
}
if (preg_match('/^(.*).$/', $formula, $matches)) {
    $formula = $matches[1] . ')';
}
$sheet->setCellValueExplicitByColumnAndRow($currentColumn , 2, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn , 2)->getNumberFormat()->setFormatCode($integerFormat);

// clienti AP
$currentColumn++;
$formula = "=SUBTOTAL(9,";
for($r = 9; $r < $highestRow; $r += 6) {
    $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
}
if (preg_match('/^(.*).$/', $formula, $matches)) {
    $formula = $matches[1] . ')';
}
$sheet->setCellValueExplicitByColumnAndRow($currentColumn , 2, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn , 2)->getNumberFormat()->setFormatCode($integerFormat);

// spesa media AP
$currentColumn++;
$XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn - 2) . 2;
$XY_clientiAP = Coordinate::stringFromColumnIndex($currentColumn - 1) . 2;
$formula = "=if($XY_clientiAP<>0,$XY_totaleIncassatoAP/$XY_clientiAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 2, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 2)->getNumberFormat()->setFormatCode($currencyFormat);

// ore AP
$currentColumn++;
$formula = "=SUBTOTAL(9,";
for($r = 9; $r < $highestRow; $r += 6) {
    $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
}
if (preg_match('/^(.*).$/', $formula, $matches)) {
    $formula = $matches[1] . ')';
}
$sheet->setCellValueExplicitByColumnAndRow($currentColumn , 2, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn , 2)->getNumberFormat()->setFormatCode($currencyFormat);

// procapite AP
$currentColumn++;
$XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn - 4) . 2;
$XY_oreAP = Coordinate::stringFromColumnIndex($currentColumn - 1) . 2;
$formula = "=if($XY_oreAP<>0,$XY_totaleIncassatoAP/$XY_oreAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 2, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 2)->getNumberFormat()->setFormatCode($currencyFormat);


// anno effettivo
$currentColumn = 4;
for ($i = 0; $i < count($incassi['elencoReparti']); $i++) {
    $currentColumn++;

    $formula = "=SUBTOTAL(9,";
    for($r = 10; $r < $highestRow; $r += 6) {
        $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
    }
    if (preg_match('/^(.*).$/', $formula, $matches)) {
        $formula = $matches[1] . ')';
    }

    $sheet->setCellValueExplicitByColumnAndRow($currentColumn, 3, $formula, DataType::TYPE_FORMULA );
    $sheet->getStyleByColumnAndRow($currentColumn, 3)->getNumberFormat()->setFormatCode($integerFormat);
}

// totale
$currentColumn++;
$formula = "=SUBTOTAL(9,";
for($r = 10; $r < $highestRow; $r += 6) {
    $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
}
if (preg_match('/^(.*).$/', $formula, $matches)) {
    $formula = $matches[1] . ')';
}
$sheet->setCellValueExplicitByColumnAndRow($currentColumn , 3, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn , 3)->getNumberFormat()->setFormatCode($integerFormat);

// clienti
$currentColumn++;
$formula = "=SUBTOTAL(9,";
for($r = 10; $r < $highestRow; $r += 6) {
    $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
}
if (preg_match('/^(.*).$/', $formula, $matches)) {
    $formula = $matches[1] . ')';
}
$sheet->setCellValueExplicitByColumnAndRow($currentColumn , 3, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn , 3)->getNumberFormat()->setFormatCode($integerFormat);

// spesa media
$currentColumn++;
$XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn - 2) . 3;
$XY_clientiAP = Coordinate::stringFromColumnIndex($currentColumn - 1) . 3;
$formula = "=if($XY_clientiAP<>0,$XY_totaleIncassatoAP/$XY_clientiAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 3, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 3)->getNumberFormat()->setFormatCode($currencyFormat);

// ore
$currentColumn++;
$formula = "=SUBTOTAL(9,";
for($r = 10; $r < $highestRow; $r += 6) {
    $formula .= Coordinate::stringFromColumnIndex($currentColumn) . $r . ',';
}
if (preg_match('/^(.*).$/', $formula, $matches)) {
    $formula = $matches[1] . ')';
}
$sheet->setCellValueExplicitByColumnAndRow($currentColumn , 3, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn , 3)->getNumberFormat()->setFormatCode($currencyFormat);

// procapite
$currentColumn++;
$XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn - 4) . 3;
$XY_oreAP = Coordinate::stringFromColumnIndex($currentColumn - 1) . 3;
$formula = "=if($XY_oreAP<>0,$XY_totaleIncassatoAP/$XY_oreAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 3, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 3)->getNumberFormat()->setFormatCode($currencyFormat);

// percentuali sui totali
$currentColumn = 2 + count($incassi['elencoReparti']) + 1;
$XY_totaleIncassatoAP = Coordinate::stringFromColumnIndex( $currentColumn) . 2;
$XY_totaleIncassato = Coordinate::stringFromColumnIndex($currentColumn) . 3;
$formula = "=if($XY_totaleIncassatoAP<>0,($XY_totaleIncassato - $XY_totaleIncassatoAP)/$XY_totaleIncassatoAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 4, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 4)->getNumberFormat()->setFormatCode($percentageFormat);

$currentColumn++;
$XY_clientiAP = Coordinate::stringFromColumnIndex( $currentColumn) . 2;
$XY_clienti = Coordinate::stringFromColumnIndex($currentColumn) . 3;
$formula = "=if($XY_clientiAP<>0,($XY_clienti - $XY_clientiAP)/$XY_clientiAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 4, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 4)->getNumberFormat()->setFormatCode($percentageFormat);

$currentColumn++;
$XY_spesaMediaAP = Coordinate::stringFromColumnIndex( $currentColumn) . 2;
$XY_spesaMedia = Coordinate::stringFromColumnIndex($currentColumn) . 3;
$formula = "=if($XY_spesaMediaAP<>0,($XY_spesaMedia - $XY_spesaMediaAP)/$XY_spesaMediaAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 4, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 4)->getNumberFormat()->setFormatCode($percentageFormat);

$currentColumn++;
$XY_spesaMediaAP = Coordinate::stringFromColumnIndex( $currentColumn) . 2;
$XY_spesaMedia = Coordinate::stringFromColumnIndex($currentColumn) . 3;
$formula = "=if($XY_spesaMediaAP<>0,($XY_spesaMedia - $XY_spesaMediaAP)/$XY_spesaMediaAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 4, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 4)->getNumberFormat()->setFormatCode($percentageFormat);

$currentColumn++;
$XY_oreAP = Coordinate::stringFromColumnIndex( $currentColumn) . 2;
$XY_ore = Coordinate::stringFromColumnIndex($currentColumn) . 3;
$formula = "=if($XY_oreAP<>0,($XY_ore - $XY_oreAP)/$XY_oreAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 4, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 4)->getNumberFormat()->setFormatCode($percentageFormat);

$currentColumn++;
$XY_procapiteAP = Coordinate::stringFromColumnIndex( $currentColumn) . 2;
$XY_procapite = Coordinate::stringFromColumnIndex($currentColumn) . 3;
$formula = "=if($XY_procapiteAP<>0,($XY_procapite - $XY_procapiteAP)/$XY_procapiteAP,0)";
$sheet->setCellValueExplicitByColumnAndRow( $currentColumn, 4, $formula, DataType::TYPE_FORMULA );
$sheet->getStyleByColumnAndRow($currentColumn, 4)->getNumberFormat()->setFormatCode($percentageFormat);

$sheet->mergeCells('A2:C7');
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
    exit;
}

