<?php
//@ini_set('memory_limit','8192M');

require '../vendor/autoload.php';
// leggo i dati da un file
//$request = file_get_contents('/Users/if65/Desktop/reportControllo.json');
$request = file_get_contents('php://input');
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

$style = new Style();

// leggo i parametri contenuti nel file
$nomeFile = $data['nomeFile'];
$file = '../temp/'.$nomeFile.'.xlsx';

// creazione del workbook
$workBook = new Spreadsheet();
$workBook->getDefaultStyle()->getFont()->setName('Arial');
$workBook->getDefaultStyle()->getFont()->setSize(12);
$workBook->getProperties()
    ->setCreator("IF65 S.p.A. (Gruppo Italmark)")
    ->setLastModifiedBy("IF65 S.p.A.")
    ->setTitle("Report Caricamento Corrispettivi Agenzia delle Entrate")
    ->setSubject("Report Caricamento Corrispettivi Agenzia delle Entrate")
    ->setDescription("Report Caricamento Corrispettivi Agenzia delle Entrate")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("IF65 Docs");


//$workBook->createSheet();
$sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
$sheet->setTitle('Testate');

$timeZone = new DateTimeZone('Europe/Rome');

$integerFormat = '###,###,##0;[Red][<0]-###,###,##0;';
$currencyFormat = '###,###,##0.00;[Red][<0]-###,###,##0.00;';

$highestRowIndex = 1;
$highestColumnIndex = 12;

// testata
// --------------------------------------------------------------------------------
$sheet->setCellValueByColumnAndRow(1, 1, strtoupper('Sede'));
$sheet->setCellValueByColumnAndRow(2, 1, strtoupper('Descrizione'));
$sheet->setCellValueByColumnAndRow(3, 1, strtoupper('Matricola'));
$sheet->setCellValueByColumnAndRow(4, 1, strtoupper('Data'));
$sheet->setCellValueByColumnAndRow(5, 1, strtoupper('Prog.'));
$sheet->setCellValueByColumnAndRow(6, 1, strtoupper('Imposta'));
$sheet->setCellValueByColumnAndRow(7, 1, strtoupper('Resi'));
$sheet->setCellValueByColumnAndRow(8, 1, strtoupper('Resi QC'));
$sheet->setCellValueByColumnAndRow(9, 1, strtoupper('Delta'));
$sheet->setCellValueByColumnAndRow(10, 1, strtoupper('Totale'));
$sheet->setCellValueByColumnAndRow(11, 1, strtoupper('Totale QC'));
$sheet->setCellValueByColumnAndRow(12, 1, strtoupper('Delta'));

$headerRowCount = 1;
$row = $headerRowCount;
foreach ($data['righe'] as $riga) {
    if ($riga['rigaPrincipale']) {
        $row += 1;

        $sheet->getCellByColumnAndRow( 1, $row )->setValueExplicit( strtoupper( $riga['sedeCodice'] ), DataType::TYPE_STRING );
        $sheet->getCellByColumnAndRow( 2, $row )->setValueExplicit( strtoupper( $riga['sedeDescrizione'] ), DataType::TYPE_STRING );
        $sheet->getCellByColumnAndRow( 3, $row )->setValueExplicit( strtoupper( $riga['matricola'] ), DataType::TYPE_STRING );

        $date = new \DateTime( $riga['data'] );
        $sheet->setCellValueByColumnAndRow( 4, $row, Date::PHPToExcel( $date->setTimezone( $timeZone )->format( 'Y-m-d' ) ) );
        $sheet->getCellByColumnAndRow( 4, $row )->getStyle()->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_DDMMYYYY );

        $sheet->getCellByColumnAndRow( 5, $row )->setValueExplicit( $riga['progressivo'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 6, $row )->setValueExplicit( $riga['imposta'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 7, $row )->setValueExplicit( $riga['ammontareResi'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 8, $row )->setValueExplicit( $riga['ammontareResiQc'], DataType::TYPE_NUMERIC );
        $formula = '=' . Coordinate::stringFromColumnIndex( 8 ) . "$row" . '-' . Coordinate::stringFromColumnIndex( 7 ) . "$row";
        $sheet->getCellByColumnAndRow( 9, $row )->setValueExplicit( $formula, DataType::TYPE_FORMULA );
        $sheet->getCellByColumnAndRow( 10, $row )->setValueExplicit( $riga['ammontare'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 11, $row )->setValueExplicit( $riga['ammontareQc'], DataType::TYPE_NUMERIC );
        $formula = '=' . Coordinate::stringFromColumnIndex( 11 ) . "$row" . '-' . Coordinate::stringFromColumnIndex( 10 ) . "$row";
        $sheet->getCellByColumnAndRow( 12, $row )->setValueExplicit( $formula, DataType::TYPE_FORMULA );
    }
}

$highestRowIndex = $row + 1;

// formattazione colonne
$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex(1)."$highestRowIndex")->getAlignment()->setHorizontal('center');
$sheet->getStyle(Coordinate::stringFromColumnIndex(4).'1:'.Coordinate::stringFromColumnIndex(4)."$highestRowIndex")->getAlignment()->setHorizontal('center');
$sheet->getStyle(Coordinate::stringFromColumnIndex(5).'1:'.Coordinate::stringFromColumnIndex(5)."$highestRowIndex")->getNumberFormat()->setFormatCode($integerFormat);
$sheet->getStyle(Coordinate::stringFromColumnIndex(6).'1:'.Coordinate::stringFromColumnIndex(12)."$highestRowIndex")->getNumberFormat()->setFormatCode($currencyFormat);

$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex($highestColumnIndex).'1')->getFont()->setBold(true);
$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex($highestColumnIndex).'1')->getAlignment()->setHorizontal('center');

for ($i = 1;$i <= $highestColumnIndex; $i++) $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($i))->setAutoSize(true);

$sheet->freezePane('A2');

// righe
$workBook->createSheet(1);
$sheet = $workBook->setActiveSheetIndex(1); // la numerazione dei worksheet parte da 0
$sheet->setTitle('Righe');

$timeZone = new DateTimeZone('Europe/Rome');

$integerFormat = '###,###,##0;[Red][<0]-###,###,##0;';
$currencyFormat = '###,###,##0.00;[Red][<0]-###,###,##0.00;';

$highestRowIndex = 1;
$highestColumnIndex = 12;

// testata
// --------------------------------------------------------------------------------
$sheet->setCellValueByColumnAndRow(1, 1, strtoupper('Sede'));
$sheet->setCellValueByColumnAndRow(2, 1, strtoupper('Descrizione'));
$sheet->setCellValueByColumnAndRow(3, 1, strtoupper('Matricola'));
$sheet->setCellValueByColumnAndRow(4, 1, strtoupper('Data'));
$sheet->setCellValueByColumnAndRow(5, 1, strtoupper('Prog.'));
$sheet->setCellValueByColumnAndRow(6, 1, strtoupper('Nat.'));
$sheet->setCellValueByColumnAndRow(7, 1, strtoupper('Al.%'));
$sheet->setCellValueByColumnAndRow(8, 1, strtoupper('Imposta'));
$sheet->setCellValueByColumnAndRow(9, 1, strtoupper('Resi'));
$sheet->setCellValueByColumnAndRow(10, 1, strtoupper('Resi QC'));
$sheet->setCellValueByColumnAndRow(11, 1, strtoupper('Delta'));
$sheet->setCellValueByColumnAndRow(12, 1, strtoupper('Totale'));
$sheet->setCellValueByColumnAndRow(13, 1, strtoupper('Totale QC'));
$sheet->setCellValueByColumnAndRow(14, 1, strtoupper('Delta'));

$headerRowCount = 1;
$row = $headerRowCount;
foreach ($data['righe'] as $riga) {
    if (! $riga['rigaPrincipale']) {
        $row += 1;

        $sheet->getCellByColumnAndRow( 1, $row )->setValueExplicit( strtoupper( $riga['sedeCodice'] ), DataType::TYPE_STRING );
        $sheet->getCellByColumnAndRow( 2, $row )->setValueExplicit( strtoupper( $riga['sedeDescrizione'] ), DataType::TYPE_STRING );
        $sheet->getCellByColumnAndRow( 3, $row )->setValueExplicit( strtoupper( $riga['matricola'] ), DataType::TYPE_STRING );

        $date = new \DateTime( $riga['data'] );
        $sheet->setCellValueByColumnAndRow( 4, $row, Date::PHPToExcel( $date->setTimezone( $timeZone )->format( 'Y-m-d' ) ) );
        $sheet->getCellByColumnAndRow( 4, $row )->getStyle()->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_DDMMYYYY );

        $sheet->getCellByColumnAndRow( 5, $row )->setValueExplicit( $riga['progressivo'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 6, $row )->setValueExplicit( $riga['natura'], DataType::TYPE_STRING );
        $sheet->getCellByColumnAndRow( 7, $row )->setValueExplicit( $riga['aliquotaIva'], DataType::TYPE_NUMERIC );

        $sheet->getCellByColumnAndRow( 8, $row )->setValueExplicit( $riga['imposta'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 9, $row )->setValueExplicit( $riga['ammontareResi'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 10, $row )->setValueExplicit( $riga['ammontareResiQc'], DataType::TYPE_NUMERIC );
        $formula = '=' . Coordinate::stringFromColumnIndex( 10 ) . "$row" . '-' . Coordinate::stringFromColumnIndex( 7 ) . "$row";
        $sheet->getCellByColumnAndRow( 11, $row )->setValueExplicit( $formula, DataType::TYPE_FORMULA );
        $sheet->getCellByColumnAndRow( 12, $row )->setValueExplicit( $riga['ammontare'], DataType::TYPE_NUMERIC );
        $sheet->getCellByColumnAndRow( 13, $row )->setValueExplicit( $riga['ammontareQc'], DataType::TYPE_NUMERIC );
        $formula = '=' . Coordinate::stringFromColumnIndex( 13 ) . "$row" . '-' . Coordinate::stringFromColumnIndex( 10 ) . "$row";
        $sheet->getCellByColumnAndRow( 14, $row )->setValueExplicit( $formula, DataType::TYPE_FORMULA );
    }
}

$highestRowIndex = $row + 1;

// formattazione colonne
$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex(1)."$highestRowIndex")->getAlignment()->setHorizontal('center');
$sheet->getStyle(Coordinate::stringFromColumnIndex(4).'1:'.Coordinate::stringFromColumnIndex(4)."$highestRowIndex")->getAlignment()->setHorizontal('center');
$sheet->getStyle(Coordinate::stringFromColumnIndex(5).'1:'.Coordinate::stringFromColumnIndex(5)."$highestRowIndex")->getNumberFormat()->setFormatCode($integerFormat);
$sheet->getStyle(Coordinate::stringFromColumnIndex(6).'1:'.Coordinate::stringFromColumnIndex(6)."$highestRowIndex")->getAlignment()->setHorizontal('center');
$sheet->getStyle(Coordinate::stringFromColumnIndex(7).'1:'.Coordinate::stringFromColumnIndex(7)."$highestRowIndex")->getNumberFormat()->setFormatCode($integerFormat);
$sheet->getStyle(Coordinate::stringFromColumnIndex(8).'1:'.Coordinate::stringFromColumnIndex(14)."$highestRowIndex")->getNumberFormat()->setFormatCode($currencyFormat);
$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex($highestColumnIndex).'1')->getFont()->setBold(true);
$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex($highestColumnIndex).'1')->getAlignment()->setHorizontal('center');

for ($i = 1;$i <= $highestColumnIndex; $i++) $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($i))->setAutoSize(true);

$sheet->freezePane('A2');

$sheet = $workBook->setActiveSheetIndex(0);

$writer = new Xlsx($workBook);
$writer->save($file);

if (file_exists($file)) {
    header('Content-Description: File Transfer');
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="'.basename($file).'"');
    header('Expires: 0');
    header('Cache-Control: must-revalidate');
    header('Pragma: public');
    header('Content-Length: ' . filesize($file));
    readfile($file);
    exit;
}

