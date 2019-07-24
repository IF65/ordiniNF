<?php
//@ini_set('memory_limit','8192M');

require '../vendor/autoload.php';
// leggo i dati da un file
$request = file_get_contents('../examples/giacenze.json');
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
    ->setTitle("Giacenze")
    ->setSubject("Giacenze")
    ->setDescription("Esportazione Giacenze")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("SM Docs");

// creazione degli Sheet (uno per ogni ordine)
    $workBook->createSheet();
    $sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
    $sheet->setTitle('Giacenze');

    $timeZone = new DateTimeZone('Europe/Rome');


    // testata colonne
    // --------------------------------------------------------------------------------
    $sheet->setCellValue('A1', strtoupper("Sede"));
    $sheet->setCellValue('B1', strtoupper("Giacenza\nIniziale"));
    $sheet->setCellValue('C1', strtoupper("Valore\nGiacenza Iniziale"));
    $sheet->setCellValue('D1', strtoupper("Giacenza"));
    $sheet->setCellValue('E1', strtoupper("Valore\nGiacenza"));
    $sheet->setCellValue('F1', strtoupper("Giacenza\nEliminati"));
    $sheet->setCellValue('G1', strtoupper("Valore\nGiacenza Eliminati"));
    $sheet->setCellValue('H1', strtoupper("Giacenza\nObsoleti"));
    $sheet->setCellValue('I1', strtoupper("Valore\nGiacenza Obsoleti"));

    // scrittura righe
    // --------------------------------------------------------------------------------
    $primaRigaDati = 2;

    $righe = $data['righe'];
    for ($i = 0; $i < count($righe); $i++) {
        $r = $i + 2;
        // righe
        $sheet->getCell('A'.$r)->setValueExplicit($righe[$i]['sede'],DataType::TYPE_STRING);
        $sheet->getCell('B'.$r)->setValueExplicit($righe[$i]['giacenzaIniziale'],DataType::TYPE_NUMERIC);
        $sheet->getCell('C'.$r)->setValueExplicit($righe[$i]['valoreGiacenzaIniziale'],DataType::TYPE_NUMERIC);
        $sheet->getCell('D'.$r)->setValueExplicit($righe[$i]['giacenza'],DataType::TYPE_NUMERIC);
        $sheet->getCell('E'.$r)->setValueExplicit($righe[$i]['valoreGiacenza'],DataType::TYPE_NUMERIC);
        $sheet->getCell('F'.$r)->setValueExplicit($righe[$i]['giacenzaEliminati'],DataType::TYPE_NUMERIC);
        $sheet->getCell('G'.$r)->setValueExplicit($righe[$i]['valoreGiacenzaEliminati'],DataType::TYPE_NUMERIC);
        $sheet->getCell('H'.$r)->setValueExplicit($righe[$i]['giacenzaObsoleti'],DataType::TYPE_NUMERIC);
        $sheet->getCell('I'.$r)->setValueExplicit($righe[$i]['valoreGiacenzaObsoleti'],DataType::TYPE_NUMERIC);
    }

    $ultimaRigaDati = count($righe) + $primaRigaDati -1;
    // formattazione
    // --------------------------------------------------------------------------------
    $sheet->getDefaultRowDimension()->setRowHeight(20);
    $sheet->getRowDimension('1')->setRowHeight(40);
    $sheet->setShowGridlines(true);

    foreach (range('A','I') as $col) {$sheet->getColumnDimension($col)->setAutoSize(true);}

    // colonne descrizione articolo + prezzi
    $sheet->getStyle(sprintf("%s%s%s%s%s",'A',$primaRigaDati - 1,':','I',$primaRigaDati - 1))->
    getAlignment()->setHorizontal('center')->setVertical('center');
    $sheet->getStyle(sprintf("%s%s%s%s%s",'A',$primaRigaDati - 1,':','I',$primaRigaDati - 1))->
    getAlignment()->setWrapText(true);
    $sheet->getStyle('A1:I1')->getAlignment()->applyFromArray(['font' => ['bold' => true]]);
    $sheet->getStyle(sprintf("%s%s%s%s%s",'B',$primaRigaDati,':','I',$ultimaRigaDati))->
    getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00;');
    $sheet->getStyle(sprintf("%s%s%s%s%s",'B',$primaRigaDati,':','B',$ultimaRigaDati))->
    getNumberFormat()->setFormatCode('###,###,##0;[Red][<0]-###,###,##0;');
    $sheet->getStyle(sprintf("%s%s%s%s%s",'D',$primaRigaDati,':','D',$ultimaRigaDati))->
    getNumberFormat()->setFormatCode('###,###,##0;[Red][<0]-###,###,##0;');
    $sheet->getStyle(sprintf("%s%s%s%s%s",'F',$primaRigaDati,':','F',$ultimaRigaDati))->
    getNumberFormat()->setFormatCode('###,###,##0;[Red][<0]-###,###,##0;');
    $sheet->getStyle(sprintf("%s%s%s%s%s",'H',$primaRigaDati,':','H',$ultimaRigaDati))->
    getNumberFormat()->setFormatCode('###,###,##0;[Red][<0]-###,###,##0;');
/*
    // quantita + sconto merce
    $sheet->getStyle(sprintf("%s%s%s%s",$colQuantitaTotale,'1:',$colSCIndex['LAST'],$primaRigaDati-1))->getFont()->setBold(true);
    $sheet->getStyle($colQuantitaTotale.'1')->getAlignment()->setHorizontal('center')->setVertical('center');
    $sheet->getStyle($colScontoMerceTotale.'1')->getAlignment()->setHorizontal('center')->setVertical('center');
    $sheet->getStyle(sprintf("%s%s%s%s%s",$colQuantitaTotale,$primaRigaDati,':',$colSCIndex['LAST'],$primaRigaDati+count($righe)-1))->
    getAlignment()->setHorizontal('center');

    // larghezza colonne (non uso volutamente autowidth)
    $sheet->getColumnDimension('A')->setWidth(25);
    $sheet->getColumnDimension('B')->setWidth(15);
    $sheet->getColumnDimension('C')->setWidth(9);
    $sheet->getColumnDimension('D')->setWidth(28);
    $sheet->getColumnDimension('E')->setWidth(14);
    $sheet->getColumnDimension('F')->setWidth(14);
    $sheet->getColumnDimension('G')->setWidth(10);
    $sheet->getColumnDimension('H')->setWidth(10);
    $sheet->getColumnDimension('I')->setWidth(4);
    $sheet->getColumnDimension('J')->setWidth(4);
    $sheet->getColumnDimension('K')->setWidth(9);
    $sheet->getColumnDimension('L')->setWidth(9);
    $sheet->getColumnDimension('M')->setWidth(6);
    $sheet->getColumnDimension('N')->setWidth(6);
    $sheet->getColumnDimension('O')->setWidth(6);
    $sheet->getColumnDimension('P')->setWidth(6);
    $sheet->getColumnDimension('Q')->setWidth(6);
    $sheet->getColumnDimension('R')->setWidth(6);
    $sheet->getColumnDimension('S')->setWidth(9);
    $sheet->getColumnDimension('T')->setWidth(9);
    $sheet->getColumnDimension('U')->setWidth(9);
    $sheet->getColumnDimension('V')->setWidth(9);
    $sheet->getColumnDimension('W')->setWidth(9);
    $sheet->getColumnDimension('X')->setWidth(9);
    $sheet->getColumnDimension('Y')->setWidth(9);

    $col = $QFirst;
    for ($i = 0; $i<count($filiali); $i++) { //<- quantita
        $sheet->getColumnDimension($col)->setWidth(4);
        $col++;
    }

    $sheet->getColumnDimension($col)->setWidth(9);

    $col = $SCFirst;
    for ($i = 0; $i<count($filiali); $i++) { //<- sconto merce
        $sheet->getColumnDimension($col)->setWidth(4);
        $col++;
    }
    // testata colonne
    $sheet->getStyle('A9:X10')->getAlignment()->setHorizontal('center')->setVertical('center');
    $sheet->getStyle('A9:X10')->getFont()->setBold(true);
    $sheet->getStyle('A9:X10')->getAlignment()->setWrapText(true);
    $sheet->getStyle('A9:X10')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THIN);
    $sheet->getStyle('A9:X10')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THIN);
    $sheet->getStyle('A9:X10')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THIN);
    $sheet->getStyle('A9:X10')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);

    //$sheet->getStyle('A9:X10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0FFF0');

    $workBook->setActiveSheetIndex(0);
}*/

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

