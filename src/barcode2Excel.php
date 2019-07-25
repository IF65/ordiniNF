<?php
//@ini_set('memory_limit','8192M');

require '../vendor/autoload.php';
// leggo i dati da un file
//$request = file_get_contents('../examples/ordini.json');

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

$righe = $data['righe'];
$ordinamento = array();

// creazione del workbook
$workBook = new Spreadsheet();
$workBook->getDefaultStyle()->getFont()->setName('Arial');
$workBook->getDefaultStyle()->getFont()->setSize(12);
$workBook->getProperties()
    ->setCreator("IF65 S.p.A. (Gruppo Italmark)")
    ->setLastModifiedBy("IF65 S.p.A.")
    ->setTitle("Caricamento Barcode")
    ->setSubject("Caricamento Barcode")
    ->setDescription("Caricamento Barcode")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("IF65 Docs");

// creazione degli Sheet (uno per ogni ordine)
$workBook->createSheet();
$sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
$sheet->setTitle('Barcode');

$timeZone = new DateTimeZone('Europe/Rome');

// testata colonne
// --------------------------------------------------------------------------------
$sheet->setCellValue('A1', strtoupper('Cod.Art.Forn.'));
$sheet->setCellValue('B1', strtoupper('Cod.Art.'));
$sheet->setCellValue('C1', strtoupper('Descrizione'));
$sheet->setCellValue('D1', strtoupper('Taglia'));
$sheet->setCellValue('E1', strtoupper('Barcode'));
$sheet->setCellValue('F1', strtoupper('Stato'));
$sheet->setCellValue('G1', strtoupper('Caricato'));
$sheet->setCellValue('H1', strtoupper('Note'));

// scrittura righe
// --------------------------------------------------------------------------------
$primaRigaDati = 2; // attenzione le righe in Excel partono da 1

for ($i = 0; $i < count($righe); $i++) {
    $R = ($i+$primaRigaDati);

    $stato = '';
    if ($righe[$i]['stato']) {
        $stato = 'OK';
    }

    $caricato = 'No';
    if ($righe[$i]['caricato']) {
        $caricato = 'Sì';
    }

    /// righe
    $sheet->getCell('A'.$R)->setValueExplicit($righe[$i]['codiceArticoloFornitore'],DataType::TYPE_STRING);
    $sheet->getCell('B'.$R)->setValueExplicit($righe[$i]['codiceArticolo'],DataType::TYPE_STRING);
    $sheet->getCell('C'.$R)->setValueExplicit($righe[$i]['descrizione'],DataType::TYPE_STRING);
    $sheet->getCell('D'.$R)->setValueExplicit($righe[$i]['taglia'],DataType::TYPE_STRING);
    $sheet->getCell('E'.$R)->setValueExplicit($righe[$i]['barcode'],DataType::TYPE_STRING);
    $sheet->getCell('F'.$R)->setValueExplicit($stato,DataType::TYPE_STRING);
    $sheet->getCell('G'.$R)->setValueExplicit($caricato,DataType::TYPE_STRING);
    $sheet->getCell('H'.$R)->setValueExplicit($righe[$i]['note'],DataType::TYPE_STRING);
}

// formattazione
// --------------------------------------------------------------------------------
$sheet->getDefaultRowDimension()->setRowHeight(20);
$sheet->setShowGridlines(true);


// larghezza colonne (non uso volutamente autowidth)
$sheet->getColumnDimension('A')->setWidth(25);
$sheet->getColumnDimension('B')->setWidth(15);
$sheet->getColumnDimension('C')->setWidth(40);
$sheet->getColumnDimension('D')->setWidth(10);
$sheet->getColumnDimension('E')->setWidth(15);
$sheet->getColumnDimension('F')->setWidth(14);
$sheet->getColumnDimension('G')->setWidth(10);
$sheet->getColumnDimension('H')->setWidth(50);

$workBook->setActiveSheetIndex(0);

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

