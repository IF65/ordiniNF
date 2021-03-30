<?php
//@ini_set('memory_limit','8192M');

require '../vendor/autoload.php';
// leggo i dati da un file
//$request = file_get_contents('/Users/if65/Desktop/dati.json');
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
USE PhpOffice\PhpSpreadsheet\Worksheet\PageSetup;


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
    ->setTitle("Ordine Acquisto")
    ->setSubject("Ordine Acquisto")
    ->setDescription("Esportazione Ordine di Acquisto")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("SM Docs");

$sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
$sheet->setTitle('Report Venditori');

$timeZone = new DateTimeZone('Europe/Rome');

$dataInizio = new \DateTime($data['data_inizio']);
$dataFine = new \DateTime($data['data_fine']);

// riquadro di testata
// --------------------------------------------------------------------------------
$sheet->setCellValue('A1', strtoupper('Sede: '. $data['negozio']));
$sheet->setCellValue('A2', 'DALLA DATA: '.$dataInizio->format('d/m/Y'));
$sheet->setCellValue('A3', 'ALLA DATA: '.$dataFine->format('d/m/Y'));

// testata colonne
// --------------------------------------------------------------------------------
$sheet->setCellValue('A4', strtoupper('VENDITORE'));
$sheet->mergeCells('A4:A5');
$sheet->setCellValue('B4', strtoupper('INCASSO'));
$sheet->mergeCells('B4:B5');
$sheet->setCellValue('C4', strtoupper('MARGINE M0'));
$sheet->mergeCells('C4:C5');
$sheet->setCellValue('D4', strtoupper('MARG. M0'));
$sheet->mergeCells('D4:D5');
$sheet->setCellValue('E4', strtoupper('VENDUTO IN PROMO'));
$sheet->mergeCells('E4:E5');
$sheet->setCellValue('F4', strtoupper('VEND. PROMO'));
$sheet->mergeCells('F4:F5');
$sheet->setCellValue('G4', strtoupper('VENDUTO SCONTATO'));
$sheet->mergeCells('G4:G5');
$sheet->setCellValue('H4', strtoupper('VEND. SCONT.'));
$sheet->mergeCells('H4:H5');
$sheet->setCellValue('I4', strtoupper('EST. GARANZIA'));
$sheet->mergeCells('I4:I5');
$sheet->setCellValue('J4', strtoupper('PESO ESTENDO'));
$sheet->mergeCells('J4:J5');
$sheet->setCellValue('K4', strtoupper('NR. SERVIZI'));
$sheet->mergeCells('K4:K5');
$sheet->setCellValue('L4', strtoupper('SCONTRINI'));
$sheet->mergeCells('L4:L5');
$sheet->setCellValue('M4', strtoupper('SCONTR. MEDIO'));
$sheet->mergeCells('M4:M5');
$sheet->setCellValue('N4', strtoupper('PEZZI'));
$sheet->mergeCells('N4:N5');
$sheet->setCellValue('O4', strtoupper('PEZZI PER SCONTR.'));
$sheet->mergeCells('O4:O5');
$sheet->setCellValue('P4', strtoupper('ELIMINATI'));
$sheet->mergeCells('P4:P5');
$sheet->setCellValue('Q4', strtoupper('ELIMINATI'));
$sheet->mergeCells('Q4:Q5');
$sheet->setCellValue('R4', strtoupper('MOVIMENTO'));
$sheet->mergeCells('R4:R5');
$sheet->setCellValue('S4', strtoupper('MOVIMENTO'));
$sheet->mergeCells('S4:S5');
$sheet->setCellValue('T4', strtoupper('SPETTACOLO'));
$sheet->mergeCells('T4:T5');
$sheet->setCellValue('U4', strtoupper('SPETTACOLO'));
$sheet->mergeCells('U4:U5');
$sheet->setCellValue('V4', strtoupper('G.P.B.'));
$sheet->mergeCells('V4:Y4');
$sheet->setCellValue('Z4', strtoupper('CREATIVITA\''));
$sheet->mergeCells('Z4:Z5');
$sheet->setCellValue('AA4', strtoupper('CREATIVITA\''));
$sheet->mergeCells('AA4:AA5');
$sheet->setCellValue('AB4', strtoupper('DIVERTIMENTO'));
$sheet->mergeCells('AB4:AB5');
$sheet->setCellValue('AC4', strtoupper('DIVERTIMENTO'));
$sheet->mergeCells('AC4:AC5');

$sheet->setCellValue('W5', strtoupper('G.E.D.'));
$sheet->setCellValue('Y5', strtoupper('P.E.D.'));

$r = 6;
foreach($data['dati'] as $datiVenditore) {
	$sheet->getColumnDimension('A')->setWidth(30);
	$sheet->getCell("A$r")->setValueExplicit(strtoupper($datiVenditore['venditore']),DataType::TYPE_STRING);

	$sheet->getColumnDimension('B')->setWidth(14);
	$sheet->getCell("B$r")->setValueExplicit($datiVenditore['venduto'],DataType::TYPE_NUMERIC);

	$formula="=IF(B$r=0,0,C$r/B$r)";
	$sheet->getColumnDimension('C')->setWidth(0);
	$sheet->getCell("C$r")->setValueExplicit($datiVenditore['margine'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('D')->setWidth(14);
	$sheet->getCell("D$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,E$r/B$r)";
	$sheet->getColumnDimension('E')->setWidth(0);
	$sheet->getCell("E$r")->setValueExplicit($datiVenditore['promo'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('F')->setWidth(14);
	$sheet->getCell("F$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,G$r/B$r)";
	$sheet->getColumnDimension('G')->setWidth(0);
	$sheet->getCell("G$r")->setValueExplicit($datiVenditore['sconto'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('H')->setWidth(14);
	$sheet->getCell("H$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,I$r/B$r)";
	$sheet->getColumnDimension('I')->setWidth(0);
	$sheet->getCell("I$r")->setValueExplicit($datiVenditore['estendo'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('J')->setWidth(14);
	$sheet->getCell("J$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$sheet->getCell("K$r")->setValueExplicit($datiVenditore['servizi'],DataType::TYPE_NUMERIC);

	$formula="=IF(L$r=0,0,B$r/L$r)";
	$sheet->getColumnDimension('L')->setWidth(0);
	$sheet->getCell("L$r")->setValueExplicit($datiVenditore['scontrini'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('M')->setWidth(14);
	$sheet->getCell("M$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(L$r=0,0,N$r/L$r)";
	$sheet->getColumnDimension('N')->setWidth(0);
	$sheet->getCell("N$r")->setValueExplicit($datiVenditore['pezzi'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('O')->setWidth(14);
	$sheet->getCell("O$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,P$r/B$r)";
	$sheet->getColumnDimension('P')->setWidth(0);
	$sheet->getCell("P$r")->setValueExplicit($datiVenditore['eliminati'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('Q')->setWidth(14);
	$sheet->getCell("Q$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,R$r/B$r)";
	$sheet->getColumnDimension('R')->setWidth(0);
	$sheet->getCell("R$r")->setValueExplicit($datiVenditore['movimento'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('S')->setWidth(14);
	$sheet->getCell("S$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,T$r/B$r)";
	$sheet->getColumnDimension('T')->setWidth(0);
	$sheet->getCell("T$r")->setValueExplicit($datiVenditore['spettacolo'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('U')->setWidth(14);
	$sheet->getCell("U$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,V$r/B$r)";
	$sheet->getColumnDimension('V')->setWidth(0);
	$sheet->getCell("V$r")->setValueExplicit($datiVenditore['ged'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('W')->setWidth(14);
	$sheet->getCell("W$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,X$r/B$r)";
	$sheet->getColumnDimension('X')->setWidth(0);
	$sheet->getCell("X$r")->setValueExplicit($datiVenditore['ped'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('Y')->setWidth(14);
	$sheet->getCell("Y$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,Z$r/B$r)";
	$sheet->getColumnDimension('Z')->setWidth(0);
	$sheet->getCell("Z$r")->setValueExplicit($datiVenditore['creativita'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('AA')->setWidth(14);
	$sheet->getCell("AA$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

	$formula="=IF(B$r=0,0,AB$r/B$r)";
	$sheet->getColumnDimension('AB')->setWidth(0);
	$sheet->getCell("AB$r")->setValueExplicit($datiVenditore['divertimento'],DataType::TYPE_NUMERIC);
	$sheet->getColumnDimension('AC')->setWidth(14);
	$sheet->getCell("AC$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
	$r++;
}
// totali
$t = $r - 1;
$sheet->getCell("A$r")->setValueExplicit('TOTALI',DataType::TYPE_STRING);

$formula="=SUM(B6:B$r)";
$sheet->getCell("B$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(C6:C$r)";
$sheet->getCell("C$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,C$r/B$r)";
$sheet->getCell("D$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(E6:E$r)";
$sheet->getCell("E$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,E$r/B$r)";
$sheet->getCell("F$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(G6:G$r)";
$sheet->getCell("G$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,G$r/B$r)";
$sheet->getCell("H$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(I6:I$r)";
$sheet->getCell("I$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,I$r/B$r)";
$sheet->getCell("J$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(K6:K$r)";
$sheet->getCell("K$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(L6:L$r)";
$sheet->getCell("L$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(L$r=0,0,B$r/L$r)";
$sheet->getCell("M$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(N6:N$r)";
$sheet->getCell("N$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(L$r=0,0,N$r/L$r)";
$sheet->getCell("O$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(P6:P$r)";
$sheet->getCell("P$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,P$r/B$r)";
$sheet->getCell("Q$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(R6:R$r)";
$sheet->getCell("R$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,R$r/B$r)";
$sheet->getCell("S$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(T6:T$r)";
$sheet->getCell("T$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,T$r/B$r)";
$sheet->getCell("U$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(V6:V$r)";
$sheet->getCell("V$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,V$r/B$r)";
$sheet->getCell("W$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(X6:X$r)";
$sheet->getCell("X$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,X$r/B$r)";
$sheet->getCell("Y$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(Z6:Z$r)";
$sheet->getCell("Z$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,Z$r/B$r)";
$sheet->getCell("AA$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

$formula="=SUM(AB6:AB$r)";
$sheet->getCell("AB$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
$formula="=IF(B$r=0,0,AB$r/B$r)";
$sheet->getCell("AC$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

/*$formula="=IF(B$r=0,0,C$r/B$r)";
$sheet->getColumnDimension('C')->setWidth(0);
$sheet->getCell("C$r")->setValueExplicit($datiVenditore['margine'],DataType::TYPE_NUMERIC);
$sheet->getColumnDimension('D')->setWidth(14);
$sheet->getCell("D$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);*/


// formattazione
// --------------------------------------------------------------------------------
$lastColumn = $sheet->getHighestColumn();
$lastRow = $sheet->getHighestRow();
$sheet->getDefaultRowDimension()->setRowHeight(40);
$sheet->setShowGridlines(true);
$sheet->getStyle("A1:$lastColumn$lastRow")->getAlignment()->setVertical('center');

// riquadro di testata
$sheet->getStyle('A1:A4')->getFont()->setBold(true);
$sheet->getStyle('A4:AE5')->getFont()->setBold(true);
$sheet->getStyle('A4:AE5')->getAlignment()->setVertical('center')->setHorizontal('center')->setWrapText(true);
$sheet->getStyle("A$r:AC$r")->getFont()->setBold(true);

// colonne
$sheet->getStyle("B6:B$r")->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00;');
$sheet->getStyle("D6:D$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("F6:F$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("H6:H$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("J6:J$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("M6:M$r")->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00;');
$sheet->getStyle("O6:O$r")->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00;');
$sheet->getStyle("Q6:Q$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("S6:S$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("U6:U$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("W6:W$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("Y6:Y$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("AA6:AA$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
$sheet->getStyle("AC6:AC$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');

$sheet->getStyle("A4:AC4")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
$sheet->getStyle("A5:AC5")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
/*$sheet->getStyle('A9:X10')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THIN);
$sheet->getStyle('A9:X10')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THIN);
$sheet->getStyle('A9:X10')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);*/
$sheet->getStyle("A$r:AC$r")->getBorders()->getAllBorders()->setBorderStyle(Border::BORDER_THIN);
//$sheet->getStyle('A9:X10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0FFF0')

$sheet->getPageSetup()->setOrientation(PageSetup::PAPERSIZE_A4);
$sheet->getPageSetup()->setOrientation(PageSetup::ORIENTATION_LANDSCAPE);
$sheet->getPageSetup()->setFitToWidth(1);
$sheet->getPageSetup()->setFitToHeight(0);

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

