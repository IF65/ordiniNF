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
	$sheet->setCellValue('D4', strtoupper('MARG. M0 %'));
	$sheet->mergeCells('D4:D5');
	$sheet->setCellValue('E4', strtoupper('VENDUTO IN PROMO'));
	$sheet->mergeCells('E4:E5');
	$sheet->setCellValue('F4', strtoupper('VEND. PROMO %'));
	$sheet->mergeCells('F4:F5');
	$sheet->setCellValue('G4', strtoupper('VENDUTO SCONTATO'));
	$sheet->mergeCells('G4:G5');
	$sheet->setCellValue('H4', strtoupper('VEND. SCONT. %'));
	$sheet->mergeCells('H4:H5');
	$sheet->setCellValue('I4', strtoupper('EST. GARANZIA'));
	$sheet->mergeCells('I4:I5');
	$sheet->setCellValue('J4', strtoupper('PESO ESTENDO'));
	$sheet->mergeCells('J4:J5');
	$sheet->setCellValue('K4', strtoupper('NR. SERVIZI'));
	$sheet->mergeCells('K4:K5');
	$sheet->setCellValue('L4', strtoupper('SCONTRINI'));
	$sheet->mergeCells('L4:L5');
	$sheet->setCellValue('M4', strtoupper('SCONTRINO MEDIO'));
	$sheet->mergeCells('M4:M5');
	$sheet->setCellValue('N4', strtoupper('PEZZI'));
	$sheet->mergeCells('N4:N5');
	$sheet->setCellValue('O4', strtoupper('PEZZI PER SCONTR.'));
	$sheet->mergeCells('O4:O5');
	$sheet->setCellValue('P4', strtoupper('ELIMINATI'));
	$sheet->mergeCells('P4:P5');
	$sheet->setCellValue('Q4', strtoupper('ELIMINATI %'));
	$sheet->mergeCells('Q4:Q5');
	$sheet->setCellValue('R4', strtoupper('MOVIMENTO'));
	$sheet->mergeCells('R4:R5');
	$sheet->setCellValue('S4', strtoupper('MOVIMENTO %'));
	$sheet->mergeCells('S4:S5');
	$sheet->setCellValue('T4', strtoupper('SPETTACOLO'));
	$sheet->mergeCells('T4:T5');
	$sheet->setCellValue('U4', strtoupper('SPETTACOLO %'));
	$sheet->mergeCells('U4:U5');
	$sheet->setCellValue('V4', strtoupper('G.P.B.'));
	$sheet->mergeCells('V4:AA4');
	$sheet->setCellValue('AB4', strtoupper('CREATIVITA\''));
	$sheet->mergeCells('AB4:AB5');
	$sheet->setCellValue('AC4', strtoupper('CREATIVITA\' %'));
	$sheet->mergeCells('AC4:AC5');
	$sheet->setCellValue('AD4', strtoupper('DIVERTIMENTO'));
	$sheet->mergeCells('AD4:AD5');
	$sheet->setCellValue('AE4', strtoupper('DIVERTIMENTO %'));
	$sheet->mergeCells('AE4:AE5');

	$sheet->setCellValue('W5', strtoupper('G.E.D.'));
	$sheet->setCellValue('Y5', strtoupper('P.E.D.'));
	$sheet->setCellValue('AA5', strtoupper('ALTRO'));

	$r = 6;
	foreach($data['dati'] as $datiVenditore) {
		$sheet->getColumnDimension('A')->setWidth(30);
		$sheet->getCell("A$r")->setValueExplicit(strtoupper($datiVenditore['venditore']),DataType::TYPE_STRING);

		$sheet->getColumnDimension('B')->setWidth(12);
		$sheet->getCell("B$r")->setValueExplicit($datiVenditore['venduto'],DataType::TYPE_NUMERIC);

		$formula="=IF(B$r=0,0,C$r/B$r)";
		$sheet->getColumnDimension('C')->setWidth(0);
		$sheet->getCell("C$r")->setValueExplicit($datiVenditore['margine'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('D')->setWidth(10);
		$sheet->getCell("D$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,E$r/B$r)";
		$sheet->getColumnDimension('E')->setWidth(0);
		$sheet->getCell("E$r")->setValueExplicit($datiVenditore['promo'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('F')->setWidth(10);
		$sheet->getCell("F$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,G$r/B$r)";
		$sheet->getColumnDimension('G')->setWidth(0);
		$sheet->getCell("G$r")->setValueExplicit($datiVenditore['sconto'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('H')->setWidth(10);
		$sheet->getCell("H$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,I$r/B$r)";
		$sheet->getColumnDimension('I')->setWidth(0);
		$sheet->getCell("I$r")->setValueExplicit($datiVenditore['estendo'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('JD')->setWidth(10);
		$sheet->getCell("J$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$sheet->getCell("K$r")->setValueExplicit($datiVenditore['servizi'],DataType::TYPE_NUMERIC);

		$formula="=IF(B$r=0,0,L$r/B$r)";
		$sheet->getColumnDimension('L')->setWidth(0);
		$sheet->getCell("L$r")->setValueExplicit($datiVenditore['scontrini'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('M')->setWidth(10);
		$sheet->getCell("M$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(N$r=0,0,L$r/N$r)";
		$sheet->getColumnDimension('N')->setWidth(0);
		$sheet->getCell("N$r")->setValueExplicit($datiVenditore['pezzi'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('O')->setWidth(10);
		$sheet->getCell("O$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,P$r/B$r)";
		$sheet->getColumnDimension('P')->setWidth(0);
		$sheet->getCell("P$r")->setValueExplicit($datiVenditore['eliminati'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('Q')->setWidth(10);
		$sheet->getCell("Q$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,R$r/B$r)";
		$sheet->getColumnDimension('R')->setWidth(0);
		$sheet->getCell("R$r")->setValueExplicit($datiVenditore['movimento'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('S')->setWidth(10);
		$sheet->getCell("S$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,T$r/B$r)";
		$sheet->getColumnDimension('T')->setWidth(0);
		$sheet->getCell("T$r")->setValueExplicit($datiVenditore['spettacolo'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('U')->setWidth(10);
		$sheet->getCell("U$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,V$r/B$r)";
		$sheet->getColumnDimension('V')->setWidth(0);
		$sheet->getCell("V$r")->setValueExplicit($datiVenditore['ged'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('W')->setWidth(10);
		$sheet->getCell("W$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,X$r/B$r)";
		$sheet->getColumnDimension('X')->setWidth(0);
		$sheet->getCell("X$r")->setValueExplicit($datiVenditore['ped'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('Y')->setWidth(10);
		$sheet->getCell("Y$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,Z$r/B$r)";
		$sheet->getColumnDimension('Z')->setWidth(0);
		$sheet->getCell("Z$r")->setValueExplicit($datiVenditore['gpb']-$datiVenditore['ped']-$datiVenditore['ged'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('AA')->setWidth(10);
		$sheet->getCell("AA$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,AB$r/B$r)";
		$sheet->getColumnDimension('AB')->setWidth(0);
		$sheet->getCell("AB$r")->setValueExplicit($datiVenditore['creativita'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('AC')->setWidth(10);
		$sheet->getCell("AC$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$formula="=IF(B$r=0,0,AD$r/B$r)";
		$sheet->getColumnDimension('AD')->setWidth(0);
		$sheet->getCell("AD$r")->setValueExplicit($datiVenditore['divertimento'],DataType::TYPE_NUMERIC);
		$sheet->getColumnDimension('D')->setWidth(9);
		$sheet->getCell("AE$r")->setValueExplicit($formula,DataType::TYPE_FORMULA);
		$r++;
	}

    // formattazione
    // --------------------------------------------------------------------------------
    $sheet->getDefaultRowDimension()->setRowHeight(20);
    $sheet->setShowGridlines(true);

    // riquadro di testata
    $sheet->getStyle('A1:A4')->getFont()->setBold(true);
    $sheet->getStyle('A4:AE5')->getFont()->setBold(true);
	$sheet->getStyle('A4:AE5')->getAlignment()->setHorizontal('center')->setWrapText(true);

	// colonne
	$sheet->getStyle("B6:B$r")->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00;');
	$sheet->getStyle("D6:D$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("F6:F$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("H6:H$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("J6:J$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("M6:M$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("O6:O$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("Q6:Q$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("S6:S$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("U6:U$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("W6:W$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("Y6:Y$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("AA6:AA$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("AC6:AC$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
	$sheet->getStyle("AE6:AE$r")->getNumberFormat()->setFormatCode('#0.00%;[Red][<0]-#0.00%;');
//foreach (range('A','X') as $col) {$sheet->getColumnDimension($col)->setAutoSize(true);}
/*
        // colonne descrizione articolo + prezzi
         $sheet->getStyle(sprintf("%s%s%s%s%s",'B',$primaRigaDati,':','C',$primaRigaDati+count($righe)-1))->
            getAlignment()->setHorizontal('center');
        $sheet->getStyle(sprintf("%s%s%s%s%s",'G',$primaRigaDati,':','J',$primaRigaDati+count($righe)-1))->
            getAlignment()->setHorizontal('center');
        $sheet->getStyle(sprintf("%s%s%s%s%s",'L',$primaRigaDati,':','X',$primaRigaDati+count($righe)-1))->
            getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00;');

        // quantita + sconto merce
        $sheet->getStyle(sprintf("%s%s%s%s",$colQuantitaTotale,'1:',$colSCIndex['LAST'],$primaRigaDati-1))->getFont()->setBold(true);
        $sheet->getStyle($colQuantitaTotale.'1')->getAlignment()->setHorizontal('center')->setVertical('center');
        $sheet->getStyle($colScontoMerceTotale.'1')->getAlignment()->setHorizontal('center')->setVertical('center');
        $sheet->getStyle(sprintf("%s%s%s%s%s",$colQuantitaTotale,$primaRigaDati,':',$colSCIndex['LAST'],$primaRigaDati+count($righe)-1))->
            getAlignment()->setHorizontal('center');

        // larghezza colonne (non uso volutamente autowidth)

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
        /*$sheet->getStyle('A9:X10')->getBorders()->getLeft()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle('A9:X10')->getBorders()->getTop()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle('A9:X10')->getBorders()->getRight()->setBorderStyle(Border::BORDER_THIN);
        $sheet->getStyle('A9:X10')->getBorders()->getBottom()->setBorderStyle(Border::BORDER_THIN);*/

        //$sheet->getStyle('A9:X10')->getFill()->setFillType(Fill::FILL_SOLID)->getStartColor()->setARGB('FFF0FFF0');

        $workBook->setActiveSheetIndex(0);


    $writer = new Xlsx($workBook);
    $writer->save($file);

    /*if (file_exists($file)) {
		header('Content-Description: File Transfer');
		header('Content-Type: application/octet-stream');
		header('Content-Disposition: attachment; filename="'.basename($file).'"');
		header('Expires: 0');
		header('Cache-Control: must-revalidate');
		header('Pragma: public');
		header('Content-Length: ' . filesize($file));
		readfile($file);
		exit;
	}*/

