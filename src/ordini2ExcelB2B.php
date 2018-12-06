<?php
	//@ini_set('memory_limit','8192M');

	require '../vendor/autoload.php';
	// leggo i dati da un file
    //$request = file_get_contents('../examples/ordiniB2B.json');
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
        ->setSubject("Ordine Acquisto B2B")
        ->setDescription("Esportazione Righe Ordini di Acquisto B2B")
        ->setKeywords("office 2007 openxml php")
        ->setCategory("SM B2B Docs");

	$workBook->createSheet();
    $sheet = $workBook->setActiveSheetIndex(0); // la numerazione dei worksheet parte da 0
    $sheet->setTitle('test');

	$timeZone = new DateTimeZone('Europe/Rome');
	
	$tipoOrdine = [0 => 'GIORNALIERO', 1 => 'STOCK', 2 => 'DROP SHIPPING'];
	$integerFormat = '###,###,##0;[Red][<0]-###,###,##0;';
	$currencyFormat = '###,###,##0.00;[Red][<0]-###,###,##0.00;';
	
	$highestRowIndex = 1;
    $highestColumnIndex = 28;
	
	// testata
	// --------------------------------------------------------------------------------
	$sheet->setCellValueByColumnAndRow(1, 1, strtoupper('Cliente'));
	$sheet->setCellValueByColumnAndRow(2, 1, strtoupper('Numero'));
	$sheet->setCellValueByColumnAndRow(3, 1, strtoupper('Rif.'));
	$sheet->setCellValueByColumnAndRow(4, 1, strtoupper('Data Ord.'));
	$sheet->setCellValueByColumnAndRow(5, 1, strtoupper('Data Comp.'));
	$sheet->setCellValueByColumnAndRow(6, 1, strtoupper('Tipo'));
	$sheet->setCellValueByColumnAndRow(7, 1, strtoupper('Bozza'));
	$sheet->setCellValueByColumnAndRow(8, 1, strtoupper('Elim.'));
	$sheet->setCellValueByColumnAndRow(9, 1, strtoupper('Codice'));
	$sheet->setCellValueByColumnAndRow(10, 1, strtoupper('Cod. GCC'));
	$sheet->setCellValueByColumnAndRow(11, 1, strtoupper('Barcode'));
	$sheet->setCellValueByColumnAndRow(12, 1, strtoupper('Descrizione'));
	$sheet->setCellValueByColumnAndRow(13, 1, strtoupper('Modello'));
	$sheet->setCellValueByColumnAndRow(14, 1, strtoupper('Marca'));
	$sheet->setCellValueByColumnAndRow(15, 1, strtoupper('Giac.'));
	$sheet->setCellValueByColumnAndRow(16, 1, strtoupper('In Ord.'));
	$sheet->setCellValueByColumnAndRow(17, 1, strtoupper('Netto'));
	$sheet->setCellValueByColumnAndRow(18, 1, strtoupper('Q.ta Ord.'));
	$sheet->setCellValueByColumnAndRow(19, 1, strtoupper('Q.ta Conf.'));
	$sheet->setCellValueByColumnAndRow(20, 1, strtoupper('Q.ta Evasa'));
	$sheet->setCellValueByColumnAndRow(21, 1, strtoupper('Prezzo'));
	$sheet->setCellValueByColumnAndRow(22, 1, strtoupper('Tot. Ord.'));
	$sheet->setCellValueByColumnAndRow(23, 1, strtoupper('Tot. Marg. Ord.'));
	$sheet->setCellValueByColumnAndRow(24, 1, strtoupper('Tot. Ev.'));
	$sheet->setCellValueByColumnAndRow(25, 1, strtoupper('Tot. Marg. Ev.'));
	$sheet->setCellValueByColumnAndRow(26, 1, strtoupper('DDT Data'));
	$sheet->setCellValueByColumnAndRow(27, 1, strtoupper('DDT Num.'));
	$sheet->setCellValueByColumnAndRow(28, 1, strtoupper('Note'));
	
	$headerRowCount = 1;
	
	foreach ($data['righe'] as $rowNum => $riga) {
		$row = $rowNum + $headerRowCount + 1;
		
		$sheet->getCellByColumnAndRow(1, $row)->setValueExplicit(strtoupper($riga['codiceCliente']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(2, $row)->setValueExplicit(strtoupper($riga['numero']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(3, $row)->setValueExplicit(strtoupper($riga['riferimentoCliente']),DataType::TYPE_STRING);
		
		$dataOrdine = new \DateTime($riga['data']);
		$sheet->setCellValueByColumnAndRow(4, $row, Date::PHPToExcel($dataOrdine->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getCellByColumnAndRow(4, $row)->getStyle()->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
		
		$dataCompetenza = new \DateTime($riga['dataCompetenza']);
		$sheet->setCellValueByColumnAndRow(5, $row, Date::PHPToExcel($dataCompetenza->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getCellByColumnAndRow(5, $row)->getStyle()->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
		
		$sheet->getCellByColumnAndRow(6, $row)->setValueExplicit($tipoOrdine[$riga['tipo']],DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(7, $row)->setValueExplicit($riga['bozza'] == true ? 'B' : '' ,DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(8, $row)->setValueExplicit($riga['eliminato'] == true ? 'E' : '' ,DataType::TYPE_STRING);
		
		$sheet->getCellByColumnAndRow(9, $row)->setValueExplicit(strtoupper($riga['codiceArticolo']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(10, $row)->setValueExplicit(strtoupper($riga['codiceArticoloGCC']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(11, $row)->setValueExplicit(strtoupper($riga['barcode']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(12, $row)->setValueExplicit(strtoupper($riga['descrizione']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(13, $row)->setValueExplicit(strtoupper($riga['modello']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(14, $row)->setValueExplicit(strtoupper($riga['marchio']),DataType::TYPE_STRING);
		
		$sheet->getCellByColumnAndRow(15, $row)->setValueExplicit($riga['giacenza'],DataType::TYPE_NUMERIC);
		$sheet->getCellByColumnAndRow(16, $row)->setValueExplicit($riga['inOrdine'],DataType::TYPE_NUMERIC);
		$sheet->getCellByColumnAndRow(17, $row)->setValueExplicit($riga['nettoNetto'],DataType::TYPE_NUMERIC);
		$sheet->getCellByColumnAndRow(18, $row)->setValueExplicit($riga['quantita'],DataType::TYPE_NUMERIC);
		$sheet->getCellByColumnAndRow(19, $row)->setValueExplicit($riga['quantitaConfermata'],DataType::TYPE_NUMERIC);
		$sheet->getCellByColumnAndRow(20, $row)->setValueExplicit($riga['quantitaEvasa'],DataType::TYPE_NUMERIC);
		$sheet->getCellByColumnAndRow(21, $row)->setValueExplicit($riga['prezzo'],DataType::TYPE_NUMERIC);
		
		$formula = '='.Coordinate::stringFromColumnIndex(18)."$row".'*'.Coordinate::stringFromColumnIndex(21)."$row";
		$sheet->getCellByColumnAndRow(22, $row)->setValueExplicit($formula,DataType::TYPE_FORMULA);
		$formula = '='.Coordinate::stringFromColumnIndex(18)."$row".'*('.Coordinate::stringFromColumnIndex(21)."$row".'-'.Coordinate::stringFromColumnIndex(17)."$row)";
		$sheet->getCellByColumnAndRow(23, $row)->setValueExplicit($formula,DataType::TYPE_FORMULA);
		$formula = '='.Coordinate::stringFromColumnIndex(20)."$row".'*'.Coordinate::stringFromColumnIndex(21)."$row";
		$sheet->getCellByColumnAndRow(24, $row)->setValueExplicit($formula,DataType::TYPE_FORMULA);
		$formula = '='.Coordinate::stringFromColumnIndex(20)."$row".'*('.Coordinate::stringFromColumnIndex(21)."$row".'-'.Coordinate::stringFromColumnIndex(17)."$row)";
		$sheet->getCellByColumnAndRow(25, $row)->setValueExplicit($formula,DataType::TYPE_FORMULA);

		$dataDDT = new \DateTime($riga['ddtData']);
		$sheet->setCellValueByColumnAndRow(26, $row, Date::PHPToExcel($dataDDT->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getCellByColumnAndRow(26, $row)->getStyle()->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
		
		$sheet->getCellByColumnAndRow(27, $row)->setValueExplicit(strtoupper($riga['ddtNumero']),DataType::TYPE_STRING);
		$sheet->getCellByColumnAndRow(28, $row)->setValueExplicit(strtoupper($riga['note']),DataType::TYPE_STRING);
	}
	
	$highestRowIndex = $row + 1;
	
	// formattazione colonne	
	$sheet->getStyle(Coordinate::stringFromColumnIndex(4).'1:'.Coordinate::stringFromColumnIndex(5)."$highestRowIndex")->getAlignment()->setHorizontal('center');
	$sheet->getStyle(Coordinate::stringFromColumnIndex(7).'1:'.Coordinate::stringFromColumnIndex(11)."$highestRowIndex")->getAlignment()->setHorizontal('center');
	$sheet->getStyle(Coordinate::stringFromColumnIndex(15).'1:'.Coordinate::stringFromColumnIndex(16)."$highestRowIndex")->getNumberFormat()->setFormatCode($integerFormat);
	$sheet->getStyle(Coordinate::stringFromColumnIndex(15).'1:'.Coordinate::stringFromColumnIndex(16)."$highestRowIndex")->getAlignment()->setHorizontal('center');
	$sheet->getStyle(Coordinate::stringFromColumnIndex(17).'1:'.Coordinate::stringFromColumnIndex(17)."$highestRowIndex")->getNumberFormat()->setFormatCode($currencyFormat);
	$sheet->getStyle(Coordinate::stringFromColumnIndex(18).'1:'.Coordinate::stringFromColumnIndex(20)."$highestRowIndex")->getNumberFormat()->setFormatCode($integerFormat);
	$sheet->getStyle(Coordinate::stringFromColumnIndex(18).'1:'.Coordinate::stringFromColumnIndex(20)."$highestRowIndex")->getAlignment()->setHorizontal('center');
	$sheet->getStyle(Coordinate::stringFromColumnIndex(21).'1:'.Coordinate::stringFromColumnIndex(25)."$highestRowIndex")->getNumberFormat()->setFormatCode($currencyFormat);
	$sheet->getStyle(Coordinate::stringFromColumnIndex(26).'1:'.Coordinate::stringFromColumnIndex(26)."$highestRowIndex")->getAlignment()->setHorizontal('center');
	        
	$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex($highestColumnIndex).'1')->getFont()->setBold(true);
	$sheet->getStyle(Coordinate::stringFromColumnIndex(1).'1:'.Coordinate::stringFromColumnIndex($highestColumnIndex).'1')->getAlignment()->setHorizontal('center');
	
	for ($i = 1;$i <= $highestColumnIndex; $i++) $sheet->getColumnDimension(Coordinate::stringFromColumnIndex($i))->setAutoSize(true);
	
	$sheet->freezePane('A2');
	
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

