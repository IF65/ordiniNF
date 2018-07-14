<?php
	//@ini_set('memory_limit','8192M');

	require '../vendor/autoload.php';
	// leggo i dati da un file
    //$request = file_get_contents('/Users/if65/Desktop/Stilnovo/ordini.json');
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
	use \PhpOffice\PhpSpreadsheet\Style\Protection;
	use PhpOffice\PhpSpreadsheet\Shared\Date;
	
	// verifico l'esistenza della cartella temp e se serve la creo
	// con mask 777.
	if (! file_exists ( '../temp' )) {
		$oldMask = umask(0);
		mkdir('../temp', 0777);
		umask($oldMask);
	}
	
	function num2alpha($n) {
        $n--;
        for($r = ""; $n >= 0; $n = intval($n / 26) - 1)
            $r = chr($n%26 + 0x41) . $r;
        return $r;
    }
	
	function RC($r = 0, $c = 0) {
		global $x, $y;
		
		$newX = $x + $c;
		$newY = $y + $r;
		return num2Alpha($newX)."$newY";
	};

    $style = new Style();

    // leggo i parametri contenuti nel file
    $nomeFile = $data['nomeFile'];
    $file = '../temp/'.$nomeFile.'.xlsx';

    $ordini = $data['ordini'];
    $ordinamento = array();
    foreach ($ordini as $key => $row) {
        $ordinamento[$key] = $row['numero'];
    }
    array_multisort($ordinamento, SORT_ASC, $ordini);

    // creazione del workbook
    $workBook = new Spreadsheet();
    $workBook->getDefaultStyle()->getFont()->setName('Arial')->setSize(12);
	$workBook->getDefaultStyle()->getAlignment()->setVertical('center')->setHorizontal('center');
    //$workBook->getDefaultStyle()->getFont();
    $workBook->getProperties()
        ->setCreator("IF65 S.p.A. (Gruppo Italmark)")
        ->setLastModifiedBy("IF65 S.p.A.")
        ->setTitle("Ordine Acquisto")
        ->setSubject("Ordine Acquisto")
        ->setDescription("Esportazione Ordine di Acquisto")
        ->setKeywords("office 2007 openxml php")
        ->setCategory("SM Docs");
		
	
		
    // creazione degli Sheet (uno per ogni ordine)
    $sheetNumber = 0;
    foreach ($ordini as $ordine) {
	    $sheetNumber++;
        if ($workBook->getSheetCount() < $sheetNumber) {
            $workBook->createSheet();
        }
        $sheet = $workBook->setActiveSheetIndex($sheetNumber-1); // la numerazione dei worksheet parte da 0
		
		// formattazione di default dello sheet
		// --------------------------------------------------------------------------------
		$sheet->getDefaultRowDimension()->setRowHeight(20);
		$sheet->setShowGridlines(true);
		//$sheet->getProtection()->setSheet(true)->setSort(False);
		
        $sheet->setTitle(preg_replace('/\//','_',$ordine['numero']));

		$timeZone = new DateTimeZone('Europe/Rome');

		$dataOrdine = new \DateTime($ordine['data']);
		$dataConsegna= new \DateTime($ordine['dataConsegna']);
		$dataConsegnaMinima= new \DateTime($ordine['dataConsegnaMinima']);
		$dataConsegnaMassima= new \DateTime($ordine['dataConsegnaMassima']);

		$filiali = $ordine['sedi'];
		$ordinamento = array();
		foreach ($filiali as $key => $row) {
			$ordinamento[$key] = $row['ordinamento'];
		}
		array_multisort($ordinamento, SORT_ASC, $filiali);
		$countFiliali = count($filiali);


        // riquadro di testata
        // --------------------------------------------------------------------------------
		$xOffset = 0;
		$yOffset = 0;
		
        $sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 1, strtoupper('fornitore'));
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 1, $ordine['fornitore']);
        $sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 1, $xOffset + 3, $yOffset + 1);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 1, strtoupper('forma di pagamento'));
        $sheet->setCellValueByColumnAndRow($xOffset + 5, $yOffset + 1, $ordine['pagamento']);
		
        $sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 2, strtoupper('numero ordine'));
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 2, $ordine['numero']);
        $sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 2, $xOffset + 3, $yOffset + 2);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 2, strtoupper('sconto cassa %'));
        $sheet->setCellValueByColumnAndRow($xOffset + 5, $yOffset + 2, $ordine['scontoCassa']);
		
		$sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 3, strtoupper('data ordine'));
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 3, Date::PHPToExcel($dataOrdine->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyleByColumnAndRow($xOffset + 2, $yOffset + 3)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
        $sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 3, $xOffset + 3, $yOffset + 3);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 3, strtoupper('spese di trasporto'));
        $sheet->setCellValueByColumnAndRow($xOffset + 5, $yOffset + 3, $ordine['speseTrasporto']);
		
		$sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 4, strtoupper('data consegna prevista'));
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 4, Date::PHPToExcel($dataConsegna->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyleByColumnAndRow($xOffset + 2, $yOffset + 4)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
        $sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 4, $xOffset + 3, $yOffset + 4);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 4, strtoupper('spese di trasporto %'));
        $sheet->setCellValueByColumnAndRow($xOffset + 5, $yOffset + 4, $ordine['speseTrasportoPerc']);
		
		$sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 5, strtoupper('data consegna minima'));
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 5, Date::PHPToExcel($dataConsegnaMinima->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyleByColumnAndRow($xOffset + 2, $yOffset + 5)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
        $sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 5, $xOffset + 3, $yOffset + 5);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 5, strtoupper('margine totale'));
		
		$x = $xOffset + 5;
		$y = $yOffset + 5;
		$formula = "=SUM(".RC(6, 18).":".RC(5 + count($ordine['righe'])*2, 18).")";
		$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
		$sheet->getStyleByColumnAndRow($x, $y)->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00; ');
	
		$sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 6, strtoupper('data consegna massimo'));
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 6, Date::PHPToExcel($dataConsegnaMassima->setTimezone($timeZone)->format('Y-m-d')));
		$sheet->getStyleByColumnAndRow($xOffset + 2, $yOffset + 6)->getNumberFormat()->setFormatCode(NumberFormat::FORMAT_DATE_DDMMYYYY);
        $sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 6, $xOffset + 3, $yOffset + 6);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 6, strtoupper('margine %'));
        
		$x = $xOffset + 5;
		$y = $yOffset + 6;
		$formula = "=IF(".RC(1,0)."<>0,ROUND(".RC(-1,0)."/(SUMPRODUCT(".RC(5,4).":".RC(9,4).",".RC(5,19).":".RC(9,19).")/100+".RC(1,0).")*100,2),0)";
		$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
		$sheet->getStyleByColumnAndRow($x, $y)->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00; ');
		
		$sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 7, strtoupper('buyer'));
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 7, $ordine['buyerCodice'].' - '.$ordine['buyerDescrizione']);
		$sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 7, $xOffset + 3, $yOffset + 7);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 7, strtoupper('totale ordine'));
		
		$x = $xOffset + 5;
		$y = $yOffset + 7;
		$formula = "=SUM(".RC(4, 19).":".RC(3 + count($ordine['righe'])*2, 19).")";
		$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
		$sheet->getStyleByColumnAndRow($x, $y)->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00; ');
		
		// formattazione
        $sheet->getStyleByColumnAndRow($xOffset + 1, $yOffset + 1, $xOffset + 1, $yOffset + 7)->getFont()->setBold(true);
		$sheet->getStyleByColumnAndRow($xOffset + 1, $yOffset + 1, $xOffset + 2, $yOffset + 7)->getAlignment()->setHorizontal('left');
		$sheet->getStyleByColumnAndRow($xOffset + 4, $yOffset + 1, $xOffset + 5, $yOffset + 7)->getAlignment()->setHorizontal('left');
        $sheet->getStyleByColumnAndRow($xOffset + 4, $yOffset + 1, $xOffset + 4, $yOffset + 7)->getFont()->setBold(true);

        // testata colonne
        // --------------------------------------------------------------------------------
        $sheet->setCellValueByColumnAndRow($xOffset + 1, $yOffset + 9, strtoupper('cod.art. fornitore'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 1, $yOffset + 9, $xOffset + 1, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 2, $yOffset + 9, strtoupper('ean'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 2, $yOffset + 9, $xOffset + 2, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 3, $yOffset + 9, strtoupper('cod. art.'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 3, $yOffset + 9, $xOffset + 3, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 4, $yOffset + 9, strtoupper('descrizione'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 4, $yOffset + 9, $xOffset + 4, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 5, $yOffset + 9, strtoupper('marca'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 5, $yOffset + 9, $xOffset + 5, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 6, $yOffset + 9, strtoupper('modello'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 6, $yOffset + 9, $xOffset + 6, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 7, $yOffset + 9, strtoupper('fam.'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 7, $yOffset + 9, $xOffset + 7, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 8, $yOffset + 9, strtoupper('s.fam.'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 8, $yOffset + 9, $xOffset + 8, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 9, $yOffset + 9, strtoupper('iva'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 9, $yOffset + 9, $xOffset + 10, $yOffset + 9);
        $sheet->setCellValueByColumnAndRow($xOffset + 9, $yOffset + 10,  '%');
        $sheet->setCellValueByColumnAndRow($xOffset + 10, $yOffset + 10,  'T');
        $sheet->setCellValueByColumnAndRow($xOffset + 11, $yOffset + 9, strtoupper('tg.'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 11, $yOffset + 9, $xOffset + 11, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 12, $yOffset + 9, strtoupper('costo'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 12, $yOffset + 9, $xOffset + 12, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 13, $yOffset + 9, strtoupper('sconti'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 13, $yOffset + 9, $xOffset + 18, $yOffset + 9);
        $sheet->setCellValueByColumnAndRow($xOffset + 13, $yOffset + 10,  'A');
        $sheet->setCellValueByColumnAndRow($xOffset + 14, $yOffset + 10,  'B');
        $sheet->setCellValueByColumnAndRow($xOffset + 15, $yOffset + 10,  'C');
        $sheet->setCellValueByColumnAndRow($xOffset + 16, $yOffset + 10,  'D');
        $sheet->setCellValueByColumnAndRow($xOffset + 17, $yOffset + 10,  'EXT.');
        $sheet->setCellValueByColumnAndRow($xOffset + 18, $yOffset + 10,  strtoupper('imp.'));
        $sheet->setCellValueByColumnAndRow($xOffset + 19, $yOffset + 9, strtoupper('costo finito'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 19, $yOffset + 9, $xOffset + 19, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 20, $yOffset + 9, strtoupper('prezzo vendita'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 20, $yOffset + 9, $xOffset + 20, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 21, $yOffset + 9, strtoupper('marg.'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 21, $yOffset + 9, $xOffset + 21, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 22, $yOffset + 9, strtoupper('marg.%'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 22, $yOffset + 9, $xOffset + 22, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 23, $yOffset + 9, strtoupper('marg. totale'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 23, $yOffset + 9, $xOffset + 23, $yOffset + 10);
        $sheet->setCellValueByColumnAndRow($xOffset + 24, $yOffset + 9, strtoupper('costo totale'));
        $sheet->mergeCellsByColumnAndRow($xOffset + 24, $yOffset + 9, $xOffset + 24, $yOffset + 10);

        $sheet->mergeCellsByColumnAndRow($xOffset + 1, $yOffset + 8, $xOffset + 24, $yOffset + 8);
        $sheet->mergeCellsByColumnAndRow($xOffset + 6, $yOffset + 1, $xOffset + 24, $yOffset + 7);
		
		// formattazione
		$sheet->getStyleByColumnAndRow($xOffset + 1, $yOffset + 9, $xOffset + 25, $yOffset + 10)->getFont()->setBold(true);
        $sheet->getStyleByColumnAndRow($xOffset + 1, $yOffset + 9, $xOffset + 25, $yOffset + 10)->getAlignment()->setHorizontal('center')->setWrapText(true);
		
		// testata elenco sedi
        // --------------------------------------------------------------------------------
		$sedi = $ordine['sedi'];
		$ordinamento = array();
		foreach ($sedi as $key => $row) {
			$ordinamento[$key] = $row['ordinamento'];
		}
		array_multisort($ordinamento, SORT_ASC, $sedi);
		
		$x = $xOffset + 25;
		$y =  $yOffset + 1;
		
		$sheet->getColumnDimensionByColumn($x)->setWidth(8);
        $sheet->setCellValueByColumnAndRow($x, $y, strtoupper('numero pezzi'));
		$sheet->getStyleByColumnAndRow($x, $y)->getFont()->setBold(true);
        $sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setTextRotation(90)->setHorizontal('center')->setVertical('bottom');
        $sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 7);
		
		$sheet->setCellValueByColumnAndRow($x, $y + 8, 'totale');
		$sheet->getStyleByColumnAndRow($x, $y + 8)->getFont()->setItalic(true)->setBold(false);
		$sheet->getStyleByColumnAndRow($x, $y + 8)->getAlignment()->setHorizontal('center')->setVertical('center');
		
		$sheet->setCellValueByColumnAndRow($x, $y + 9, 'vent.');
		$sheet->getStyleByColumnAndRow($x, $y + 9)->getFont()->setItalic(true)->setBold(false);
		$sheet->getStyleByColumnAndRow($x, $y + 9)->getAlignment()->setHorizontal('center')->setVertical('center');
		
		$x++;
        foreach ($sedi as $codice => $sede) {
			$sheet->getColumnDimensionByColumn($x)->setWidth(4);
			$sheet->getColumnDimensionByColumn($x + 1)->setWidth(4);
			
			$sheet->setCellValueByColumnAndRow($x, $y, strtoupper($sede['descrizione']));
			$sheet->getStyleByColumnAndRow($x, $y)->getFont()->setBold(true);
			$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setTextRotation(90)->setHorizontal('center')->setVertical('bottom');
			$sheet->mergeCellsByColumnAndRow($x, $y, $x + 1, $y + 7);
			
			$sheet->setCellValueByColumnAndRow($x, $y + 8, 'o');
			$sheet->getStyleByColumnAndRow($x, $y + 8)->getFont()->setItalic(true);
			$sheet->getStyleByColumnAndRow($x, $y + 8)->getAlignment()->setHorizontal('center')->setVertical('center');
			$sheet->setCellValueByColumnAndRow($x + 1, $y + 8, 'sm');
			$sheet->getStyleByColumnAndRow($x + 1, $y + 8)->getFont()->setItalic(true);
			$sheet->getStyleByColumnAndRow($x + 1, $y + 8)->getAlignment()->setHorizontal('center')->setVertical('center');
			$sheet->setCellValueByColumnAndRow($x, $y + 9, 'vent.');
			$sheet->getStyleByColumnAndRow($x, $y + 9)->getFont()->setItalic(true);
			$sheet->getStyleByColumnAndRow($x, $y + 9)->getAlignment()->setHorizontal('center')->setVertical('center');
			$sheet->mergeCellsByColumnAndRow($x, $y + 9, $x + 1, $y + 9);
			
			$x += 2;
        }
		
		// righe ordine
        // --------------------------------------------------------------------------------
		$firstDataRow = 11;
		
		foreach ($ordine['righe'] as $index => $riga) {		
			$y = $yOffset + $firstDataRow + ($index * 2);
			
			$x = $xOffset + 1;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['codiceArticoloFornitore'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('left');
			$sheet->getColumnDimensionByColumn($x)->setWidth(25);
			
			$x = $xOffset + 2;
			$barcode = $riga['barcode'];
			if (count($barcode)) {
				$sheet->setCellValueExplicitByColumnAndRow($x, $y, $barcode[0],DataType::TYPE_STRING,DataType::TYPE_STRING);
			}
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			//$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
			$sheet->getColumnDimensionByColumn($x)->setWidth(15);
			
			$x = $xOffset + 3;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['codice'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			//$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 4;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['descrizione'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('left')->setWrapText(true);
			$sheet->getColumnDimensionByColumn($x)->setWidth(50);
			
			$x = $xOffset + 5;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['marca'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('left')->setWrapText(true);
			$sheet->getColumnDimensionByColumn($x)->setWidth(14);
			
			$x = $xOffset + 6;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['modello'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('left')->setWrapText(true);
			$sheet->getColumnDimensionByColumn($x)->setWidth(14);
			
			$x = $xOffset + 7;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['famiglia'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			//$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
			$sheet->getColumnDimensionByColumn($x)->setWidth(10);
			
			$x = $xOffset + 8;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['sottoFamiglia'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			//$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
			$sheet->getColumnDimensionByColumn($x)->setWidth(10);
			
			$x = $xOffset + 9;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['iva'],DataType::TYPE_STRING);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			//$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
			$sheet->getColumnDimensionByColumn($x)->setWidth(4);
			
			$x = $xOffset + 10;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['tipoIva'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			//$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
			$sheet->getColumnDimensionByColumn($x)->setWidth(4);
			
			$x = $xOffset + 11;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['taglia'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			//$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 12;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['listino'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 13;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['scontoA'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(6);
			
			$x = $xOffset + 14;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['scontoB'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(6);
			
			$x = $xOffset + 15;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['scontoC'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(6);
			
			$x = $xOffset + 16;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['scontoD'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(6);
			
			$x = $xOffset + 17;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['scontoExtra'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(6);
			
			$x = $xOffset + 18;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['scontoImporto'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(6);
			
			$x = $xOffset + 19;
			$formula = "=IF(".RC(0,46)."+".RC(0,6).">0,ROUND((".RC(0,-7)."*(100-".RC(0,-6).")/100*(100-".RC(0,-5).")/100*(100-".RC(0,-4).")/100*(100-".RC(0,-3).")/100*(100-".RC(0,-2).")/100-".RC(0,-1).")*".RC(0,6)."/(".RC(0,46)."+".RC(0,6)."),2),ROUND((".RC(0,-7)."*(100-".RC(0,-6).")/100*(100-".RC(0,-5).")/100*(100-".RC(0,-4).")/100*(100-".RC(0,-3).")/100*(100-".RC(0,-2).")/100-".RC(0,-1)."),2))";
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 20;
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $riga['prezzo'],DataType::TYPE_NUMERIC);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 21;
			$formula = "=ROUND(".RC(0, -1)."*(100/(100+".RC(0, -12)."))-".RC(0, -2).",2)";
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 22;
			$formula = "=IF(".RC(0,-2)."<>0,ROUND(".RC(0,-1)."/(".RC(0,-2)."*(100/(100+".RC(0,-13).")))*100,2),0)";
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 23;
			$formula = "=ROUND(".RC(0,-2)."*".RC(0,2).",2)";
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 24;
			//$formula = "=ROUND(".RC(0,-5)."*(".RC(0,1)."-(SUMPRODUCT(--(MOD(COLUMN(".RC(0,2).":".RC(0,65).")-COLUMN(".RC(0,2).")+1,2)=0),".RC(0,2).":".RC(0,65)."))),2)";
			//$formula = "=ROUND(S$y*(Y$y-(SUMPRODUCT(--(MOD(COLUMN(Z$y:CK$y)-COLUMN(Z$y)+1,2)=0),Z$y:CK$y))),2)";
			$formula = "=ROUND(".RC(0,-5)."*(".RC(0,1);
			$offset = 3;
			foreach ($sedi as $sede) {
				$formula .= '-'.RC(0, $offset);
				$offset += 2;
			}
			$formula .= "),2)";
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
			$sheet->mergeCellsByColumnAndRow($x, $y, $x, $y + 1);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$x = $xOffset + 25;
			$formula = "=SUM(".RC(0, 1).':'.RC(0, count($sedi)*2).')';
			$sheet->setCellValueExplicitByColumnAndRow($x, $y, $formula,DataType::TYPE_FORMULA);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			$formula = "=SUM(".RC(1, 1).':'.RC(1, count($sedi)*2).')';
			$sheet->setCellValueExplicitByColumnAndRow($x, $y + 1, $formula,DataType::TYPE_FORMULA);
			$sheet->getColumnDimensionByColumn($x)->setWidth(9);
			
			// quantita
			$x++;
			foreach ($sedi as $codice => $sede) {
				foreach($riga['quantita'] as $quantita) {
					if ($quantita['sede'] == $sede['codice']) {
						$quantitaVentilata = 0;
						foreach ($quantita['ventilazione'] as $key => $value) {
							$quantitaVentilata -= $value;
						}
						
						$quantitaInOrdine = $quantita['quantita']*1;
						$quantitaInScontoMerce =$quantita['scontoMerce']*1;
						
						$sheet->setCellValueExplicitByColumnAndRow($x, $y, $quantitaInOrdine, DataType::TYPE_NUMERIC);
						$sheet->getStyleByColumnAndRow($x, $y)->getAlignment()->setHorizontal('center');
						
						$sheet->setCellValueExplicitByColumnAndRow($x + 1, $y, $quantitaInScontoMerce, DataType::TYPE_NUMERIC);
						$sheet->getStyleByColumnAndRow($x + 1, $y)->getAlignment()->setHorizontal('center');
						
						if ($quantitaVentilata != 0) {
							$sheet->setCellValueExplicitByColumnAndRow($x, $y + 1,($quantitaInOrdine + $quantitaInScontoMerce) * -1, DataType::TYPE_NUMERIC);
						}
						$sheet->getStyleByColumnAndRow($x, $y + 1)->getAlignment()->setHorizontal('center');
						$sheet->getStyleByColumnAndRow($x, $y + 1)->getFont()->setItalic(true)->setBold(false);
						$sheet->mergeCellsByColumnAndRow($x, $y + 1, $x + 1, $y + 1);
					}
				}
				
				$x += 2;
			}
			
			// calcolo della ventilazione totale della riga
			$ventilazioneTotale = [];
			foreach($riga['quantita'] as $quantita) {
				foreach($quantita['ventilazione'] as $key => $value) {
					if (array_key_exists($key, $ventilazioneTotale)) {
						$ventilazioneTotale[$key] += $value;
					} else {
						$ventilazioneTotale[$key] = $value;
					}
				}
			}
			
			$x = 26;
			foreach ($sedi as $codice => $sede) {
				foreach($ventilazioneTotale as $codiceSede => $quantitaVentilata) {
					if ($codiceSede == $sede['codice']) {
						$sheet->setCellValueExplicitByColumnAndRow($x, $y + 1,$quantitaVentilata, DataType::TYPE_NUMERIC);
					}
				}
				$x += 2;
			}
		}
		
		/*$sheet->getStyleByColumnAndRow($xOffset + 1, 11, $xOffset + 18, $yOffset + 1000)
			->getProtection()
			->setLocked(Protection::PROTECTION_UNPROTECTED);
		$sheet->getStyleByColumnAndRow($xOffset + 1, 20, $xOffset + 20, $yOffset + 1000)
			->getProtection()
			->setLocked(Protection::PROTECTION_UNPROTECTED);
		$sheet->getStyleByColumnAndRow($xOffset + 26, 11, $xOffset + 89, $yOffset + 1000)
			->getProtection()
			->setLocked(Protection::PROTECTION_UNPROTECTED);*/ 
		$sheet->getStyleByColumnAndRow($xOffset + 12, $yOffset + 11, $xOffset + 12 + 12, $yOffset + 11 + count($ordine['righe'])*2)->getNumberFormat()->setFormatCode('###,###,##0.00;[Red][<0]-###,###,##0.00; ');
		$sheet->getStyleByColumnAndRow($xOffset + 25, $yOffset + 11, $xOffset + 25 + count($sedi)*2, $yOffset + 11 + count($ordine['righe'])*2)->getNumberFormat()->setFormatCode('###,###,##0;[Red][<0]-###,###,##0; ');
        
		$workBook->setActiveSheetIndex(0);
	}

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
