<?php
    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;

    $timeZone = new DateTimeZone('Europe/Rome');

    // verifico che il file sia stato effettivamente caricato
	if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
	  	echo 'Non hai inviato nessun file...';
	  	exit;
	}

    if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
        $inputFileName = "/phpUpload/".$_FILES['userfile']['name'];
    //if(1) { //<-debug
		//$inputFileName = "/Users/if65/Desktop/ordini_560252582.xlsx";//<-debug

        /** Create a new Xls Reader  **/
        $reader = new Xlsx();
        //    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
        //    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xml();
        //    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Ods();
        //    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Slk();
        //    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Gnumeric();
        //    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Csv();
        /** Load $inputFileName to a Spreadsheet Object  **/
        $reader->setReadDataOnly(true);
        $reader->setLoadAllSheets();

        $ordini = [];

        $spreadsheet = $reader->load($inputFileName);
        for ($sheetNumber = 0; $sheetNumber < $spreadsheet->getSheetCount(); $sheetNumber++) {
            $worksheet = $spreadsheet->getSheet($sheetNumber);
            $rows = [];
            foreach ($worksheet->getRowIterator() AS $row) {
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(FALSE); // This loops through all cells,
                $cells = [];
                foreach ($cellIterator as $cell) {
                    $cells[] = $cell->getValue();
                }
                $rows[] = $cells;
            }

            // verifica formale file
            if ($rows[0][0] == 'FORNITORE' and $rows[6][3] == 'TOTALE ORDINE' and $rows[8][23] == 'COSTO TOTALE') {
                // determino la posizione di inzio e fine delle sedi
                $colonnaInizioSedi = 25;
                $colonnaFineSedi = 0;
                for ($index = count($rows[0]); $index >= 0; $index--) {
                    if (preg_match('/^(\w\w(?:\w|\d)+)\s\-.*$/', $rows[0][$index])) {
                         $colonnaFineSedi = $index;
                         break;
                    }
                }
                $numeroSedi = ($colonnaFineSedi - $colonnaInizioSedi) > 0 ? ($colonnaFineSedi - $colonnaInizioSedi)/2 + 1 : 0;
                
                $rigaInizioArticoli = 10;
                $rigaFineArticoli = 0;
                for ($index = count($rows); $index >= 0; $index--) {
                    if (preg_match('/^\=SUM\(/', $rows[$index][24])) {
                         $rigaFineArticoli = $index;
                         break;
                    }
                }
                $numeroArticoli = ($rigaFineArticoli - $rigaInizioArticoli + 1) > 0 ? ($rigaFineArticoli - $rigaInizioArticoli + 1)/2 : 0;

                
                if ($numeroSedi && $numeroArticoli) {
                    $ordine = [];

                    $ordine['fornitore'] = $rows[0][1];
                    $ordine['numeroOrdine'] = $rows[1][1];
                    $ordine['dataOrdine'] = Date::excelToDateTimeObject($rows[2][1], $timeZone)->format('c');
                    $ordine['dataConsegnaPrevista'] = Date::excelToDateTimeObject($rows[3][1], $timeZone)->format('c');
                    $ordine['dataConsegnaMinima'] = Date::excelToDateTimeObject($rows[4][1], $timeZone)->format('c');
                    $ordine['dataConsegnaMassima'] = Date::excelToDateTimeObject($rows[5][1], $timeZone)->format('c');
                    $ordine['category'] = $rows[6][1];
                    $ordine['formaPagamento'] = $rows[0][4];
                    $ordine['scontoCassaPerc'] = $rows[1][4];
                    $ordine['speseTrasportoVal'] = $rows[2][4];
                    $ordine['speseTrasportoPerc'] = $rows[3][4];
					
                    $righe = [];
                    for ($indexArticolo = 0; $indexArticolo < $numeroArticoli; $indexArticolo++) {
                        
                        $indexRow = $rigaInizioArticoli + $indexArticolo*2;
                        
                        $riga = [];
                        $riga['codiceArticoloFornitore'] = $rows[$indexRow][0];
                        $riga['barcode'] = $rows[$indexRow][1];
                        $riga['codiceArticolo'] = $rows[$indexRow][2];
                        $riga['descrizione'] = $rows[$indexRow][3];
                        $riga['marca'] = $rows[$indexRow][4];
                        $riga['modello'] = $rows[$indexRow][5];
                        $riga['famiglia'] = $rows[$indexRow][6];
                        $riga['sottoFamiglia'] = $rows[$indexRow][7];
                        $riga['ivaAliquota'] = $rows[$indexRow][8];
                        $riga['ivaCodice'] = $rows[$indexRow][9];
                        $riga['taglia'] = $rows[$indexRow][10];
                        $riga['costo'] = $rows[$indexRow][11];
                        $riga['scontoA'] = $rows[$indexRow][12];
                        $riga['scontoB'] = $rows[$indexRow][13];
                        $riga['scontoC'] = $rows[$indexRow][14];
                        $riga['scontoD'] = $rows[$indexRow][15];
                        $riga['scontoExtra'] = $rows[$indexRow][16];
                        $riga['scontoImporto'] = $rows[$indexRow][17];
                        $riga['prezzoVendita'] = $rows[$indexRow][19];
						
						$riga['quantitaTotale'] = 0;
                        $quantita = [];
                        for ($indexSede = 0; $indexSede < $numeroSedi; $indexSede++) {
                            $indexColumn = $colonnaInizioSedi + $indexSede*2;
                            
                            if (preg_match ( '/^(\w\w(?:\w|\d)+)\s\-.*$/', $rows[0][$indexColumn], $matches)) {
                            	 if ($rows[$indexRow][$indexColumn] != 0) {
                                    $quantita[$matches[1]] = $rows[$indexRow][$indexColumn]*1;
                                    $riga['quantitaTotale'] += $rows[$indexRow][$indexColumn]*1;
                                }
                            }
                        }
                        $ventilazione = [];
                        for ($indexSede = 0; $indexSede < $numeroSedi; $indexSede++) {
                            $indexColumn = $colonnaInizioSedi + $indexSede*2;
                            
                            if (preg_match ( '/^(\w\w(?:\w|\d)+)\s\-.*$/', $rows[0][$indexColumn], $matches)) {
                            	 if ($rows[$indexRow + 1][$indexColumn] > 0) {
                                    $ventilazione[$matches[1]] = $rows[$indexRow + 1][$indexColumn]*1;
                                }
                            }
                        }
                        if (! empty($ventilazione)) {
                            $quantita['ventilazione'] = $ventilazione;
                        }
                        $riga['quantita'] = $quantita;
						
						$riga['scontoMerceTotale'] = 0;
                        $scontoMerce = [];
                        for ($indexSede = 0; $indexSede < $numeroSedi; $indexSede++) {
                            $indexColumn = $colonnaInizioSedi + $indexSede*2 + 1;
                            
                            if (preg_match ( '/^(\w\w(?:\w|\d)+)\s\-.*$/', $rows[0][$indexColumn], $matches)) {
                                if ($rows[$indexRow][$indexColumn] != 0) {
                                    $scontoMerce[$matches[1]] = $rows[$indexRow][$indexColumn]*1;
                                    $riga['scontoMerceTotale'] = $rows[$indexRow][$indexColumn]*1;
                                }
                            }
                        }
                        $riga['scontoMerce'] = $scontoMerce;

                        $righe[] = $riga;
                    }

                    $ordine['righe'] = $righe;
                    $ordine['sedi'] = $sedi;

                    $ordini[] = $ordine;
                }

            }
        }
        
        // trasformo l'array associativo in formato test
        $txtResponse = "";
        foreach ($ordini as $ordine) {
            $txtResponse .= "INIZIO ORDINE\n";

            
            $txtResponse .= $ordine['fornitore'] . "\t";
            $txtResponse .= $ordine['numeroOrdine'] . "\t";
            $txtResponse .= (new DateTime($ordine['dataOrdine']))->format('d/m/Y') . "\t";
            $txtResponse .= (new DateTime($ordine['dataConsegnaPrevista']))->format('d/m/Y') . "\t";
            $txtResponse .= (new DateTime($ordine['dataConsegnaMinima']))->format('d/m/Y') . "\t";
            $txtResponse .= (new DateTime($ordine['dataConsegnaMassima']))->format('d/m/Y') . "\t";
            $txtResponse .= $ordine['category'] . "\t";
            $txtResponse .= $ordine['formaPagamento'] . "\t";
            $txtResponse .= number_format($ordine['scontoCassaPerc']*1,2,',','') . "\t";
            $txtResponse .= number_format($ordine['speseTrasportoVal']*1,2,',','') . "\t";
            $txtResponse .= number_format($ordine['speseTrasportoPerc']*1,2,',','') . "\n";
            
            $txtResponse .= "INIZIO RIGHE\n";
            foreach($ordine['righe'] as $riga) {
                $txtResponse .= $riga['codiceArticoloFornitore'] . "\t";
                $txtResponse .= $riga['barcode'] . "\t";
                $txtResponse .= $riga['codiceArticolo'] . "\t";
                $txtResponse .= $riga['descrizione'] . "\t";
                $txtResponse .= $riga['marca'] . "\t";
                $txtResponse .= $riga['modello'] . "\t";
                $txtResponse .= $riga['famiglia'] . "\t";
                $txtResponse .= $riga['sottoFamiglia'] . "\t";
                $txtResponse .= number_format($riga['ivaAliquota']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['ivaCodice']*1,2,',','') . "\t";
                $txtResponse .= $riga['taglia'] . "\t";
                $txtResponse .= number_format($riga['costo']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['scontoA']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['scontoB']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['scontoC']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['scontoD']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['scontoExtra']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['scontoImporto']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['prezzoVendita']*1,2,',','') . "\t";
                $txtResponse .= number_format($riga['quantitaTotale']*1,2,',','') . "\n";
                $txtResponse .= "INIZIO QUANTITA\n";
                foreach($riga['quantita'] as $key => $value) {
                    if ($key != 'ventilazione') {
                        if ($value > 0) {
                            $txtResponse .= $key . "\t";
                            $txtResponse .= number_format($value*1,2,',','') . "\n";
                        }
                    } else {
                        $txtResponse .= "INIZIO VENTILAZIONE\n";
                        foreach($value as $sede => $quantita) {
                            if ($quantita > 0) {
                                $txtResponse .= $sede . "\t";
                                $txtResponse .= number_format($quantita*1,2,',','') . "\n";
                            }
                        }
                        $txtResponse .= "FINE VENTILAZIONE\n";
                    }
                    
                }
                $txtResponse .= "FINE QUANTITA\n";
                
        }
            $txtResponse .= "FINE RIGHE\n";
            
            $txtResponse .= "FINE ORDINE\n";
        }
        

        echo $txtResponse;
    } else {
		echo json_encode($_FILES, true);
	}
