<?php
    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;

    $timeZone = new DateTimeZone('Europe/Rome');

    //verifico che il file sia stato effettivamente caricato
	if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
	  	echo 'Non hai inviato nessun file...';
	  	//echo json_encode($_FILES, true);
		exit;
	}

    if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
        $inputFileName = "/phpUpload/".$_FILES['userfile']['name'];
    //if(1) {
        //$inputFileName = "/Users/if65/Desktop/test.xlsx";

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
            
            $ordine = [];

            $ordine['fornitore'] = '';
            $ordine['numeroOrdine'] = '';
            $ordine['dataOrdine'] = Date::excelToDateTimeObject(new DateTime(), $timeZone)->format('c');
            $ordine['dataConsegnaPrevista'] = Date::excelToDateTimeObject(new DateTime(), $timeZone)->format('c');
            $ordine['dataConsegnaMinima'] = Date::excelToDateTimeObject(new DateTime(), $timeZone)->format('c');
            $ordine['dataConsegnaMassima'] = Date::excelToDateTimeObject(new DateTime(), $timeZone)->format('c');
            $ordine['category'] = '';
            $ordine['formaPagamento'] = '';
            $ordine['scontoCassaPerc'] = '';
            $ordine['speseTrasportoVal'] = '';
            $ordine['speseTrasportoPerc'] = '';
            $ordine['nomeFoglio'] = ($spreadsheet->getSheetNames())[$sheetNumber];
            
            $colonne = [];
            $colonne['codiceArticoloFornitore'] = -1;
            $colonne['barcode'] = -1;
            $colonne['codiceArticolo'] = -1;
            $colonne['descrizione'] = -1;
            $colonne['marca'] = -1;
            $colonne['modello'] = -1;
            $colonne['famiglia'] = -1;
            $colonne['sottoFamiglia'] = -1;
            $colonne['ivaAliquota'] = -1;
            $colonne['ivaCodice'] = -1;
            $colonne['taglia'] = -1;
            $colonne['costo'] = -1;
            $colonne['scontoA'] = -1;
            $colonne['scontoB'] = -1;
            $colonne['scontoC'] = -1;
            $colonne['scontoD'] = -1;
            $colonne['scontoExtra'] = -1;
            $colonne['scontoImporto'] = -1;
            $colonne['prezzoVendita'] = -1;
            $colonne['EB1'] = -1;
            $colonne['EB2'] = -1;
            $colonne['EB3'] = -1;
            $colonne['EB4'] = -1;
            $colonne['EB5'] = -1;
            $colonne['EBM1'] = -1;
            
            $righe = [];
            foreach ($rows as $rowIndex => $row) {
                $riga = [];

                if ($rowIndex == 0) {
                    // riconoscimento colonne
                    foreach ($row as $columnIndex => $column) {
                        if (preg_match('/^articolo fornitore$/i', $column)): 
                            $colonne['codiceArticoloFornitore'] = $columnIndex;
                        elseif (preg_match('/^descrizione$/i', $column)): 
                            $colonne['descrizione'] = $columnIndex;
                        elseif (preg_match('/^multiplo$/i', $column)): 
                            $colonne['multiplo'] = $columnIndex;
                        elseif (preg_match('/^ean$/i', $column)): 
                            $colonne['barcode'] = $columnIndex;
                        elseif (preg_match('/^prezzo$/i', $column)): 
                            $colonne['prezzoVendita'] = $columnIndex;
                        elseif (preg_match('/^costo$/i', $column)): 
                            $colonne['costo'] = $columnIndex;
                        elseif (preg_match('/^A$/i', $column)): 
                            $colonne['scontoA'] = $columnIndex;
                        elseif (preg_match('/^B$/i', $column)): 
                            $colonne['scontoB'] = $columnIndex;
                        elseif (preg_match('/^C$/i', $column)): 
                            $colonne['scontoC'] = $columnIndex;
                        elseif (preg_match('/^D$/i', $column)): 
                            $colonne['scontoD'] = $columnIndex;
                        elseif (preg_match('/^importo$/i', $column)): 
                            $colonne['scontoImporto'] = $columnIndex;
                        elseif (preg_match('/^extra$/i', $column)): 
                            $colonne['scontoExtra'] = $columnIndex;
                        elseif (preg_match('/^EB1$/i', $column)): 
                            $colonne['EB1'] = $columnIndex;
                        elseif (preg_match('/^EB2$/i', $column)): 
                            $colonne['EB2'] = $columnIndex;
                        elseif (preg_match('/^EB3$/i', $column)): 
                            $colonne['EB3'] = $columnIndex;
                        elseif (preg_match('/^EB4$/i', $column)): 
                            $colonne['EB4'] = $columnIndex;
                        elseif (preg_match('/^EB5$/i', $column)): 
                            $colonne['EB5'] = $columnIndex;
                        elseif (preg_match('/^EBM1$/i', $column)): 
                            $colonne['EBM1'] = $columnIndex;
                        endif;
                    }
                    
                } else {
                    $riga['codiceArticoloFornitore'] = ($colonne['codiceArticoloFornitore'] >= 0 && isset($row[$colonne['codiceArticoloFornitore']])) ?  $row[$colonne['codiceArticoloFornitore']] : '';
                    $riga['barcode'] = ($colonne['barcode'] >= 0 && isset($row[$colonne['barcode']])) ?  $row[$colonne['barcode']] : '';
                    $riga['codiceArticolo'] = '';
                    $riga['descrizione'] = ($colonne['descrizione'] >= 0 && isset($row[$colonne['descrizione']])) ?  $row[$colonne['descrizione']] : '';
                    $riga['marca'] = '';
                    $riga['modello'] = '';
                    $riga['famiglia'] = '';
                    $riga['sottoFamiglia'] = '';
                    $riga['ivaAliquota'] = 22;
                    $riga['ivaCodice'] = 16;
                    $riga['taglia'] = '';
                    $riga['costo'] = ($colonne['costo'] >= 0 && isset($row[$colonne['costo']])) ?  $row[$colonne['costo']]*1 : 0;
                    $riga['scontoA'] = ($colonne['scontoA'] >= 0 && isset($row[$colonne['scontoA']])) ?  $row[$colonne['scontoA']]*1 : 0;
                    $riga['scontoB'] = ($colonne['scontoB'] >= 0 && isset($row[$colonne[scontoAB]])) ?  $row[$colonne['scontoB']]*1 : 0;
                    $riga['scontoC'] = ($colonne['scontoC'] >= 0 && isset($row[$colonne['scontoC']])) ?  $row[$colonne['scontoC']]*1 : 0;
                    $riga['scontoD'] = ($colonne['scontoD'] >= 0 && isset($row[$colonne['scontoD']])) ?  $row[$colonne['scontoD']]*1 : 0;
                    $riga['scontoExtra'] = ($colonne['scontoExtra'] >= 0 && isset($row[$colonne['scontoExtra']])) ?  $row[$colonne['scontoExtra']]*1 : 0;
                    $riga['scontoImporto'] = ($colonne['scontoImporto'] >= 0 && isset($row[$colonne['scontoImporto']])) ?  $row[$colonne['scontoImporto']]*1 : 0;
                    $riga['prezzoVendita'] = ($colonne['prezzoVendita'] >= 0 && isset($row[$colonne['prezzoVendita']])) ?  $row[$colonne['prezzoVendita']]*1 : 0;
                    
                    $riga['quantitaTotale'] = 0;
                    $quantita = [];
                    if($colonne['EB1'] >= 0 && isset($row[$colonne['EB1']])) {
                        $quantita['EB1'] = $row[$colonne['EB1']] * 1;
                        $riga['quantitaTotale'] += $quantita['EB1'];               
                    }
                    if($colonne['EB2'] >= 0 && isset($row[$colonne['EB2']])) {
                        $quantita['EB2'] = $row[$colonne['EB2']] * 1;
                        $riga['quantitaTotale'] += $quantita['EB2'];               
                    }
                    if($colonne['EB3'] >= 0 && isset($row[$colonne['EB3']])) {
                        $quantita['EB3'] = $row[$colonne['EB3']] * 1;
                        $riga['quantitaTotale'] += $quantita['EB3'];               
                    }
                    if($colonne['EB4'] >= 0 && isset($row[$colonne['EB4']])) {
                        $quantita['EB4'] = $row[$colonne['EB4']] * 1;
                        $riga['quantitaTotale'] += $quantita['EB4'];               
                    }
                    if($colonne['EB5'] >= 0 && isset($row[$colonne['EB5']])) {
                        $quantita['EB5'] = $row[$colonne['EB5']] * 1;
                        $riga['quantitaTotale'] += $quantita['EB5'];               
                    }
                    if($colonne['EBM1'] >= 0 && isset($row[$colonne['EBM1']])) {
                        $quantita['EBM1'] = $row[$colonne['EBM1']] * 1;
                        $riga['quantitaTotale'] += $quantita['EBM1'];               
                    }
                    $riga['quantita'] = $quantita;
                    
                    $riga['scontoMerceTotale'] = 0;
                    $scontoMerce = [];
                
                    $righe[] = $riga;
                }
            }

            $ordine['righe'] = $righe;
            $ordine['sedi'] = $sedi;

            $ordini[] = $ordine;
        }

        echo json_encode(array("recordCount" => count($ordini), "values" => $ordini));
    } else {
		echo json_encode($_FILES, true);
	}
