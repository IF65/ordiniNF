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
		//$inputFileName = "/Users/if65/Desktop/Sviluppo/ordiniNF/temp/FAB.xlsx";//<-debug

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
            
            $highestRow = $worksheet->getHighestRow(); // e.g. 10
            $highestColumn = $worksheet->getHighestColumn(); // e.g 'F'     
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
            
            $indiceColonne = [];
            
            $dataCorrente = new DateTime(null, $timeZone);
            
            $rows = [];
            foreach ($worksheet->getRowIterator() as $row) {
                $cellIterator = $row->getCellIterator();
                $cellIterator->setIterateOnlyExistingCells(False); // This loops through all cells,
                $cells = [];
                foreach ($cellIterator as $cell) {
                    $cells[] = $cell->getValue();
                }
                $rows[] = $cells;
            }
                         
            $ordine = [];
            
            $ordine['fornitore'] = '';
            $ordine['numeroOrdine'] = '';
            $ordine['dataOrdine'] = Date::excelToDateTimeObject($dataCorrente, $timeZone)->format('c');
            $ordine['dataConsegnaPrevista'] = Date::excelToDateTimeObject($dataCorrente, $timeZone)->format('c');
            $ordine['dataConsegnaMinima'] = Date::excelToDateTimeObject($dataCorrente, $timeZone)->format('c');
            $ordine['dataConsegnaMassima'] = Date::excelToDateTimeObject($dataCorrente, $timeZone)->format('c');
            $ordine['category'] = '';
            $ordine['formaPagamento'] = '';
            $ordine['scontoCassaPerc'] = 0;
            $ordine['speseTrasportoVal'] = 0;
            $ordine['speseTrasportoPerc'] = 0;
            
            $righe = [];
            foreach ($rows as $row) {
               
                $barcode = isset($row[1]) ? trim($row[1]) : '';
                if (preg_match('/^(\d{8}|\d{13}|\d{12})$/', $barcode)) {
                    
                    $fornitore = strtoupper(trim($row[7]));
                    $famiglia = isset($row[3]) ? trim($row[3]) : '';
                    $sottofamiglia = isset($row[4]) ? trim($row[4]) : '999';
                    $descrizione = strtoupper(trim($row[5]));
                    $marca = strtoupper(trim($row[6]));
                    $costo = isset($row[11]) ? $row[11]*1 : 0;
                    $prezzoVendita = isset($row[17]) ? $row[17]*1 : 0;
                    $scontoA = isset($row[12]) ? $row[12]*1 : 0;
                    $scontoB = isset($row[13]) ? $row[13]*1 : 0;
                    $scontoC = isset($row[14]) ? $row[14]*1 : 0;
                    $scontoExtra = isset($row[15]) ? $row[15]*1 : 0;
                    
                    $quantitaEB1 = isset($row[18]) ? $row[18]*1 : 0;
                    $quantitaEB3 = isset($row[19]) ? $row[19]*1 : 0;
                    $quantitaEB4 = isset($row[20]) ? $row[20]*1 : 0;
                    $quantitaEB5 = isset($row[21]) ? $row[21]*1 : 0;
                    $quantitaEBM1 = isset($row[22]) ? $row[22]*1 : 0;
                    
                    if (preg_match('/^F\w+/', $fornitore) && preg_match('/^\d{9}/', $famiglia) && preg_match('/^\d{3}/', $sottofamiglia) &&
                        $descrizione != '' && $marca != '' && $costo != 0 && $prezzoVendita != 0) {
                        
                        $ordine['fornitore'] = $fornitore;
                        
                        $riga = [];
                        $riga['codiceArticolo'] = isset($row[0]) ? trim($row[0]) : '';
                        $riga['barcode'] = $barcode;
                        $riga['codiceArticoloFornitore'] = isset($row[2]) ? trim($row[2]) : '';
                        $riga['famiglia'] = $famiglia;
                        $riga['sottoFamiglia'] = $sottofamiglia;
                        $riga['descrizione'] = $descrizione;
                        $riga['marca'] = $marca;
                        $riga['fornitore'] = $fornitore;
                        $riga['uxi'] = isset($row[8]) ? $row[8]*1 : 1;
                        $riga['ivaCodice'] = isset($row[9]) ? $row[9]*1 : 16;
                        $riga['ivaAliquota'] = isset($row[10]) ? $row[10]*1 : 22;
                        $riga['costo'] = $costo;
                        $riga['scontoA'] = $scontoA;
                        $riga['scontoB'] = $scontoB;
                        $riga['scontoC'] = $scontoC;
                        $riga['scontoExtra'] = $scontoExtra;
                        $riga['costoFinito'] = round($costo * (100 - $scontoA/100)/100 * (100 - $scontoB/100)/100 * (100 - $scontoC/100)/100 * (100 - $scontoExtra/100)/100,2);
                        $riga['prezzoVendita'] = $prezzoVendita;
                        $riga['quantitaTotale'] = $quantitaEB1 + $quantitaEB3 + $quantitaEB4 + $quantitaEB5 + $quantitaEBM1;
                        $riga['scontoMerceTotale'] = 0;
                        $quantita = [];
                        if ($quantitaEB1 != 0) {
                            $quantita[] = ['descrizione' => 'EB1 - Sonico', 'quantita' => $quantitaEB1, 'scontoMerce' => 0, 'sede' => 'EB1'];
                        }
                        if ($quantitaEB3 != 0) {
                            $quantita[] = ['descrizione' => 'EB3 - Roe\' Volciano', 'quantita' => $quantitaEB3, 'scontoMerce' => 0, 'sede' => 'EB3'];
                        }
                        if ($quantitaEB4 != 0) {
                            $quantita[] = ['descrizione' => 'EB4 - Castel Goffredo', 'quantita' => $quantitaEB4, 'scontoMerce' => 0, 'sede' => 'EB4'];
                        }
                        if ($quantitaEB5 != 0) {
                            $quantita[] = ['descrizione' => 'EB5 - Pralboino', 'quantita' => $quantitaEB5, 'scontoMerce' => 0, 'sede' => 'EB5'];
                        }
                        if ($quantitaEBM1 != 0) {
                            $quantita[] = ['descrizione' => 'EBM1 - Magazzino', 'quantita' => $quantitaEBM1, 'scontoMerce' => 0, 'sede' => 'EBM1'];
                        }
                        $riga['quantita'] = $quantita;
                        
                        $righe[] = $riga;
                    }
                }
            }

            $ordine['righe'] = $righe;

            $ordini[] = $ordine;
            
        }

        echo json_encode(array("recordCount" => count($ordini), "values" => $ordini));
    } else {
		echo json_encode($_FILES, true);
	}
