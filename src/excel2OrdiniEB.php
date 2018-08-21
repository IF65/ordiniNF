<?php
    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;

    $timeZone = new DateTimeZone('Europe/Rome');

    // verifico che il file sia stato effettivamente caricato
	/*if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
	  	echo 'Non hai inviato nessun file...';
	  	exit;
	}

    if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
        $inputFileName = "/phpUpload/".$_FILES['userfile']['name'];*/
    if(1) { //<-debug
		$inputFileName = "/Users/if65/Desktop/Sviluppo/ordiniNF/temp/ERREBI Ecobrico - Errebi 2018.06.08 proposta d'assortimento 4x100.xlsx";//<-debug

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
            // leggo le intestazioni delle colonne
            for ($col = 1; $col <= $highestColumnIndex; ++$col) {
                $value = $worksheet->getCellByColumnAndRow($col, 3)->getValue();
                if (preg_match ( '/descrizione/i', $value, $matches)) {
                    $indiceColonne['descrizione'] = $i;
                }
                if (preg_match ( '/EAN/i', $value, $matches)) {
                    $indiceColonne['ean'] = $col;
                }
                if (preg_match ( '/codice/i', $value, $matches)) {
                    $indiceColonne['codice'] = $col;
                }
                if (preg_match ( '/descrizione/i', $value, $matches)) {
                    $indiceColonne['descrizione'] = $col;
                }
                if (preg_match ( '/prezzo/i', $value, $matches)) {
                    $indiceColonne['prezzo'] = $col;
                }
                if (preg_match ( '/sconto extra/i', $value, $matches)) {
                    $indiceColonne['sconto extra'] = $col;
                }
                if (preg_match ( '/vendita/i', $value, $matches)) {
                    $indiceColonne['vendita'] = $col;
                }
                if (preg_match ( '/famiglia/i', $value, $matches)) {
                    $indiceColonne['famiglia'] = $col;
                }
                if (preg_match ( '/quantita/i', $value, $matches)) {
                    $indiceColonne['quantita'] = $col;
                }
                if (preg_match ( '/conf/i', $value, $matches)) {
                    $indiceColonne['confezione'] = $col;
                }
            }
            
            $dataCorrente = new DateTime(null, $timeZone);
            
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
            if ($rows[0][0] == 'FORNITORE') {
                         
                $ordine = [];
                
                $ordine['fornitore'] = 'FERREBI';
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
                for ($indexArticolo = 3; $indexArticolo < $highestRow; $indexArticolo++) {
                    
                    $riga = [];
                    $riga['codiceArticoloFornitore'] = $rows[$indexRow][$indiceColonne['codice']];
                    $riga['barcode'] = $rows[$indexRow][$indiceColonne['ean']];
                    $riga['codiceArticolo'] = '';
                    $riga['descrizione'] = $rows[$indexRow][$indiceColonne['descrizione']];
                    $riga['marca'] = '';
                    $riga['modello'] = '';
                    $riga['famiglia'] = $rows[$indexRow][$indiceColonne['famiglia']];
                    $riga['sottoFamiglia'] ='';
                    $riga['ivaAliquota'] = 22;
                    $riga['ivaCodice'] = 16;
                    $riga['taglia'] = '';
                    $riga['costo'] = $rows[$indexRow][$indiceColonne['prezzo']];
                    $riga['scontoA'] = 0;
                    $riga['scontoB'] = 0;
                    $riga['scontoC'] = 0;
                    $riga['scontoD'] = 0;
                    $riga['scontoExtra'] = $indiceColonne['sconto extra'];
                    $riga['scontoImporto'] = 0;
                    $riga['prezzoVendita'] = $rows[$indexRow][$indiceColonne['vendita']];
                    $riga['quantita'] = $rows[$indexRow][$indiceColonne['quantita']];
                    
                    $righe[] = $riga;
                }

                $ordine['righe'] = $righe;
                $ordine['sedi'] = $sedi;

                $ordini[] = $ordine;
                
            }
        }

        echo json_encode(array("recordCount" => count($ordini), "values" => $ordini));
    } else {
		echo json_encode($_FILES, true);
	}
