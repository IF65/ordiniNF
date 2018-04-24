<?php
    require './vendor/autoload.php';
    
    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;
    
    $inputFileName = './testOrdine.xlsx';
    
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
    
    $spreadsheet = $reader->load($inputFileName);
    
    $worksheet = $spreadsheet->getActiveSheet();
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
    
    $timeZone = new DateTimeZone('Europe/Rome');
    
    // verifica formale file
    if ($rows[0][0] == 'FORNITORE' and $rows[6][3] == 'TOTALE ORDINE' and $rows[8][23] == 'COSTO TOTALE') {
        // determino la posizione di inzio e fine delle sedi
        $inizioSedi = 0;
        $fineSedi = 0;
        for ($i = 0; $i < count($rows[2]); $i++) {
            if ($rows[2][$i] == 'TOTALE PEZZI') {
                $inizioSedi = $i + 1;
            }
            
            if ($rows[2][$i] == 'TOTALE SCONTO MERCE') {
                $fineSedi = $i - 1;
            }
        }
        
        if ($inizioSedi != 0 and $fineSedi != 0) {
            $numeroSedi = $fineSedi -$inizioSedi +1;
            
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
            for ($i = 10; $i < count($rows); $i++) {
                $riga = [];
                
                $riga['codiceArticoloFornitore'] = $rows[$i][0];
                $riga['barcode'] = $rows[$i][1];
                $riga['codiceArticolo'] = $rows[$i][2];
                $riga['descrizione'] = $rows[$i][3];
                $riga['marca'] = $rows[$i][4];
                $riga['modello'] = $rows[$i][5];
                $riga['famiglia'] = $rows[$i][6];
                $riga['sottoFamiglia'] = $rows[$i][7];
                $riga['ivaAliquota'] = $rows[$i][8];
                $riga['ivaCodice'] = $rows[$i][9];
                $riga['taglia'] = $rows[$i][10];
                $riga['costo'] = $rows[$i][11];
                $riga['scontoA'] = $rows[$i][12];
                $riga['scontoB'] = $rows[$i][13];
                $riga['scontoC'] = $rows[$i][14];
                $riga['scontoD'] = $rows[$i][15];
                $riga['scontoExtra'] = $rows[$i][16];
                $riga['scontoImporto'] = $rows[$i][17];
                $riga['prezzoVendita'] = $rows[$i][19];
                
                $quantita = [];
                for ($j = $inizioSedi; $j <= $fineSedi; $j++) {
                    if (preg_match ( '/^(\w\w(?:\w|\d)+)\s\-.*$/', $rows[2][$j], $matches)) {
                        if ($rows[$i][$j] != 0) {
                            $quantita[$matches[1]] = $rows[$i][$j];
                        }
                    }
                }
                $riga['quantita'] = $quantita;
                
                $inizioSedi += $numeroSedi + 1;
                $fineSedi += $numeroSedi + 1;
                $scontoMerce = [];
                for ($j = $inizioSedi; $j <= $fineSedi; $j++) {
                    if (preg_match ( '/^(\w\w(?:\w|\d)+)\s\-.*$/', $rows[2][$j], $matches)) {
                        if ($rows[$i][$j] != 0) {
                            $scontoMerce[$matches[1]] = $rows[$i][$j];
                        }
                    }
                }
                $riga['scontoMerce'] = $scontoMerce;
                
                $righe[] = $riga;
            }
            
            $ordine['righe'] = $righe;
            
            $text = json_encode($ordine, true);
            print_r($text);
        }
        
    }
    
        
        
