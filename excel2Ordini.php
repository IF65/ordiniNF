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
    
    // verifica formale file
    if ($rows[0][0] == 'FORNITORE' and $rows[6][3] == 'TOTALE ORDINE' and $rows[8][23] == 'COSTO TOTALE') {
        $fornitore = $rows[0][1];
        $numeroOrdine = $rows[1][1];
        $dataOrdine = $rows[2][1];
        $dataConsegnaPrevista = $rows[3][1];
        $dataConsegnaMinima = $rows[4][1];
        $dataConsegnaMassima = $rows[5][1];
        $category = $rows[6][1];
        
        $timeZone = new DateTimeZone('Europe/Rome');
        
        $data = Date::excelToDateTimeObject($dataOrdine, $timeZone);
        $testArray = [];
        $testArray['data'] = $data;
        $testArrayJson = json_encode($testArray, true);
        
        for ($i = 10; $i < count($rows); $i++) {
            echo $rows[$i][0]."\n";
        }
    }
    
    
    
    
    function xls2tstamp($date) {
        return ((($date > 25568) ? $date : 25569) * 86400) - ((70 * 365 + 19) * 86400);
    }
        
        
