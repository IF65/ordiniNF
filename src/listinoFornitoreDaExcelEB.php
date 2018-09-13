<?php
    ini_set('memory_limit', -1);
    
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
		//$inputFileName = "/Users/if65/Desktop/ORECA del 01.01.2013 agg.articoli dal 14.03.2018.xlsx";//<-debug

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
        $sheetNumber = 0;
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
                     
        $listino = [];
        foreach ($rows as $row) {
            $barcode = isset($row[1]) ? trim($row[1]) : '';
            if (preg_match('/^(\d{8}|\d{13}|\d{12})$/', $barcode)) {
                $codiceArticolo = isset($row[0]) ? trim($row[0]) : '';
                $uxi = isset($row[2]) ? $row[2]*1 : 1;
                $descrizione = isset($row[3]) ? trim($row[3]) : '';
                $aliquotaIva = isset($row[4]) ? $row[4]*1 : 0;
                $lordo = isset($row[6]) ? $row[6]*1 : 0;
                $scontoA = isset($row[7]) ? $row[7]*1 : 0;
                $scontoB = isset($row[8]) ? $row[8]*1 : 0;
                $scontoC = isset($row[9]) ? $row[9]*1 : 0;
                $prezzoVendita = isset($row[10]) ? $row[10]*1 : 0;
                
                if ($lordo != 0) {
                    $listino[$barcode] = [
                        'codiceArticolo' => $codiceArticolo,
                        'uxi' => $uxi,
                        'descrizione' => $descrizione,
                        'aliquotaIva' => $aliquotaIva,
                        'lordo' => $lordo,
                        'scontoA' => $scontoA,
                        'scontoB' => $scontoB,
                        'scontoC' => $scontoC,
                        'prezzoVendita' => $prezzoVendita
                    ];
                }
            }
        }            

        echo json_encode($listino);
    } else {
		echo json_encode($_FILES, true);
	}
