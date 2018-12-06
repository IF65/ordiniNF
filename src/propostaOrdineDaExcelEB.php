<?php
    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;

    $timeZone = new DateTimeZone('Europe/Rome');

    // verifico che il file sia stato effettivamente caricato
	/*if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
	  	echo 'Non hai inviato nessun file...';
	  	exit;
	}*/

    //if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
        //$inputFileName = "/phpUpload/".$_FILES['userfile']['name'];
        if(1) { //<-debug
		$inputFileName = "/Users/if65/Desktop/Sviluppo/ordiniNF/temp/listino FISCHER.xlsx";//<-debug

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
                         
            $propostaOrdine = [];
            
            foreach ($rows as $row) {
                $barcode = isset($row[0]) ? trim($row[0]) : '';
                $scontoExtra = isset($row[1]) ? $row[1]*1 : 0;
                $sede = isset($row[2]) ? trim($row[2]) : '';
                $quantita = isset($row[3]) ? $row[3]*1 : 0;
                
                if (preg_match('/^(\d{8}|\d{13}|\d{12})$/', $barcode) && preg_match('/^(SM|EB|SP|RU|RS)(\w|\d)/', $sede) && $quantita  != 0) {
                    $propostaOrdine[] = [
                        'barcode' => $barcode,
                        'extra' => $scontoExtra,
                        'sede' => $sede,
                        'quantita' => $quantita
                    ];
                }
            }            
        }

        echo json_encode($propostaOrdine);
    } else {
		echo json_encode($_FILES, true);
	}
