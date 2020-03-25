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
		$inputFileName = "/Users/if65/Desktop/Template Caricamento Promozioni.xlsx";//<-debug

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

        $template = [];

        $spreadsheet = $reader->load($inputFileName);
        foreach ($spreadsheet->getSheetNames() as $sheetName) {
            $worksheet = $spreadsheet->getSheetByName($sheetName);
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

            $promozioni = [];
            for ($i = 1;$i < count($rows); $i++) {
                if ($sheetName == '0054') {
                    $promozioni[] =
                        [
                            "idIf65" => $rows[$i][0] == null ? 0 : $rows[$i][0],
                            "numero" => $rows[$i][1] == null ? 0 : $rows[$i][1],
                            "promovar" => $rows[$i][2] == null ? 0 : $rows[$i][2],
                            "denominazione" => $rows[$i][3] == null ? "" : $rows[$i][3],
                            "percentuale" => $rows[$i][4] == null ? 0 : $rows[$i][4],
                            "dataInizio" => Date::excelToDateTimeObject( $rows[$i][5], $timeZone )->format( 'c' ),
                            "dataFine" => Date::excelToDateTimeObject( $rows[$i][6], $timeZone )->format( 'c' ),
                            "barcodeGruppo1" => $rows[$i][7] == null ? "" : $rows[$i][7],
                            "barcodeGruppo2" => $rows[$i][8] == null ? "" : $rows[$i][8],
                            "codiciArticoloGruppo1" => $rows[$i][9] == null ? "" : $rows[$i][9],
                            "codiciArticoloGruppo2" => $rows[$i][10] == null ? "" : $rows[$i][10],
                            "aderenti" => $rows[$i][11] == null ? "" : $rows[$i][11],
                            "aderentiGruppo" => $rows[$i][12] == null ? "" : $rows[$i][12]
                        ];
                }
            }

            if (count($promozioni)){
                $template[$sheetName] = $promozioni;
            }
        }

        echo json_encode($template);
    } else {
		echo json_encode($_FILES, true);
	}
