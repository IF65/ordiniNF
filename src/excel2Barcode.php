<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Shared\Date;

$timeZone = new DateTimeZone('Europe/Rome');

$debug = false;

if ($debug) {
    $inputFileName = "/Users/if65/Desktop/Porchi/ean_carhartt_41.xlsx";
} else {
    if (!isset( $_FILES['userfile'] ) || !is_uploaded_file( $_FILES['userfile']['tmp_name'] )) {
        echo 'Non hai inviato nessun file...';
        exit;
    }

    if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/" . $_FILES['userfile']['name'] )) {
        $inputFileName = "/phpUpload/" . $_FILES['userfile']['name'];
    }
}



/** Create a new Xls Reader  **/
$reader = new Xlsx();
$reader->setReadDataOnly(true);
$reader->setLoadAllSheets();

$dati = [];

$spreadsheet = $reader->load($inputFileName);

$maxRowCount = 0;
foreach ($spreadsheet->getSheetNames() as $sheetName) {
    $worksheet = $spreadsheet->getSheetByName($sheetName);
    $count = $worksheet->getHighestRow();
    if ($maxRowCount < $count) {
        $maxRowCount = $count;
    }
}

if ($maxRowCount < 10000) {
    foreach ($spreadsheet->getSheetNames() as $sheetName) {
        $worksheet = $spreadsheet->getSheetByName( $sheetName );
        $rows = [];
        foreach ($worksheet->getRowIterator() AS $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells( FALSE ); // This loops through all cells,
            $cells = [];
            foreach ($cellIterator as $cell) {
                $cells[] = $cell->getValue();
            }
            $rows[] = $cells;
        }

        $currentSheetRows = [];
        foreach ($rows as $row) {
            $currentSheetRow['codiceArticoloFornitore'] = "$row[0]";
            $currentSheetRow['taglia'] = "$row[1]";
            $currentSheetRow['barcode'] = "$row[2]";
            $currentSheetRows[] = $currentSheetRow;
        }
        $dati[$sheetName] = $currentSheetRows;
    }

    echo json_encode($dati);
} else {
    http_response_code(408);//troppe righe
}

