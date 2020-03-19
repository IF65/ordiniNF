<?php
require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Shared\Date;

$timeZone = new DateTimeZone('Europe/Rome');

// verifico che il file sia stato effettivamente caricato
if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
    echo 'Non hai inviato nessun file...';
    //echo json_encode($_FILES, true);
    exit;
}

if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
    $inputFileName = "/phpUpload/".$_FILES['userfile']['name'];

    //$inputFileName = "/Users/if65/Desktop/Catalina/test.xlsx";

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

    //carico il file
    $spreadsheet = $reader->load($inputFileName);

    //utilizzo per il caricamento dei dati solo il foglio 0
    $worksheet = $spreadsheet->getSheet(0);
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
    if (count($rows[0]) >= 7) {
        for ($i = 1; $i < count($rows); $i++) {
            $promozione = [];
            $promozione['mclu'] = $rows[$i][0];
            $promozione['fascia'] = $rows[$i][1];
            $promozione['valoreFacciale'] = $rows[$i][2];
            $promozione['tipologia'] = $rows[$i][3];
            $promozione['minimoSpesa'] = $rows[$i][4];
            $promozione['inizioValidita'] = Date::excelToDateTimeObject($rows[$i][5], Date::getDefaultTimezone())->format('Y-m-d');
            $promozione['fineValidita'] = Date::excelToDateTimeObject($rows[$i][6], Date::getDefaultTimezone())->format('Y-m-d');
            $promozione['barcode'] = (count($rows[$i]) > 7 && $rows[$i][7] <> null) ? $barcode = $rows[$i][7] : 0;

            $promozioni[] = $promozione;
        }
    }

    echo json_encode(array("recordCount" => count($promozioni), "values" => $promozioni));
} else {
    echo json_encode($_FILES, true);
}
