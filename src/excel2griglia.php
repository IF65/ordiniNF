<?php

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Shared\Date;
use PhpOffice\PhpSpreadsheet\Reader\Exception as ReaderException;
use PhpOffice\PhpSpreadsheet\Exception as SpreadsheetException;

$timeZone = new DateTimeZone('Europe/Rome');

// verifico che il file sia stato effettivamente caricato

if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
    echo 'Non hai inviato nessun file...';
    //echo json_encode($_FILES, true);
    exit;
}


if (move_uploaded_file($_FILES['userfile']['tmp_name'], "/phpUpload/" . $_FILES['userfile']['name'])) {
    $inputFileName = "/phpUpload/" . $_FILES['userfile']['name'];
    //$inputFileName = "/Users/if65/Desktop/DRY XF48.xlsx";

    /** Create a new Xls Reader  **/
    $reader = new Xlsx();
    //    $reader = new Csv();
    /** Load $inputFileName to a Spreadsheet Object  **/
    $reader->setReadDataOnly(true);
    $reader->setLoadAllSheets();

    $matrix = [];

    try {
        $spreadsheet = $reader->load($inputFileName);
        $worksheet = $spreadsheet->getSheet(0);

        $rows = [];
        foreach ($worksheet->getRowIterator() as $row) {
            $cellIterator = $row->getCellIterator();
            $cellIterator->setIterateOnlyExistingCells(false); // This loops through all cells,
            $cells = [];
            foreach ($cellIterator as $cell) {
                $cells[] = $cell->getValue();
            }
            $rows[] = $cells;
        }

        $pos = [];
        foreach ($rows[0] as $column => $cell) {
            if (preg_match('/^(?:SP|MC)\d+\s(?:MIN|MAX)$/', $cell)) {
                $pos[$column] = $cell;
            }
        }

        for ($r = 1; $r < count($rows); $r++) {
            $id = str_pad($rows[$r][0], 2, '0', STR_PAD_LEFT)
                . '-'
                . str_pad($rows[$r][1], 7, '0', STR_PAD_LEFT);

            $matrix[$id] = [
                'stagionalita' => $rows[$r][0],
                'codice' => $rows[$r][1],
                'codice_fornitore' => $rows[$r][2],
                'descrizione' => $rows[$r][3],
                'taglia' => $rows[$r][4],
                'griglia' => $rows[$r][5],
                'lotto_minimo' => $rows[$r][6],
                'marca' => $rows[$r][7]
            ];
            $sedi = [];
            for ($c = 8; $c < count($rows[$r]); $c++) {
                if (preg_match('/^((?:SP|MC)\d+)\s(MIN|MAX)$/', $pos[$c], $matches)) {
                    $sedi[$matches[1]][$matches[2]] = is_null($rows[$r][$c]) ? 0 : $rows[$r][$c];
                }
            }
            $matrix[$id]['sedi'] = $sedi;
        }
    } catch (ReaderException|SpreadsheetException $e) {
    }

    echo json_encode(array("recordCount" => count($matrix), "values" => $matrix));
} else {
    echo json_encode($_FILES, true);
}
