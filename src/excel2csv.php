<?php
    @ini_set('memory_limit','65536M');
    
    require '/Users/if65/Desktop/Sviluppo/ordiniNF/vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Writer\Csv;

    if (count($argv) > 1) {
        $inputFileName = $argv[1];
        if (file_exists($inputFileName)) {
            $outputFileName = preg_replace('/((?:\.xlsx|\.xls))$/','',$inputFileName);
    
            if ($inputFileName != '') {
                $reader = new Xlsx();
                $reader->setReadDataOnly(true);
    
                $spreadsheet = $reader->load($inputFileName);
                $sheetCount = $spreadsheet->getSheetCount();
                for ($i = 0; $i < $sheetCount; $i++) {
                    $sheet = $spreadsheet->getSheet($i);
    
                    $writer = new Csv($spreadsheet);
                    $writer->setDelimiter(';');
                    $writer->setEnclosure('');
                    $writer->setLineEnding("\r\n");
                    $writer->setSheetIndex($i);
                    
                    $writer->save($outputFileName."_$i.csv");
                    
                    unset($writer);
                }
                
                unset($reader);
            }
        } else {
            echo "il file $inputFileName non esiste!\n";
        }
    }
?>