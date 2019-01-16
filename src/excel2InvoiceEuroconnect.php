<?php
    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;

    $debug = false;
    
    $timeZone = new DateTimeZone('Europe/Rome');
    
    $inputFileName = '';

    if ($debug) {
        $inputFileName = "/Users/if65/Desktop/001-21-2019tracciato_excel_importazione_righe_dett_Fattura_versione_1.2_REV-2.xlsx";
    } else {
        if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
            echo 'Non hai inviato nessun file...';
            exit;
        }
        
        if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
            $inputFileName = "/phpUpload/".$_FILES['userfile']['name'];
        }
    }

    if($inputFileName != '') {
		try {
            $reader = new Xlsx();
            $reader->setLoadAllSheets();

            $spreadsheet = IOFactory::load($inputFileName);
            $sheets = $spreadsheet->getSheetNames();
    
            echo json_encode(["recordCount" => count($sheets), "values" => $sheets, "error" => 0]);
        } catch( InvalidArgumentException $e ) {
            echo json_encode(["recordCount" => 0, "values" => [], "errorCode" => 200, "erroMessage" => $e->getMessage()]);
        }
    } else {
		echo json_encode(["recordCount" => 0, "values" => [], "errorCode" => 100, "errorMessage" => 'Nessun file name impostato']);
	}
