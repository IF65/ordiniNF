<?php
    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;

    $debug = true;
    
    $timeZone = new DateTimeZone('Europe/Rome');
    
    $inputFileName = '';

    if ($debug) {
        $inputFileName = "/Users/if65/Desktop/001-21-2019tracciato_excel_importazione_righe_dett_Fattura_versione_1.2_REV-2.xlsx";
        $sheetName = '2.2   <DatiBeniServizi>';
    } else {
        if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
            echo 'Non hai inviato nessun file...';
            exit;
        }
        
        if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
            $inputFileName = "/phpUpload/".$_FILES['userfile']['name'];
            
            $sheetName = '';
            foreach (getallheaders() as $name => $value) {
                if ($name == 'Sheet-Name') {
                    $sheetName = $value;
                }
            }
        }
    }

    if ($inputFileName != '' &&  $sheetName != '') {
		try {
            $reader = new Xlsx();
            $reader->setLoadSheetsOnly($sheetName);
            $reader->setReadDataOnly(true);

            $spreadsheet = IOFactory::load($inputFileName);
            $worksheet = $spreadsheet->getSheetByName($sheetName);
            
            $highestRow = $worksheet->getHighestRow(); // e.g. 10
            $highestColumn = $worksheet->getHighestColumn(); // e.g 'F'     
            $highestColumnIndex = \PhpOffice\PhpSpreadsheet\Cell\Coordinate::columnIndexFromString($highestColumn);
            
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
            
            $fattura = [];
            // lettura testata
            $testata = [];
            
            $testata['tipoCogeMitico'] = isset($rows[0][3]) ? $rows[0][3] : '';
            $testata['contoRicavoMitico'] = isset($rows[1][3]) ? $rows[1][3] : '';
            $testata['codiceRapportoCliente'] = isset($rows[2][3]) ? $rows[2][3] : '';
            $testata['codiceCliente'] = isset($rows[3][3]) ? $rows[3][3] : '';
            
            $testata['codiceDestinatario'] = isset($rows[4][3]) ? $rows[4][3] : '';
            $testata['pecDestinatario'] = isset($rows[5][3]) ? $rows[5][3] : '';
            
            $testata['idPaese'] = isset($rows[9][3]) ? $rows[9][3] : '';
            $testata['idCodice'] = isset($rows[10][3]) ? $rows[10][3] : '';
            $testata['codiceFiscale'] = isset($rows[11][3]) ? $rows[11][3] : '';
            
            $testata['denominazione'] = isset($rows[13][3]) ? $rows[13][3] : '';
            $testata['nome'] = isset($rows[14][3]) ? $rows[14][3] : '';
            $testata['cognome'] = isset($rows[15][3]) ? $rows[15][3] : '';
            $testata['codiceEori'] = isset($rows[16][3]) ? $rows[16][3] : '';
            
            $testata['indirizzo'] = isset($rows[18][3]) ? $rows[18][3] : '';
            $testata['numeroCivico'] = isset($rows[19][3]) ? $rows[19][3] : '';
            $testata['cap'] = isset($rows[20][3]) ? $rows[20][3] : '';
            $testata['comune'] = isset($rows[21][3]) ? $rows[21][3] : '';
            $testata['provincia'] = isset($rows[22][3]) ? $rows[22][3] : '';
            $testata['nazione'] = isset($rows[23][3]) ? $rows[23][3] : '';
            
            $testata['tipoDocumento'] = isset($rows[26][3]) ? $rows[26][3] : '';
            $testata['divisa'] = isset($rows[27][3]) ? $rows[27][3] : '';
            $testata['data'] = Date::excelToDateTimeObject($rows[28][3])->format('c');
            $testata['numero'] = isset($rows[29][3]) ? $rows[29][3] : '';
            
            $testata['bolloVirtuale'] = isset($rows[31][3]) ? $rows[31][3] : '';
            $testata['bolloImporto'] = isset($rows[32][3]) ? $rows[32][3] : '0';
            
            $testata['scontoMaggiorazioneTipo'] = isset($rows[34][3]) ? $rows[34][3] : '';
            $testata['scontoMaggiorazionePercentuale'] = isset($rows[35][3]) ? $rows[35][3] : '';
            $testata['scontoMaggiorazioneImporto'] = isset($rows[36][3]) ? $rows[36][3] : '';
            
            $testata['causale'] = isset($rows[37][3]) ? $rows[37][3] : '';
            
            $testata['codiceCommessaConvenzione'] = isset($rows[39][3]) ? $rows[39][3] : '';
            $testata['codiceCup'] = isset($rows[40][3]) ? $rows[40][3] : '';
            $testata['codiceCig'] = isset($rows[41][3]) ? $rows[41][3] : '';
            
            $testata['condizioniPagamento'] = isset($rows[43][3]) ? $rows[43][3] : '';
            
            $testata['modalitaPagamento'] = isset($rows[45][3]) ? $rows[45][3] : '';
            $testata['dataPagamento'] = isset($rows[28][3]) ? Date::excelToDateTimeObject($rows[28][3])->format('c') : '';
            $testata['giorniPagamento'] = isset($rows[47][3]) ? $rows[47][3] : '';
            $testata['dataScadenzaPagamento'] = isset($rows[48][3]) ? Date::excelToDateTimeObject($rows[48][3])->format('c') : '';
            $testata['iban'] = isset($rows[49][3]) ? $rows[49][3] : '';
            
            
            $righe = [];
            foreach ($rows as $num => $row) {
                if ($num > 53) {
                    $riga = [];
                    
                    $riga['idDocumento'] = isset($row[0]) ? $row[0] : '';
                    $riga['data'] = isset($row[1]) ? Date::excelToDateTimeObject($row[1])->format('c') : '';
                    $riga['numItem'] = isset($row[2]) ? $row[2] : '';
                    $riga['numeroDdt'] = isset($row[3]) ? $row[3] : '';
                    $riga['dataDdt'] = isset($row[4]) ? Date::excelToDateTimeObject($row[4])->format('c') : '';
                    $riga['tipoCessione'] = isset($row[5]) ? $row[5] : '';
                    
                    $riga['codice1Tipo'] = isset($row[6]) ? $row[6] : '';
                    $riga['codice1Valore'] = isset($row[7]) ? $row[7] : '';
                    $riga['codice2Tipo'] = isset($row[8]) ? $row[8] : '';
                    $riga['codice2Valore'] = isset($row[9]) ? $row[9] : '';
                    $riga['codice3Tipo'] = isset($row[10]) ? $row[10] : '';
                    $riga['codice3Valore'] = isset($row[11]) ? $row[11] : '';
                    
                    $riga['descrizione'] = isset($row[12]) ? $row[12] : '';
                    $riga['quantita'] = isset($row[13]) ? $row[13] * 1: 0;
                    $riga['unitaMisura'] = isset($row[14]) ? $row[14] : '';
                    $riga['prezzoUnitario'] = isset($row[15]) ? $row[15] * 1: 0;
                    
                    $riga['sconto1Tipo'] = isset($row[16]) ? $row[16] : '';
                    $riga['sconto1Percentuale'] = isset($row[17]) ? $row[17] * 1: 0;
                    $riga['sconto1Importo'] = isset($row[18]) ? $row[18] * 1: 0;
                    $riga['sconto2Tipo'] = isset($row[19]) ? $row[19] : '';
                    $riga['sconto2Percentuale'] = isset($row[20]) ? $row[20] * 1: 0;
                    $riga['sconto2Importo'] = isset($row[21]) ? $row[21] * 1: 0;
                    
                    $riga['prezzoTotale'] = $riga['prezzoUnitario'] * $riga['quantita'];//isset($row[22]) ? $row[22] * 1: 0;
                    $riga['aliquotaIva'] = isset($row[23]) ? $row[23] * 1: 0;
                    $riga['natura'] = isset($row[24]) ? $row[24] : '';
                    $riga['riferimentoNormativo'] = isset($row[25]) ? $row[25] : '';
                    $riga['riferimentoAmministrativo'] = isset($row[26]) ? $row[26] : '';
                    
                    $riga['tipoDato'] = isset($row[27]) ? $row[27] : '';
                    $riga['riferimentoTesto'] = isset($row[28]) ? $row[28] : '';
                    $riga['riferimentoNumero'] = isset($row[29]) ? $row[29] * 1: 0;
                    $riga['riferimentoData'] = isset($row[30]) ? Date::excelToDateTimeObject($row[30])->format('c') : '';
                    $righe[] = $riga;
                }
            }
            $fattura = ['testata' => $testata, 'righe' => $righe];
    
            echo json_encode(["invoice" => $fattura, "errorCode" => 0, "erroMessage" => '']);
        } catch( InvalidArgumentException $e ) {
            echo json_encode(["invoice" => [], "errorCode" => 200, "erroMessage" => $e->getMessage()]);
        }
    } else {
		echo json_encode(["invoice" => [], "errorCode" => 100, "errorMessage" => 'Nessun file name impostato']);
	}
