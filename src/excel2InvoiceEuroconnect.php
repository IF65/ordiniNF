<?php
    @ini_set('memory_limit','8192M');
    
    require '../vendor/autoload.php';

    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
    use PhpOffice\PhpSpreadsheet\Shared\Date;

    $debug = false;
    
    $timeZone = new DateTimeZone('Europe/Rome');
    
    $inputFileName = '';

    if ($debug) {
        $inputFileName = "/Users/if65/Desktop/test_fatture_excel.xlsx";
        $sheetName = ' ITMK 2019';
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
                    $cells[] = $cell->getCalculatedValue();//getValue();
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
            
            $testata['tipoRitenuta'] = isset($rows[39][3]) ? $rows[39][3] : '';
            $testata['importoRitenuta'] = isset($rows[40][3]) ? $rows[40][3] : '0';
            $testata['aliquotaRitenuta'] = isset($rows[41][3]) ? $rows[41][3] : '0';
            $testata['causalePagamentoRitenuta'] = isset($rows[42][3]) ? $rows[42][3] : '';
            
            $testata['codiceCommessaConvenzione'] = isset($rows[44][3]) ? $rows[44][3] : '';
            $testata['codiceCup'] = isset($rows[45][3]) ? $rows[45][3] : '';
            $testata['codiceCig'] = isset($rows[46][3]) ? $rows[46][3] : '';
            
            $testata['idDocumento'] = isset($rows[48][3]) ? $rows[48][3] : '';
            $testata['dataDocumento'] = isset($rows[49][3]) ? Date::excelToDateTimeObject($rows[49][3])->format('c') : '';

            $testata['condizioniPagamento'] = isset($rows[51][3]) ? $rows[51][3] : '';
            
            $testata['modalitaPagamento'] = isset($rows[53][3]) ? $rows[53][3] : '';
            $testata['dataPagamento'] = isset($rows[54][3]) ? Date::excelToDateTimeObject($rows[54][3])->format('c') : '';
            $testata['giorniPagamento'] = isset($rows[55][3]) ? $rows[55][3] : '';
            $testata['dataScadenzaPagamento'] = isset($rows[56][3]) ? Date::excelToDateTimeObject($rows[56][3])->format('c') : '';
            $testata['iban'] = isset($rows[57][3]) ? $rows[57][3] : '';
            
            $righe = [];
            foreach ($rows as $num => $row) {
                if ($num > 61) {
                    $riga = [];
                    
                    $riga['idDocumento'] = isset($row[0]) ? $row[0] : '';
                    $riga['data'] = isset($row[1]) ? Date::excelToDateTimeObject($row[1])->format('c') : '';
                    $riga['numItem'] = isset($row[2]) ? $row[2] : '';
                    $riga['idDocumento2'] = isset($row[3]) ? $row[3] : '';
                    $riga['data2'] = isset($row[4]) ? Date::excelToDateTimeObject($row[4])->format('c') : '';
                    
                    $riga['numeroDdt'] = isset($row[5]) ? $row[5] : '';
                    $riga['dataDdt'] = isset($row[6]) ? Date::excelToDateTimeObject($row[6])->format('c') : '';
                    $riga['tipoCessione'] = isset($row[7]) ? $row[7] : '';
                    
                    $riga['codice1Tipo'] = isset($row[8]) ? $row[8] : '';
                    $riga['codice1Valore'] = isset($row[9]) ? $row[9] : '';
                    $riga['codice2Tipo'] = isset($row[10]) ? $row[10] : '';
                    $riga['codice2Valore'] = isset($row[11]) ? $row[11] : '';
                    $riga['codice3Tipo'] = isset($row[12]) ? $row[12] : '';
                    $riga['codice3Valore'] = isset($row[13]) ? $row[13] : '';
                    
                    $riga['descrizione'] = isset($row[14]) ? $row[14] : '';
                    $riga['quantita'] = isset($row[15]) ? $row[15] * 1: 0;
                    $riga['unitaMisura'] = isset($row[16]) ? $row[16] : '';
                    $riga['prezzoUnitario'] = isset($row[17]) ? $row[17] * 1: 0;
                    
                    $riga['sconto1Tipo'] = isset($row[18]) ? $row[18] : '';
                    $riga['sconto1Percentuale'] = isset($row[19]) ? $row[19] * 100: 0;
                    $riga['sconto1Importo'] = isset($row[20]) ? $row[20] * 1: 0;
                    $riga['sconto2Tipo'] = isset($row[21]) ? $row[21] : '';
                    $riga['sconto2Percentuale'] = isset($row[22]) ? $row[22] * 1: 0;
                    $riga['sconto2Importo'] = isset($row[23]) ? $row[23] * 1: 0;
                    
                    $riga['prezzoTotale'] = isset($row[24]) ? $row[24] * 1: 0;
                    $riga['aliquotaIva'] = isset($row[25]) ? $row[25] : '0';
                    $riga['ritenuta'] = isset($row[26]) ? $row[26] : '';
                    $riga['natura'] = isset($row[27]) ? $row[27] : '';
                    $riga['riferimentoNormativo'] = isset($row[28]) ? $row[28] : '';
                    $riga['riferimentoAmministrativo'] = isset($row[29]) ? $row[29] : '';
                    
                    $riga['tipoDato'] = isset($row[30]) ? $row[30] : '';
                    $riga['riferimentoTesto'] = isset($row[31]) ? $row[31] : '';
                    $riga['riferimentoNumero'] = isset($row[32]) ? $row[32] * 1: 0;
                    $riga['riferimentoData'] = isset($row[33]) ? Date::excelToDateTimeObject($row[33])->format('c') : '';
                    
                    $riga['tipoDato2'] = isset($row[34]) ? $row[34] : '';
                    $riga['riferimentoTesto2'] = isset($row[35]) ? $row[35] : '';
                    $riga['riferimentoNumero2'] = isset($row[36]) ? $row[36] * 1: 0;
                    $riga['riferimentoData2'] = isset($row[37]) ? Date::excelToDateTimeObject($row[37])->format('c') : '';
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
