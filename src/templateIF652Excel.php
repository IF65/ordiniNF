<?php
//@ini_set('memory_limit','8192M');

require '../vendor/autoload.php';
// leggo i dati da un file
//$request = file_get_contents('/Users/if65/Desktop/Catalina/request.json');
$request = file_get_contents('php://input');
$data = json_decode($request, true);

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Cell\DataType;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Style;
use PhpOffice\PhpSpreadsheet\Style\NumberFormat;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;
use PhpOffice\PhpSpreadsheet\Style\Border;
use PhpOffice\PhpSpreadsheet\Shared\Date;

// verifico l'esistenza della cartella temp e se serve la creo
// con mask 777.
if (! file_exists ( '../temp' )) {
    $oldMask = umask(0);
    mkdir('../temp', 0777);
    umask($oldMask);
}

$timeZone = new DateTimeZone('Europe/Rome');

// leggo i parametri contenuti nel file
$nomeFile = $data['nomeFile'];
$file = '../temp/'.$nomeFile.'.xlsx';

$promozioni = $data['promozioni'];
$sedi = $data['sedi'];
$aderentiGruppi = $data['aderentiGruppi'];
$tipiPromozione = $data['tipoPromozione'];

// creazione del workbook
$workBook = new Spreadsheet();
$workBook->getDefaultStyle()->getFont()->setName('Arial');
$workBook->getDefaultStyle()->getFont()->setSize(12);
$workBook->getProperties()
    ->setCreator("IF65 S.p.A. (Gruppo Italmark)")
    ->setLastModifiedBy("IF65 S.p.A.")
    ->setTitle("Promozioni IF65")
    ->setSubject("Promozioni IF65")
    ->setDescription("Esportazione Promozioni IF65")
    ->setKeywords("office 2007 openxml php")
    ->setCategory("SM Docs");

$style = new Style();

$sheetCount = $workBook->getSheetCount();
while ($sheetCount <> 0) {
    $workBook->removeSheetByIndex($sheetCount - 1);
    $sheetCount = $workBook->getSheetCount();
}

$workBook->getDefaultStyle()->getAlignment()->setVertical( 'center' );

sort($tipiPromozione);
foreach ($tipiPromozione as $sheetCount => $tipoPromozione) {
    if (preg_match( '/^(....)$/', $tipoPromozione ) && $tipoPromozione != '0000') {
        $sheet = $workBook->createSheet()->setTitle( $tipoPromozione );
        $sheet->getDefaultRowDimension()->setRowHeight( 20 );
        $sheet->getRowDimension( 1 )->setRowHeight( 40 );


        // 0054 - Sconto Set in %
        if ($tipoPromozione == '0054') {
            $rigaDiTestataTesto = [
                'id Catalina', '#', 'P.Var', 'Denominazione', '%', 'Inizio', 'Fine', "Gruppo\naderenti", 'Aderenti',
                "Barcode\nGruppo 1", "Barcode\nGruppo 2", "Barcode\nGruppo 3"
            ];
            $rigaDiTestataLarghezza = [
                12, 12, 7, 40, 7, 10, 10, 15, 25, 20, 20, 20
            ];
        } elseif ($tipoPromozione == '0481') {
            $rigaDiTestataTesto = [
                'id Catalina', '#', 'P.Var', 'Denominazione', "Soglia", "Importo", 'Inizio', 'Fine', "Gruppo\naderenti", 'Aderenti', 'Reparti'
            ];
            $rigaDiTestataLarghezza = [
                12, 12, 7, 40, 10, 10, 10, 10, 15, 25, 20, 20
            ];
        } elseif ($tipoPromozione == '0503') {
            $rigaDiTestataTesto = [
                'id Catalina', '#', 'P.Var', 'Denominazione', "Soglia", "Importo", 'Inizio', 'Fine', "Gruppo\naderenti", 'Aderenti'
            ];
            $rigaDiTestataLarghezza = [
                12, 12, 7, 40, 10, 10, 10, 10, 15, 25, 20
            ];
        } else {
            $rigaDiTestataTesto = ['id Catalina', '#', 'P.Var', 'Denominazione'];
            $rigaDiTestataLarghezza = [ 12, 12, 7, 40];
        }

        for ($i = 0; $i < count( $rigaDiTestataTesto ); $i++) {
            $sheet->setCellValueExplicitByColumnAndRow($i + 1, 1, strtoupper( $rigaDiTestataTesto[$i]),  DataType::TYPE_STRING);
            $sheet->getColumnDimensionByColumn($i + 1)->setWidth( $rigaDiTestataLarghezza[$i] );
        }
        $range = Coordinate::stringFromColumnIndex( 1 ) . '1:' . Coordinate::stringFromColumnIndex( count( $rigaDiTestataTesto ) ) . '1';
        $sheet->getStyle( $range )
            ->getAlignment()
            ->setHorizontal( 'center' )
            ->setVertical( 'center' )
            ->setWrapText( true );
        $sheet->getStyle( $range )
            ->getFont()
            ->setBold( true );
    }
}

foreach ($promozioni as $promozione) {

    $sheet = $workBook->getSheetByName($promozione['tipo']);
    if ($sheet <> null) {
        $gruppiBarcode = [];
        foreach ($promozione['articoli'] as $articolo) {
            if (key_exists($articolo['gruppo'], $gruppiBarcode)) {
                $gruppiBarcode[$articolo['gruppo']]['barcode'][] =$articolo['barcode'];
            } else {
                $gruppiBarcode[$articolo['gruppo']] = ['barcode' => [$articolo['barcode']], 'molteplicita' => $articolo['molteplicita']];
            }
            sort($gruppiBarcode[$articolo['gruppo']]['barcode']);
        }
        $reparti = [];
        foreach ($promozione['articoli'] as $articolo) {
            $reparti[] = $articolo['codiceReparto'];
        }
        sort($reparti);

        $aderenti = $promozione['sedi'];
        sort($aderenti);

        $gruppoAderentiSelezionato = '';
        foreach ($aderentiGruppi as $aderenteDescrizione => $aderenteSedi) {
            if (count($aderenti) == count($aderenteSedi) && array_diff($aderenti, $aderenteSedi) === array_diff($aderenteSedi, $aderenti)) {
                $gruppoAderentiSelezionato = $aderenteDescrizione;
            }
        }

        $newRowIndex = 1 + $sheet->getHighestRow();

        $sheet->setCellValueExplicitByColumnAndRow( 1, $newRowIndex, $promozione['codiceCatalina'], DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( 1, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

        $sheet->setCellValueExplicitByColumnAndRow( 2, $newRowIndex, $promozione['codice'], DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( 2, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

        $sheet->setCellValueExplicitByColumnAndRow( 3, $newRowIndex, $promozione['ricompense'][0]['promovar'], DataType::TYPE_NUMERIC );
        $sheet->getStyleByColumnAndRow( 3, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

        $sheet->setCellValueExplicitByColumnAndRow( 4, $newRowIndex, $promozione['descrizione'], DataType::TYPE_STRING );

        if ($promozione['tipo'] == '0054') {
            $sheet->setCellValueExplicitByColumnAndRow( 5, $newRowIndex, $promozione['ricompense'][0]['ammontare'] / 100, DataType::TYPE_NUMERIC );
            $sheet->getStyleByColumnAndRow( 5, $newRowIndex )->getNumberFormat()->setFormatCode('0.00%');
            $sheet->getStyleByColumnAndRow( 5, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

            $sheet->setCellValueByColumnAndRow( 6, $newRowIndex, $promozione['dataInizio'] );
            $sheet->getStyleByColumnAndRow( 6, $newRowIndex )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_YYYYMMDD2 );
            $sheet->getStyleByColumnAndRow( 6, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

            $sheet->setCellValueByColumnAndRow( 7, $newRowIndex, $promozione['dataFine'] );
            $sheet->getStyleByColumnAndRow( 7, $newRowIndex )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_YYYYMMDD2 );
            $sheet->getStyleByColumnAndRow( 7, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

            $sheet->setCellValueExplicitByColumnAndRow( 8, $newRowIndex, $gruppoAderentiSelezionato, DataType::TYPE_STRING );
            $sheet->getStyleByColumnAndRow( 8, $newRowIndex )
                ->getAlignment()
                ->setHorizontal( 'center' )
                ->setVertical( 'center' )
                ->setWrapText( true );

            if(count($aderenti) && $gruppoAderentiSelezionato == '') {
                $sheet->setCellValueExplicitByColumnAndRow( 9, $newRowIndex, implode(";", $aderenti), DataType::TYPE_STRING );
                $sheet->getStyleByColumnAndRow( 9, $newRowIndex )
                    ->getAlignment()
                    ->setHorizontal( 'center' )
                    ->setVertical( 'center' )
                    ->setWrapText( true );
            }

            if(key_exists('1', $gruppiBarcode)) {
                $sheet->setCellValueExplicitByColumnAndRow( 10, $newRowIndex, implode(";", $gruppiBarcode['1']['barcode']), DataType::TYPE_STRING );
                $sheet->getStyleByColumnAndRow( 10, $newRowIndex )
                    ->getAlignment()
                    ->setHorizontal( 'center' )
                    ->setVertical( 'center' )
                    ->setWrapText( true );
            }
            if(key_exists('2', $gruppiBarcode)) {
                $sheet->setCellValueExplicitByColumnAndRow( 11, $newRowIndex, implode(";", $gruppiBarcode['2']['barcode']), DataType::TYPE_STRING );
                $sheet->getStyleByColumnAndRow( 11, $newRowIndex )
                    ->getAlignment()
                    ->setHorizontal( 'center' )
                    ->setVertical( 'center' )
                    ->setWrapText( true );
            }
            if(key_exists('3', $gruppiBarcode)) {
                $sheet->setCellValueExplicitByColumnAndRow( 12, $newRowIndex, implode(";", $gruppiBarcode['1']['barcode']), DataType::TYPE_STRING );
                $sheet->getStyleByColumnAndRow( 12, $newRowIndex )
                    ->getAlignment()
                    ->setHorizontal( 'center' )
                    ->setVertical( 'center' )
                    ->setWrapText( true );
            }
        } elseif ($promozione['tipo'] == '0481') {
            $sheet->setCellValueExplicitByColumnAndRow( 5, $newRowIndex, $promozione['ricompense'][0]['soglia'] , DataType::TYPE_NUMERIC );
            $sheet->getStyleByColumnAndRow( 5, $newRowIndex )->getNumberFormat()->setFormatCode(numberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $sheet->getStyleByColumnAndRow( 5, $newRowIndex )->getAlignment()->setHorizontal( 'right' );

            $sheet->setCellValueExplicitByColumnAndRow( 6, $newRowIndex, $promozione['ricompense'][0]['ammontare'] , DataType::TYPE_NUMERIC );
            $sheet->getStyleByColumnAndRow( 6, $newRowIndex )->getNumberFormat()->setFormatCode(numberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $sheet->getStyleByColumnAndRow( 6, $newRowIndex )->getAlignment()->setHorizontal( 'right' );

            $sheet->setCellValueByColumnAndRow( 7, $newRowIndex, $promozione['dataInizio'] );
            $sheet->getStyleByColumnAndRow( 7, $newRowIndex )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_YYYYMMDD2 );
            $sheet->getStyleByColumnAndRow( 7, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

            $sheet->setCellValueByColumnAndRow( 8, $newRowIndex, $promozione['dataFine'] );
            $sheet->getStyleByColumnAndRow( 8, $newRowIndex )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_YYYYMMDD2 );
            $sheet->getStyleByColumnAndRow( 8, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

            $sheet->setCellValueExplicitByColumnAndRow( 9, $newRowIndex, $gruppoAderentiSelezionato, DataType::TYPE_STRING );
            $sheet->getStyleByColumnAndRow( 9, $newRowIndex )
                ->getAlignment()
                ->setHorizontal( 'center' )
                ->setVertical( 'center' )
                ->setWrapText( true );

            if(count($aderenti) && $gruppoAderentiSelezionato == '') {
                $sheet->setCellValueExplicitByColumnAndRow( 10, $newRowIndex, implode(";", $aderenti), DataType::TYPE_STRING );
                $sheet->getStyleByColumnAndRow( 10, $newRowIndex )
                    ->getAlignment()
                    ->setHorizontal( 'center' )
                    ->setVertical( 'center' )
                    ->setWrapText( true );
            }

            $sheet->setCellValueExplicitByColumnAndRow( 11, $newRowIndex, implode(";", $reparti), DataType::TYPE_STRING );
            $sheet->getStyleByColumnAndRow( 11, $newRowIndex )
                ->getAlignment()
                ->setHorizontal( 'center' )
                ->setVertical( 'center' )
                ->setWrapText( true );
        } elseif ($promozione['tipo'] == '0503') {
            $sheet->setCellValueExplicitByColumnAndRow( 5, $newRowIndex, $promozione['ricompense'][0]['soglia'] , DataType::TYPE_NUMERIC );
            $sheet->getStyleByColumnAndRow( 5, $newRowIndex )->getNumberFormat()->setFormatCode(numberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $sheet->getStyleByColumnAndRow( 5, $newRowIndex )->getAlignment()->setHorizontal( 'right' );

            $sheet->setCellValueExplicitByColumnAndRow( 6, $newRowIndex, $promozione['ricompense'][0]['ammontare'] , DataType::TYPE_NUMERIC );
            $sheet->getStyleByColumnAndRow( 6, $newRowIndex )->getNumberFormat()->setFormatCode(numberFormat::FORMAT_NUMBER_COMMA_SEPARATED1);
            $sheet->getStyleByColumnAndRow( 6, $newRowIndex )->getAlignment()->setHorizontal( 'right' );

            $sheet->setCellValueByColumnAndRow( 7, $newRowIndex, $promozione['dataInizio'] );
            $sheet->getStyleByColumnAndRow( 7, $newRowIndex )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_YYYYMMDD2 );
            $sheet->getStyleByColumnAndRow( 7, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

            $sheet->setCellValueByColumnAndRow( 8, $newRowIndex, $promozione['dataFine'] );
            $sheet->getStyleByColumnAndRow( 8, $newRowIndex )->getNumberFormat()->setFormatCode( NumberFormat::FORMAT_DATE_YYYYMMDD2 );
            $sheet->getStyleByColumnAndRow( 8, $newRowIndex )->getAlignment()->setHorizontal( 'center' );

            $sheet->setCellValueExplicitByColumnAndRow( 9, $newRowIndex, $gruppoAderentiSelezionato, DataType::TYPE_STRING );
            $sheet->getStyleByColumnAndRow( 9, $newRowIndex )
                ->getAlignment()
                ->setHorizontal( 'center' )
                ->setVertical( 'center' )
                ->setWrapText( true );

            if(count($aderenti) && $gruppoAderentiSelezionato == '') {
                $sheet->setCellValueExplicitByColumnAndRow( 10, $newRowIndex, implode(";", $aderenti), DataType::TYPE_STRING );
                $sheet->getStyleByColumnAndRow( 10, $newRowIndex )
                    ->getAlignment()
                    ->setHorizontal( 'center' )
                    ->setVertical( 'center' )
                    ->setWrapText( true );
            }
        }
    }

}

$writer = new Xlsx($workBook);
$writer->save($file);

if (file_exists($file)) {
    header('Content-Description: File Transfer');
    header('Content-Type: application/octet-stream');
    header('Content-Disposition: attachment; filename="'.basename($file).'"');
    header('Expires: 0');
    header('Cache-Control: must-revalidate');
    header('Pragma: public');
    header('Content-Length: ' . filesize($file));
    readfile($file);
    exit;
}