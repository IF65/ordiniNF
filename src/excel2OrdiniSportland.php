<?php

declare(strict_types=1);

@ini_set('memory_limit', '16384M');

require '../vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx;
use PhpOffice\PhpSpreadsheet\Shared\Date;

$timeZone = new DateTimeZone('Europe/Rome');

// verifico che il file sia stato effettivamente caricato
if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
    echo 'Non hai inviato nessun file...';
    //echo json_encode($_FILES, true);
    exit;
}

if (move_uploaded_file($_FILES['userfile']['tmp_name'], "/phpUpload/" . $_FILES['userfile']['name'])) {
    $inputFileName = "/phpUpload/" . $_FILES['userfile']['name'];
/*    if (true) {
        $inputFileName = "/Users/if65/Desktop/AKUSS23.xlsx";*/

    /** Create a new Xls Reader  **/
    $reader = new Xlsx();
    /** Load $inputFileName to a Spreadsheet Object  **/
    $reader->setReadDataOnly(true);
    $reader->setLoadAllSheets();

    $workbook = $reader->load($inputFileName);
    $worksheet = $workbook->getSheet(0);

    try {
        $highestRow = $worksheet->getHighestRow();
        $highestColumn = $worksheet->getHighestColumn();
        $highestColumnIndex = Coordinate::columnIndexFromString($highestColumn);

        // Definizione sedi
        $stores = [];
        for ($columnIndex = 53; $columnIndex <= 73; $columnIndex++) {
            $value = $worksheet->getCellByColumnAndRow($columnIndex, 1)->getValue();
            if ($value != '') {
                $stores[$columnIndex] = $value;
            }
        }

        // Definizione taglie
        $sizes = [];
        for ($rowIndex = 1; $rowIndex <= 9; $rowIndex++) {
            $sizeId = $worksheet->getCellByColumnAndRow(28, $rowIndex)->getValue() ?? '';
            if ($sizeId != '') {
                for ($columnIndex = 29; $columnIndex <= 49; $columnIndex++) {
                    $size = (string)$worksheet->getCellByColumnAndRow($columnIndex, $rowIndex)->getValue() ?? '';
                    if ($size != '') {
                        $sizes[$sizeId][$columnIndex] = $size;
                    }
                }
            }
        }

        /** @var  Article[] $articles */
        $articles = [];

        //Delivery date
        $deliveryDates = [];
        for ($index = 11; $index <= $highestRow; $index += 5) {
            $brand = $worksheet->getCellByColumnAndRow(2, $index)->getValue() ?? '';
            $minDeliveryDate = Date::excelToDateTimeObject($worksheet->getCellByColumnAndRow(25, 11)->getValue());
            $maxDeliveryDate = Date::excelToDateTimeObject($worksheet->getCellByColumnAndRow(26, 11)->getValue());

            if (!key_exists($brand, $deliveryDates)) {
                $deliveryDates[$brand] = ['min' => $minDeliveryDate, 'max' => $maxDeliveryDate];
            }
        }

        for ($index = 11; $index <= $highestRow; $index += 5) {
            $value = $worksheet->getCellByColumnAndRow(1, $index)->getValue();
            if ($value == '') {
                break;
            }

            $sizeId = $worksheet->getCellByColumnAndRow(28, $index)->getValue();

            $usedSizes = [];

            $packageClass1Items = [];
            $packageClass2Items = [];
            $packageClass3Items = [];
            $packageClass4Items = [];
            for ($columnIndex = 29; $columnIndex <= 49; $columnIndex++) {
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass1Items[$sizes[$sizeId][$columnIndex]] = ['quantity' => $value, 'articleCode' => ''];
                    $usedSizes[$sizes[$sizeId][$columnIndex]] = 0;
                }
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index + 1)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass2Items[$sizes[$sizeId][$columnIndex]] = ['quantity' => $value, 'articleCode' => ''];
                    $usedSizes[$sizes[$sizeId][$columnIndex]] = 0;
                }
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index + 2)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass3Items[$sizes[$sizeId][$columnIndex]] = ['quantity' => $value, 'articleCode' => ''];
                    $usedSizes[$sizes[$sizeId][$columnIndex]] = 0;
                }
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index + 3)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass4Items[$sizes[$sizeId][$columnIndex]] = ['quantity' => $value, 'articleCode' => ''];
                    $usedSizes[$sizes[$sizeId][$columnIndex]] = 0;
                }
            }

            /** @var string[] $packageClass1Stores */
            $packageClass1Stores = [];
            /** @var string[] $packageClass2Stores */
            $packageClass2Stores = [];
            /** @var string[] $packageClass3Stores */
            $packageClass3Stores = [];
            /** @var string[] $packageClass4Stores */
            $packageClass4Stores = [];
            for ($columnIndex = 53; $columnIndex <= 73; $columnIndex++) {
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index)->getValue() ?? '';
                if ($value != '') {
                    $packageClass1Stores[] = $stores[$columnIndex];
                }
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index + 1)->getValue() ?? '';
                if ($value != '') {
                    $packageClass2Stores[] = $stores[$columnIndex];
                }
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index + 2)->getValue() ?? '';
                if ($value != '') {
                    $packageClass3Stores[] = $stores[$columnIndex];
                }
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index + 3)->getValue() ?? '';
                if ($value != '') {
                    $packageClass4Stores[] = $stores[$columnIndex];
                }
            }

            $articles[] = new Article(
                $worksheet->getCellByColumnAndRow(13, $index)->getValue() ?? '',
                $worksheet->getCellByColumnAndRow(14, $index)->getValue() ?? '',
                $worksheet->getCellByColumnAndRow(1, $index)->getValue() ?? '',
                $worksheet->getCellByColumnAndRow(2, $index)->getValue() ?? '',
                (int)$worksheet->getCellByColumnAndRow(3, $index)->getValue() ?? 0,
                $worksheet->getCellByColumnAndRow(15, $index)->getValue() ?? '',
                $worksheet->getCellByColumnAndRow(4, $index)->getValue() ?? '', // settore
                $worksheet->getCellByColumnAndRow(5, $index)->getValue() ?? '',//reparto
                $worksheet->getCellByColumnAndRow(6, $index)->getValue() ?? '',
                $worksheet->getCellByColumnAndRow(7, $index)->getValue() ?? '',
                $worksheet->getCellByColumnAndRow(8, $index)->getValue() ?? '',
                (string)$worksheet->getCellByColumnAndRow(9, $index)->getValue() ?? '',
                (string)$worksheet->getCellByColumnAndRow(10, $index)->getValue() ?? '',
                $worksheet->getCellByColumnAndRow(16, $index)->getValue() ?? '',
                (float)$worksheet->getCellByColumnAndRow(17, $index)->getValue() ?? 0.0,
                (float)$worksheet->getCellByColumnAndRow(21, $index)->getCalculatedValue() ?? 0.0,
                (float)$worksheet->getCellByColumnAndRow(18, $index)->getValue() * 100 ?? 0.0,
                (float)$worksheet->getCellByColumnAndRow(19, $index)->getValue() * 100 ?? 0.0,
                (float)$worksheet->getCellByColumnAndRow(20, $index)->getValue() * 100 ?? 0.0,
                (float)$worksheet->getCellByColumnAndRow(22, $index)->getValue() ?? 0.0,
                (float)$worksheet->getCellByColumnAndRow(23, $index)->getValue() ?? 0.0,
                (float)$worksheet->getCellByColumnAndRow(24, $index)->getValue() ?? 0.0,
                $deliveryDates[$worksheet->getCellByColumnAndRow(2, $index)->getValue() ?? '']['min'],
                $deliveryDates[$worksheet->getCellByColumnAndRow(2, $index)->getValue() ?? '']['max'],
                $packageClass1Items,
                $packageClass2Items,
                $packageClass3Items,
                $packageClass4Items,
                $packageClass1Stores,
                $packageClass2Stores,
                $packageClass3Stores,
                $packageClass4Stores,
                array_keys($usedSizes)
            );
        }

        echo json_encode(array("recordCount" => count($articles), "values" => $articles));
        //file_put_contents('/Users/if65/Desktop/ordineSP.json', json_encode(array("recordCount" => count($articles), "values" => $articles)));
    } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
        die ("Error");
    }

    exit;
}

class Article
{
    public string $code;//Articolo
    public string $description;//Descrizione
    public string $supplierCode;//FORNITORE
    public string $brand;//MARCHIO
    public int $seasonCode;//STAGIONE
    public string $seasonDescription;//STAG
    public string $sector;//SETTORE
    public string $department;//REPARTO
    public string $gender; //GENDER
    public string $line;//LINEA
    public string $notes;//NOTE
    public string $familyCode;//FAM
    public string $subFamilyCode;//S.FAM
    public string $variation;//VARIANTE
    public float $grossAmount;// Acquisto Lordo
    public float $netAmount;// Acquisto Netto (prezzo finito)
    public float $discount1;//SCONTO 1
    public float $discount2;//SCONTO 2
    public float $discount3;//SCONTO 3
    public float $priceList1;//	LIST 1
    public float $priceList3;//LIST3
    public float $markUp;//Margine
    public DateTime $minDeliveryDate;//CONS MIN
    public DateTime $maxDeliveryDate;//CONS MAX
    public array $packageClass1Items;
    public array $packageClass2Items;
    public array $packageClass3Items;
    public array $packageClass4Items;

    public array $packageClass1Stores;
    public array $packageClass2Stores;
    public array $packageClass3Stores;
    public array $packageClass4Stores;

    public array $usedSizes;

    /**
     * @param string $code
     * @param string $description
     * @param string $supplierCode
     * @param string $brand
     * @param int $seasonCode
     * @param string $seasonDescription
     * @param string $sector
     * @param string $department
     * @param string $gender
     * @param string $line
     * @param string $notes
     * @param string $familyCode
     * @param string $subFamilyCode
     * @param string $variation
     * @param float $grossAmount
     * @param float $netAmount
     * @param float $discount1
     * @param float $discount2
     * @param float $discount3
     * @param float $priceList1
     * @param float $priceList3
     * @param float $markUp
     * @param DateTime $minDeliveryDate
     * @param DateTime $maxDeliveryDate
     * @param array $packageClass1Items
     * @param array $packageClass2Items
     * @param array $packageClass3Items
     * @param array $packageClass4Items
     * @param array $packageClass1Stores
     * @param array $packageClass2Stores
     * @param array $packageClass3Stores
     * @param array $packageClass4Stores
     * @param array $usedSizes
     */
    public function __construct(
        string $code,
        string $description,
        string $supplierCode,
        string $brand,
        int $seasonCode,
        string $seasonDescription,
        string $sector,
        string $department,
        string $gender,
        string $line,
        string $notes,
        string $familyCode,
        string $subFamilyCode,
        string $variation,
        float $grossAmount,
        float $netAmount,
        float $discount1,
        float $discount2,
        float $discount3,
        float $priceList1,
        float $priceList3,
        float $markUp,
        DateTime $minDeliveryDate,
        DateTime $maxDeliveryDate,
        array $packageClass1Items,
        array $packageClass2Items,
        array $packageClass3Items,
        array $packageClass4Items,
        array $packageClass1Stores,
        array $packageClass2Stores,
        array $packageClass3Stores,
        array $packageClass4Stores,
        array $usedSizes
    ) {
        $this->code = $code;
        $this->description = $description;
        $this->supplierCode = $supplierCode;
        $this->brand = $brand;
        $this->seasonCode = $seasonCode;
        $this->seasonDescription = $seasonDescription;
        $this->sector = $sector;
        $this->department = $department;
        $this->gender = $gender;
        $this->line = $line;
        $this->notes = $notes;
        $this->familyCode = $familyCode;
        $this->subFamilyCode = $subFamilyCode;
        $this->variation = $variation;
        $this->grossAmount = $grossAmount;
        $this->netAmount = $netAmount;
        $this->discount1 = $discount1;
        $this->discount2 = $discount2;
        $this->discount3 = $discount3;
        $this->priceList1 = $priceList1;
        $this->priceList3 = $priceList3;
        $this->markUp = $markUp;
        $this->minDeliveryDate = $minDeliveryDate;
        $this->maxDeliveryDate = $maxDeliveryDate;
        $this->packageClass1Items = $packageClass1Items;
        $this->packageClass2Items = $packageClass2Items;
        $this->packageClass3Items = $packageClass3Items;
        $this->packageClass4Items = $packageClass4Items;
        $this->packageClass1Stores = $packageClass1Stores;
        $this->packageClass2Stores = $packageClass2Stores;
        $this->packageClass3Stores = $packageClass3Stores;
        $this->packageClass4Stores = $packageClass4Stores;
        $this->usedSizes = $usedSizes;
    }


}