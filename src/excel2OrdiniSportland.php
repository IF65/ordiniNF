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
/*if (true) {
    $inputFileName = "/Users/if65/Desktop/Sportland/Ordini per Test caricamento _Modello 2023_/Mod.2023 x Prova_Garmin.xlsx";*/

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
                    $size = $worksheet->getCellByColumnAndRow($columnIndex, $rowIndex)->getValue() ?? '';
                    if ($size != '') {
                        $sizes[$sizeId][$columnIndex] = $size;
                    }
                }
            }
        }

        /** @var  Article[] $articles */
        $articles = [];

        $minDeliveryDate = Date::excelToDateTimeObject($worksheet->getCellByColumnAndRow(25, 11)->getValue());
        $maxDeliveryDate = Date::excelToDateTimeObject($worksheet->getCellByColumnAndRow(26, 11)->getValue());

        for ($index = 11; $index <= $highestRow; $index += 5) {
            $value = $worksheet->getCellByColumnAndRow(1, $index)->getValue();
            if ($value == '') {
                break;
            }

            $sizeId = $worksheet->getCellByColumnAndRow(28, $index)->getValue();

            /** @var PackageItem[] $packageClass1Items */
            $packageClass1Items = [];
            /** @var PackageItem[] $packageClass2Items */
            $packageClass2Items = [];
            /** @var PackageItem[] $packageClass3Items */
            $packageClass3Items = [];
            /** @var PackageItem[] $packageClass4Items */
            $packageClass4Items = [];
            for ($columnIndex = 29; $columnIndex <= 49; $columnIndex++) {
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass1Items[] = new PackageItem($sizeId, $sizes[$sizeId][$columnIndex], $value);
                }
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index + 1)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass2Items[] = new PackageItem($sizeId, $sizes[$sizeId][$columnIndex], $value);
                }
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index + 2)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass3Items[] = new PackageItem($sizeId, $sizes[$sizeId][$columnIndex], $value);
                }
                $value = (int)$worksheet->getCellByColumnAndRow($columnIndex, $index + 3)->getValue() ?? 0;
                if ($value != 0) {
                    $packageClass4Items[] = new PackageItem($sizeId, $sizes[$sizeId][$columnIndex], $value);
                }
            }

            /** @var string[] $PackageClass1Stores */
            $PackageClass1Stores = [];
            /** @var string[] $PackageClass2Stores */
            $PackageClass2Stores = [];
            /** @var string[] $PackageClass3Stores */
            $PackageClass3Stores = [];
            /** @var string[] $PackageClass4Stores */
            $PackageClass4Stores = [];
            for ($columnIndex = 53; $columnIndex <= 73; $columnIndex++) {
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index)->getValue() ?? '';
                if ($value != '') {
                    $PackageClass1Stores[] = $stores[$columnIndex];
                }
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index + 1)->getValue() ?? '';
                if ($value != '') {
                    $PackageClass2Stores[] = $stores[$columnIndex];
                }
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index + 2)->getValue() ?? '';
                if ($value != '') {
                    $PackageClass3Stores[] = $stores[$columnIndex];
                }
                $value = $worksheet->getCellByColumnAndRow($columnIndex, $index + 3)->getValue() ?? '';
                if ($value != '') {
                    $PackageClass4Stores[] = $stores[$columnIndex];
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
                $minDeliveryDate,
                $maxDeliveryDate,
                $packageClass1Items,
                $packageClass2Items,
                $packageClass3Items,
                $packageClass4Items,
                $PackageClass1Stores,
                $PackageClass2Stores,
                $PackageClass3Stores,
                $PackageClass4Stores
            );
        }

        echo json_encode(array("recordCount" => count($articles), "values" => $articles));
        //file_put_contents('/Users/if65/Desktop/ordineSP.json', json_encode(array("recordCount" => count($articles), "values" => $articles)));
    } catch (\PhpOffice\PhpSpreadsheet\Exception $e) {
        die ("Error");
    }

    exit;
}

class PackageItem
{
    public string $code;
    public string $size;
    public int $quantity;

    /**
     * @param string $code
     * @param string $size
     * @param int $quantity
     */
    public function __construct(string $code, string $size, int $quantity)
    {
        $this->code = $code;
        $this->size = $size;
        $this->quantity = $quantity;
    }


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

    /** @var PackageItem[] $packageClass1Items */
    public array $packageClass1Items;
    /** @var PackageItem[] $packageClass2Items */
    public array $packageClass2Items;
    /** @var PackageItem[] $packageClass3Items */
    public array $packageClass3Items;
    /** @var PackageItem[] $packageClass4Items */
    public array $packageClass4Items;
    /** @var string[] $PackageClass1Stores */
    public array $PackageClass1Stores;
    /** @var string[] $PackageClass2Stores */
    public array $PackageClass2Stores;
    /** @var string[] $PackageClass3Stores */
    public array $PackageClass3Stores;
    /** @var string[] $PackageClass4Stores */
    public array $PackageClass4Stores;

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
     * @param PackageItem[] $packageClass1Items
     * @param PackageItem[] $packageClass2Items
     * @param PackageItem[] $packageClass3Items
     * @param PackageItem[] $packageClass4Items
     * @param string[] $PackageClass1Stores
     * @param string[] $PackageClass2Stores
     * @param string[] $PackageClass3Stores
     * @param string[] $PackageClass4Stores
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
        array $PackageClass1Stores,
        array $PackageClass2Stores,
        array $PackageClass3Stores,
        array $PackageClass4Stores
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
        $this->PackageClass1Stores = $PackageClass1Stores;
        $this->PackageClass2Stores = $PackageClass2Stores;
        $this->PackageClass3Stores = $PackageClass3Stores;
        $this->PackageClass4Stores = $PackageClass4Stores;
    }

}