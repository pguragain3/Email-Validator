<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use Egulias\EmailValidator\EmailValidator;
use Egulias\EmailValidator\Validation\RFCValidation;
use Egulias\EmailValidator\Validation\DNSCheckValidation;
use Egulias\EmailValidator\Validation\MessageIDValidation;
use Egulias\EmailValidator\Validation\NoRFCWarningsValidation;
use Egulias\EmailValidator\Validation\MultipleValidationWithAnd;


$validator = new EmailValidator();
$multipleValidations = new MultipleValidationWithAnd([
    new RFCValidation(),
    new DNSCheckValidation(),
    new NoRFCWarningsValidation(),
    new MessageIDValidation(),
]);



// Let's traverse the images directory
$fileSystemIterator = new FilesystemIterator('inputs');

foreach ($fileSystemIterator as $fileInfo) {
    $filename = $fileInfo->getFilename();
    $reader = new \PhpOffice\PhpSpreadsheet\Reader\Xlsx();
    $reader->setReadDataOnly(true); 
    $reader->setReadEmptyCells(false);
    $spreadsheet = $reader->load("inputs/" . $filename);

    $d = $spreadsheet->getSheet(0)->toArray();

    $sheetData = $spreadsheet->getActiveSheet()->toArray();
    $i = 1;
    unset($sheetData[0]);
    $array = [];
    foreach ($sheetData as $t) {
        // process element here;
        //ietf.org has MX records signaling a server with email capabilities
        $check = $validator->isValid($t[0] ?? 'null', $multipleValidations); //true
        $data_from_excel[] = ['Email' => $t[0], 'Validity' => $check == 1 ? 'True' : 'False'];
        $i++;
    }
    // Creates New Spreadsheet 
    $spreadsheet = new Spreadsheet();

    // Retrieve the current active worksheet 
    $sheet = $spreadsheet->getActiveSheet();

    //set your own column header
    $column_header = ["Email", "Validity"];
    $j = 1;
    foreach ($column_header as $x_value) {
        $sheet->setCellValueByColumnAndRow($j, 1, $x_value);
        $j = $j + 1;
    }

    //set value row
    for ($i = 0; $i < count($data_from_excel); $i++) {

        //set value for indi cell
        $row = $data_from_excel[$i];

        $j = 1;

        foreach ($row as $x => $x_value) {
            $sheet->setCellValueByColumnAndRow($j, $i + 2, $x_value);
            $j = $j + 1;
        }
    }

    $spreadsheet->getActiveSheet()->getColumnDimension('A')->setWidth(40);
    $spreadsheet->getActiveSheet()->getColumnDimension('B')->setWidth(30);
    // Write an .xlsx file  
    $writer = new Xlsx($spreadsheet);

    // Save .xlsx file to the files directory 
    $writer->save('outputs/' . $filename);
}
