<?php


require_once __DIR__ . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';

use Hatem\FilesApp\WordTemplateProcessing;

$data = [
    'info' => [
        'reportName' => 'كشف بيانات لبلبلبلاب عن يوم الخميس الموافق 1/5/1999'
    ],
    'rows' => []
];
$faker = Faker\Factory::create('ar_SA');

foreach (range(1, 100) as $id){
    $data['rows'][] = [
        'id'         => $id,
        //'job'        => $faker->jobTitle(),
        'name'       => $faker->name('male'),
        'armyId'     => $faker->randomNumber(9),
        'nationalId' => $faker->randomNumber(9),
        'place'      => $faker->city
    ];
}


$template_path   = __DIR__ . DIRECTORY_SEPARATOR . 'src' . DIRECTORY_SEPARATOR . 'storage' . DIRECTORY_SEPARATOR . 'template.docx';
$largeExcel_path = __DIR__ . DIRECTORY_SEPARATOR . 'src' . DIRECTORY_SEPARATOR . 'storage' . DIRECTORY_SEPARATOR . 'largeExcel.xlsx';
$largeCsv_path   = __DIR__ . DIRECTORY_SEPARATOR . 'src' . DIRECTORY_SEPARATOR . 'storage' . DIRECTORY_SEPARATOR . 'largeCsv.csv';

/*
$wordApp = new WordTemplateProcessing($template_path, $data);
$wordApp->generate('Report_' . date('Ymd'));

$xlsxApp = new \Hatem\FilesApp\Excel();
$xlsxApp->splitLargeExcelSheetToMultipleFiles($largeExcel_path, 1, 2, 1000);
$xlsxApp->splitLargeCsvFileToMultipleFiles($largeCsv_path, 100);
$xlsxApp->convertCsvFileToXlsx($largeCsv_path, 'convertedFileFromCsvToXlsx');

*/

$xlsxApp = new \Hatem\FilesApp\Excel();
$xlsxApp->convertXlsxFileToCsv($largeExcel_path, 'convertedFileFromXlsxToCsv');

/*
#split xlsx to multiple by rows
#convert xlsx to csv
#split csv to multiple
#convert csv to xlsx
#make new template in word

colllect mulltiple csv into one
collect multiple xlsx into one file
make report from design support rtl and arabic language
mostafa soliman new app
insert fast in data base from excel sheet with loop



echo (php:://output)  what this mean
https://stackoverflow.com/questions/4348802/how-can-i-output-a-utf-8-csv-in-php-that-excel-will-read-properly
*/