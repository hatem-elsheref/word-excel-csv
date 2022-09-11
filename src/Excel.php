<?php


namespace Hatem\FilesApp;


use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Reader\Xlsx as XlsxReader;
use PhpOffice\PhpSpreadsheet\Spreadsheet;

class Excel
{
    const SAVED_PATH = __DIR__ . DIRECTORY_SEPARATOR . 'storage' . DIRECTORY_SEPARATOR;

    public function splitLargeExcelSheetToMultipleFiles($filePath = null, $rowHeaderNumber = 0, $rowsToSkip = 1, $maxSize = 10, $savedDir = "splitData")
    {
        if (is_null($filePath))
            die("please, enter valid file path");

        $reader = new XlsxReader();
        $rows = $reader->load($filePath)->getActiveSheet()->toArray();
        $header = $rows[$rowHeaderNumber];
        array_splice($rows, 0, $rowsToSkip);
        $countOfSmallFiles = ceil(count($rows) / $maxSize);
        if (!file_exists(self::SAVED_PATH . $savedDir))
            mkdir(self::SAVED_PATH . $savedDir, 0777, true);

        for($i = 0; $i < $countOfSmallFiles; $i++){
            $currentData = [$header, ...array_splice($rows, 0, $maxSize)];
            $spreadsheet = new Spreadsheet();
            $sheet = $spreadsheet->getActiveSheet();
            $sheet->fromArray($currentData, NULL, 'A1');
            $writer = new Xlsx($spreadsheet);
            $writer->save( self::SAVED_PATH . $savedDir . DIRECTORY_SEPARATOR . $i . '.xlsx');
        }
    }

    public function splitLargeCsvFileToMultipleFiles($filePath = null, $maxSize = 10, $savedDir = "splitDataCsv")
    {

        $fileHandlerToRead = fopen($filePath, 'r') or die("filed to open given file");
        $header = []; $isStart = true; $fileIndex = 0; $counter = 0; $data = [];

        while($row = fgetcsv($fileHandlerToRead, 1000, ",")){
            if($isStart){
                $header = $row;
                $isStart = false;
                continue;
            }

            if (!file_exists(self::SAVED_PATH . $savedDir))
                mkdir(self::SAVED_PATH . $savedDir, 0777, true);

            if($counter == $maxSize){
                //start make new file then empty
                $fileHandlerToWrite = fopen(self::SAVED_PATH . $savedDir . DIRECTORY_SEPARATOR . $fileIndex . '.csv', 'w');
                fputcsv($fileHandlerToWrite, $header);
                foreach ($data as $rowData) {
                    fputcsv($fileHandlerToWrite, $rowData);
                }
                fclose($fileHandlerToWrite);
                $data = [];
                $counter = 0;
                $fileIndex++;
            }

            $data[] = $row;
            $counter++;
        }
    }

    public function convertXlsxFileToCsv($xlsxFile, $csvFileName)
    {
        #method1
        $reader = IOFactory::createReader("Xlsx");
        $spreadsheet = $reader->load($xlsxFile);
        $writer = IOFactory::createWriter($spreadsheet, "Csv");
        $writer->setSheetIndex(0);
        $writer->setDelimiter(';');
        $writer->save(self::SAVED_PATH . pathinfo($csvFileName, PATHINFO_FILENAME) . '.csv');


        #method2
        /*
        $reader = IOFactory::createReader("Xlsx");
        $rows = $reader->load($xlsxFile)->getActiveSheet()->toArray();
        $fileHandlerToWrite = fopen(self::SAVED_PATH . pathinfo($csvFileName, PATHINFO_FILENAME) . '.csv', 'w');
        foreach($rows as $index => $row){
            fputcsv($fileHandlerToWrite, $row);
        }
        fclose($fileHandlerToWrite);
        */
    }

    public function convertCsvFileToXlsx($csvFile, $xlsxFileName)
    {
        $reader = IOFactory::createReader('Csv');
        $file = $reader->load($csvFile);

        $writer = IOFactory::createWriter($file, 'Xlsx');
        $writer->save(self::SAVED_PATH . pathinfo($xlsxFileName, PATHINFO_FILENAME) . '.xlsx');
    }

    public function collectMultipleXlsxFilesInOneFile()
    {

    }


    public function collectMultipleCsvFilesInOneFile()
    {

    }

    public function insertRowsInDatabase()
    {

    }

    public function generateWordReportFromExcelFile()
    {

    }
    public function generateWordReportFromExcelFileBasedOnTemplate()
    {

    }


}