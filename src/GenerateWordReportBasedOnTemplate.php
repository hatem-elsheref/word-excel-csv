<?php

namespace Hatem\FilesApp;

use PhpOffice\PhpWord\TemplateProcessor;

class GenerateWordReportBasedOnTemplate
{
    private $templatePath = null;
    private $data = null;
    CONST SAVE_PATH = __DIR__ . DIRECTORY_SEPARATOR . 'storage' . DIRECTORY_SEPARATOR . 'results';

    public function __construct(string $templatePath, $data = null)
    {
        $this->templatePath = $templatePath;
        $this->data = $data;
    }

    public function generate(string $reportFileName = 'newFile.docs')
    {
        $templateProcessor = new TemplateProcessor($this->templatePath);
        $templateProcessor->setValue('reportName', $this->data['info']['reportName']);
        $templateProcessor->cloneRowAndSetValues('id', $this->data['rows']);
        $templateProcessor->saveAs(self::SAVE_PATH . DIRECTORY_SEPARATOR . $this->prepareFileName($reportFileName));
    }


    private function prepareFileName($fileName)
    {
        if(!empty($fileName)){
            $fileNameParts = explode(".", $fileName);
            if(end($fileNameParts) == 'docx'){
                array_splice($fileNameParts, -1, 1);
            }
            return implode('.', $fileNameParts) . '.docx';
        }
        return "";
    }

}