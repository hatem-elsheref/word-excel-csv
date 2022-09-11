<?php


require_once __DIR__ . DIRECTORY_SEPARATOR . 'vendor' . DIRECTORY_SEPARATOR . 'autoload.php';

use Hatem\FilesApp\GenerateWordReportBasedOnTemplate;

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


$template_path = __DIR__ . DIRECTORY_SEPARATOR . 'src' . DIRECTORY_SEPARATOR . 'storage' . DIRECTORY_SEPARATOR . 'template.docx';

$wordApp = new GenerateWordReportBasedOnTemplate($template_path, $data);
$wordApp->generate('report-' . date('d-m-Y'));