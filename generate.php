<?php

require 'vendor/autoload.php';
use PhpOffice\PhpSpreadsheet\Spreadsheet;

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Fill;


// connection
$connection = pg_connect('host=localhost port=5432 dbname=timesheet user=postgres password=postgres');
if (!$connection)
{
    exit("dabase connection error");
}

// get data from table
$data = pg_query($connection, "SELECT * FROM public.timesheets"); // modify query where id_user, start date, end date

// print_r(count(pg_fetch_array($data)));
// print_r(pg_fetch_all($data));

// load template
$reader = IOFactory::createReader('Xlsx');
$spreadsheet = new Spreadsheet();

$spreadsheet = $reader->load('celerates-template.xlsx');

// constant variable
$timesheet_sheet_index = 1;
$employee_name_cell = 'N1';
$timesheet_start_date_cell = 'D5';
$timesheet_end_date_cell = 'D6';

// fillable column
$date_column = "C";
$activity_column = "D";
$project_name_column = "E";
$work_desc_column = "F";
$start_time_column = "G";
$end_time_column = "H";
$break_hour_column = "I";



// dynamic variable
$employee_name = 'Purna Satria Nugraha';
$timesheet_start_date = '11-11-2021';
$timesheet_end_date = '11-12-2021';
$filename = $employee_name.$timesheet_start_date.$timesheet_start_date;

// insert data info
$spreadsheet->getSheet($timesheet_sheet_index)->setCellValue($employee_name_cell, $employee_name);
$spreadsheet->getSheet($timesheet_sheet_index)->setCellValue($timesheet_start_date_cell, $timesheet_start_date);
$spreadsheet->getSheet($timesheet_sheet_index)->setCellValue($timesheet_end_date_cell, $timesheet_end_date);

// loop insert timesheet data

$contentStartRow = 9;
$currentContentRow = 9;

while($item=pg_fetch_array($data)){
    print_r($item);
    // insert row before
    // $spreadsheet->getSheet($timesheet_sheet_index)->insertNewRowBefore($currentContentRow+1, 1);

    // fill data
    $spreadsheet->getSheet($timesheet_sheet_index)
        ->setCellValue($date_column.$currentContentRow, $item['date'])
        ->setCellValue($activity_column.$currentContentRow, $item['task'])
        ->setCellValue($project_name_column.$currentContentRow, $item['project_id'])
        ->setCellValue($start_time_column.$currentContentRow, $item['clockin'])
        ->setCellValue($end_time_column.$currentContentRow, $item['clockout'])
        ->setCellValue($break_hour_column.$currentContentRow, 1);

    $currentContentRow++;
}





$writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
$writer->save($filename.".xlsx")

?>