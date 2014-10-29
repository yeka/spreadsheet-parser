<?php
error_reporting(E_ALL ^ E_NOTICE);
require_once 'excel_reader2.php';
$xls = new Spreadsheet_Excel_Reader("example.xls");
//$data = new Spreadsheet_Excel_Reader("PG-ID-DBAL-Master-v2.3.xlsx");
//print_r($xls->sheets[0]['cellsInfo']);
for ($row = 1; $row <= $xls->rowcount(); $row++)
{
    $rows = array();
    for ($col = 1; $col <= $xls->colcount(); $col++)
    {
        $type = $xls->type($row, $col);

        switch ($type) {
            case 'number':
                $val = $xls->raw($row, $col);
                break;
            case 'date':
                $raw = $xls->raw($row, $col);
                // Converting Excel Raw Date
                $days = floor($raw);
                $secs = round(($raw - $days) * 24 * 3600);
                $date = new DateTime('1899-12-30');

                $date->modify("+{$days} day");
                $date->modify("+{$secs} second");

                $val = $date;
                break;
            default:
                $val = $xls->val($row, $col);
        }
        $rows[] = $val;
//        if (isset($special_col[$col]))
//        {
//            switch($special_col[$col]) {
//                case 'raw':
//                    $rows[$col - 1] = $xls->raw($row, $col);
//                    break;
//                case 'val':
//                    $rows[$col - 1] = $xls->val($row, $col);
//                    break;
//                case 'format':
//                default:
//                    $format = $xls->format($row, $col);
//                    $val = $xls->val($row, $col);
//                    $rows[$col - 1] = substr($format.$val, - (strlen($format)));
//                    break;
//            }
//
//        }
//        elseif ($xls->type($row, $col) == 'number')
//        {
//            $rows[$col - 1] = $xls->raw($row, $col);
//        }
//        elseif ($xls->type($row, $col) == 'date')
//        {
//            $raw = $xls->raw($row, $col);
//            // Converting Excel Raw Date
//            $days = floor($raw);
//            $secs = round(($raw - $days) * 24 * 3600);
//            $date = new DateTime('1899-12-30');
//
//            //$i = new DateInterval("P{$days}DT{$secs}S");
//
//            //$date->add($i);
//            $date->modify("+{$days} day");
//            $date->modify("+{$secs} second");
//
//
//
//            $rows[$col - 1] = $date->format('Y-m-d');
//        }
//        else
//        {
//            $rows[$col - 1] = $xls->val($row, $col);
//        }
    }

    if ($row == 13) print_r($rows);
}