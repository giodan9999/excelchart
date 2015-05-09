<?php

require_once("PHPExcel/Classes/PHPExcel.php");

$ea = new PHPExcel();

$ea->getProperties()
   ->setCreator('Taylor Ren')
   ->setTitle('PHPExcel Demo')
   ->setLastModifiedBy('Taylor Ren')
   ->setDescription('A demo to show how to use PHPExcel to manipulate an Excel file')
   ->setSubject('PHP Excel manipulation')
   ->setKeywords('excel php office phpexcel lakers')
   ->setCategory('programming');

$ews = $ea->getSheet(0);
$ews->setTitle('Data');

$ews->setCellValue('a1', 'ID');
$ews->setCellValue('b1', 'Date');
$ews->setCellValue('c1', 'Value');

$data = array(array('1','2015-05-06','30.6'),
    array('2','2015-05-07','33.6'),
    array('3','2015-05-08','28.6'),
    array('4','2015-05-09','43.6'),
    array('5','2015-05-10','53.6'),
    array('6','2015-05-11','23.6'));

foreach($data as $key => $item) {
    $ews->setCellValue('a'.($key+2), $item[0]);
    $ews->setCellValue('b'.($key+2), $item[1]);
    $ews->setCellValue('c'.($key+2), $item[2]);
}

$data=array();
$ews->fromArray($data, ' ', 'A2');

$dsl = array(new PHPExcel_Chart_DataSeriesValues('String', 'Data!$C$1', NULL, 1));
$xal = array(new PHPExcel_Chart_DataSeriesValues('String', 'Data!$B$2:$B$7', NULL, 6));
$dsv = array(new PHPExcel_Chart_DataSeriesValues('String', 'Data!$C$2:$C$7', NULL, 6));

$ds = new PHPExcel_Chart_DataSeries(
    PHPExcel_Chart_DataSeries::TYPE_LINECHART,
    PHPExcel_Chart_DataSeries::GROUPING_STANDARD,
    range(0, count($dsv)-1),
    $dsl,
    $xal,
    $dsv
);

$pa=new PHPExcel_Chart_PlotArea(NULL, array($ds));
$legend=new PHPExcel_Chart_Legend(PHPExcel_Chart_Legend::POSITION_RIGHT, NULL, false);
$title=new PHPExcel_Chart_Title('BaoBiao');

$chart= new PHPExcel_Chart(
    'chart1',
    $title,
    $legend,
    $pa,
    true,
    0,
    NULL,
    NULL
);

$chart->setTopLeftPosition('D1');
$chart->setBottomRightPosition('Q25');
$ews->addChart($chart);

$writer = PHPExcel_IOFactory::createWriter($ea, 'Excel2007');
$writer->setIncludeCharts(true);
$writer->save('output.xlsx');


?>