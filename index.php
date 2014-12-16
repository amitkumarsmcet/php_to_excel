<?php

function getNameFromNumber($num) {
		$numeric = $num % 26;
		$letter = chr(65 + $numeric);
		$num2 = intval($num / 26);
		if ($num2 > 0) {
			return getNameFromNumber($num2 - 1) . $letter;
		} else {
			return $letter;
		}
	}
/* data to display in excel */
$data = array(
	'0' => array('1','images/1.jpg','Jhon','22 Jump Street','Taxas','USA','jhon@gmail.com','789654123','Active'),
	'1' => array('2','images/2.jpg','Rob','21 Jump Street','Washington','USA','rob@gmail.com','984563217','Inactive'),
	'2' => array('3','images/3.jpg','Mike','Downtown','NewYork','USA','mike@gmail.com','784563210','Active')
	);

$header_row = array("Sno", "Image", "Name", "Address", "City", "Country","Email","Phone","Status"); // Define custom header row 
$header_width = array(10,15,20,25,20,20,40,15,15); // define custom header width

// include classes of excel
include('PHPExcel.php');
include('PHPExcel/Cell/AdvancedValueBinder.php');
include('PHPExcel/IOFactory.php');

// define report title
$ReportTitle ='User List Using PHP Excel';

// load sample of excel file
PHPExcel_Cell::setValueBinder( new PHPExcel_Cell_AdvancedValueBinder() );
$objPHPExcel = PHPExcel_IOFactory::load('user_list.xlsx');
			
$objPHPExcel->getProperties()->setCreator($ReportTitle);
$objPHPExcel->getProperties()->setLastModifiedBy($ReportTitle);
$objPHPExcel->getProperties()->setTitle($ReportTitle);
$objPHPExcel->getProperties()->setSubject($ReportTitle);
$objPHPExcel->getProperties()->setDescription($ReportTitle);
$objPHPExcel->getProperties()->setKeywords("office 2007 openxml php");
$objPHPExcel->getProperties()->setCategory($ReportTitle);
$objColGCondition = new PHPExcel_Style_Conditional();
$objPHPExcel->setActiveSheetIndex(0);
			
$col = 0;
$col_name = 0;
$flag	=	0;

foreach($header_row as $head_name){	
	$address	= getNameFromNumber($col_name++);
	
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getBorders()->getRight()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getBorders()->getLeft()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getBorders()->getBottom()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getBorders()->getTop()->setBorderStyle(PHPExcel_Style_Border::BORDER_THIN);

	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getAlignment()->setHorizontal(PHPExcel_Style_Alignment::HORIZONTAL_CENTER);

	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getAlignment()->setWrapText(true);

	// setColor
	// set heading COLOR
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getFont()->getColor()->setARGB(PHPExcel_Style_Color::COLOR_WHITE);
	// Set heading Allignment
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getAlignment()->setVertical(PHPExcel_Style_Alignment::VERTICAL_CENTER);
	// FORMATING
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1' )->getFill()->setFillType(PHPExcel_Style_Fill::FILL_SOLID);
	$objPHPExcel->getActiveSheet()->getStyle( $address.'1'  )->getFill()->getStartColor()->setARGB('FF0070C0');		

	$objPHPExcel->getActiveSheet()->getColumnDimension($address,1)->setWidth($header_width[$col] );
	//$objPHPExcel->getActiveSheet()->_text_wrap = 1;
	 

	$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow( ($col) , 1, $head_name, $pDataType = PHPExcel_Cell_DataType::TYPE_STRING);
	$col++;

}

$row_num = 2;
foreach($data as $row => $row_data){				
	foreach($row_data as $col_name => $cell_value){
		$objPHPExcel->getActiveSheet()->getRowDimension($row_num)->setRowHeight(50);
		if(is_numeric($cell_value) ){
			$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow( ($col_name) , ($row_num), $cell_value, $pDataType = PHPExcel_Cell_DataType::TYPE_NUMERIC);
		}else{
			if( $col_name == 1 ){
				$objDrawing = new PHPExcel_Worksheet_Drawing();
				$objDrawing->setName('PHPExcel logo');
				$objDrawing->setDescription('PHPExcel logo');
				$objDrawing->setPath($cell_value);       // filesystem reference for the image file
				$objDrawing->setHeight(50);                 // sets the image height to 36px (overriding the actual image height); 
				$objDrawing->setCoordinates(getNameFromNumber($col_name).$row_num);    // pins the top-left corner of the image to cell D24
				$objDrawing->setOffsetX(10);                // pins the top left corner of the image at an offset of 10 points horizontally to the right of the top-left corner of the cell
				$objDrawing->setOffsetY(10);
				$objDrawing->setWorksheet($objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow());
			}else{
				$objPHPExcel->getActiveSheet()->setCellValueExplicitByColumnAndRow( ($col_name) , ($row_num), $cell_value, $pDataType = PHPExcel_Cell_DataType::TYPE_STRING);
			}
		}		
	}
	$row_num++;
}
//die;

$file_name	= 'user_list'.date('m-d-Y');

header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
header('Content-Disposition: attachment;filename='.$file_name.'.xlsx');
header('Cache-Control: max-age=0');
$objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
$objWriter->save('php://output');
exit;
?>