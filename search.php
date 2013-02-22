<?php

//Below is my classes file I use to connect to DBs.  You can comment out the next line if you have your own model and edit the DB connection strings below if you have 
//your own model and/or prefer to use the standard PHP mysql or mssql connection strings
include_once "classes.php";
//This file should be placed in the PHPExcel root directory, you will need to reference the below files
include_once "Classes/PHPExcel/Reader/Excel2007.php";
include_once "Classes/PHPExcel/Reader/Excel5.php";
include_once "Classes/PHPExcel.php";
include_once "Classes/PHPExcel/Writer/Excel2007.php";

//Create a new object for the PHPExcel reader to operate on
$objReader = new PHPExcel_Reader_Excel2007();
$objReader->setReadDataOnly(true);
$objPHPExcel = $objReader->load($_FILES['file']['tmp_name']);
$rowIterator = $objPHPExcel->getActiveSheet()->getRowIterator();

$array_data = array();

//Our example Excel file has columns A through T.  Because of this, we create an array with values A -> T.  
//We then fetch all of the data from columns A through T in rows 1 to the last data row of the sheet
foreach($rowIterator as $row){
    $cellIterator = $row->getCellIterator();
    $cellIterator->setIterateOnlyExistingCells(false);
    if(1 == $row->getRowIndex()) continue;
    $rowIndex = $row->getRowIndex();
    $array_data[$rowIndex] = array("A"=>"", "B"=>"", "C"=>"", "D"=>"", "E"=>"", "F"=>"", "G"=>"", "H"=>"",
        "I"=>"", "J"=>"", "K"=>"", "L"=>"", "M"=>"", "N"=>"", "O"=>"", "P"=>"", "Q"=>"", "R"=>"", "S"=>"", "T"=>"");
    foreach($cellIterator as $cell){
        if("A" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("B" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("C" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("D" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("E" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("F" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("G" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("H" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("I" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("J" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("K" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("L" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("M" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("N" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("O" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("P" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("Q" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
        else if("R" == $cell->getColumn()){
            $array_data[$rowIndex][$cell->getColumn()] = $cell->getCalculatedValue();
        }
    }
}

//The below creates a new object from the db class in the classes.php file
$connection = new db;
$this->mysqlconnect("server address", "user name" "password");
$this->mysqlselectdb("database name");

//Grab a unique ID based on a column value in your Excel sheet and puts the unique value in column S
foreach($array_data as &$row){
    mysqlquery("SELECT your_column
				FROM your_table
				WHERE your_column = '".$row[L]."'");
    $value = mssql_fetch_array($this->results);
    $row[S] = $value["site_id"];
}

//Does a look-up for the value you want to return based on the unique ID you acquired in the previous foreach iteration
foreach($array_data as &$row){
    msqlquery("SELECT your_other_column
			   FROM your_table
			   WHERE your_id_column = '".$row[S]."'");
    $value = mssql_fetch_array($this->results);
    $row[T] = $value["id"];
}

$objPHPExcel = new PHPExcel();
$sheet = $objPHPExcel->getActiveSheet();
$sheet->setCellValue("A1", "STMS Account Number");
$sheet->setCellValue("B1", "Subscriber Last Name");
$sheet->setCellValue("C1", "Subscriber First Name");
$sheet->setCellValue("D1", "Activation Date");
$sheet->setCellValue("E1", "Disconnect Date");
$sheet->setCellValue("F1", "Account Type");
$sheet->setCellValue("G1", "Account Status");
$sheet->setCellValue("H1", "Dealer Name");
$sheet->setCellValue("I1", "Corp ID");
$sheet->setCellValue("J1", "Chain ID");
$sheet->setCellValue("K1", "Store Name");
$sheet->setCellValue("L1", "Property ID");
$sheet->setCellValue("M1", "Service Street Number");
$sheet->setCellValue("N1", "Service Street Name");
$sheet->setCellValue("O1", "Service Apt Suite Number");
$sheet->setCellValue("P1", "Service City Name");
$sheet->setCellValue("Q1", "Service State Code");
$sheet->setCellValue("R1", "Service Zip");
$sheet->setCellValue("S1", "Site ID");
$sheet->setCellValue("T1", "Address at Property?");
$sheet->getStyle("A1:T1")->getFont()->setBold(true);
$sheet->setAutoFilter("A1:T1");
$sheet->getColumnDimension("A")->setWidth(20);
$sheet->getColumnDimension("B")->setWidth(20);
$sheet->getColumnDimension("C")->setWidth(20);
$sheet->getColumnDimension("D")->setWidth(20);
$sheet->getColumnDimension("E")->setWidth(20);
$sheet->getColumnDimension("F")->setWidth(20);
$sheet->getColumnDimension("G")->setWidth(20);
$sheet->getColumnDimension("H")->setWidth(20);
$sheet->getColumnDimension("I")->setWidth(20);
$sheet->getColumnDimension("J")->setWidth(20);
$sheet->getColumnDimension("K")->setWidth(20);
$sheet->getColumnDimension("L")->setWidth(20);
$sheet->getColumnDimension("M")->setWidth(20);
$sheet->getColumnDimension("N")->setWidth(20);
$sheet->getColumnDimension("O")->setWidth(20);
$sheet->getColumnDimension("P")->setWidth(20);
$sheet->getColumnDimension("Q")->setWidth(20);
$sheet->getColumnDimension("R")->setWidth(20);
$sheet->getColumnDimension("S")->setWidth(20);
$sheet->getColumnDimension("T")->setWidth(20);


//Write all of the data into the new Excel file
foreach($array_data as $rowkey=>$rowvalue){
	foreach($rowvalue as $key=>$value){
		$sheet->setCellValue($key.$rowkey, $value);
	}
}

//Create the new file.  Name of the file will be its original name + "Revised"
$objWriter = new PHPExcel_Writer_Excel2007($objPHPExcel);
$newfilename = basename($_FILES["file"]["name"], ".xlsx")." Revised.xlsx";
$objWriter->save($newfilename);

//Header clean-up so we can download file correctly using all major browsers
header("Content-type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
header('Content-Disposition: attachment; filename = "'.$newfilename.'"');
ob_clean();
flush();
readfile($newfilename);
unlink($newfilename);
exit

?>