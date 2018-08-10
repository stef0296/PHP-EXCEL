<?php
    require 'Dependencies/vendor/autoload.php';
    use PhpOffice\PhpSpreadsheet\IOFactory;
    use PhpOffice\PhpSpreadsheet\Spreadsheet;

    $spreadsheet = new Spreadsheet();
    $header = array('Dummy Company', 'Dummy Report Header', 'Username', 'Product ID', 'Random placement header', 'dadawda');

    $connection = mysqli_connect('localhost', 'root', 'password123', 'test');

    if(!$connection){
        die("Connection failed");
    }else{
        
        $userinfo = array();
        $tableinfo = array();
        $rowAssign = 1;
        //Query to get column names
        $query = "DESC demo";
        $result= mysqli_query($connection, $query);
        if(!$result){
            die("Query Failed");
        }

        while($row=mysqli_fetch_row($result)){
            $tableinfo[] = $row[0];
        }
        
        //Query to get table data
        $query = "SELECT * FROM demo";
        $result = mysqli_query($connection, $query);
        if(!$result){
            die("Query failed");
        }
        while($row = mysqli_fetch_row($result)){
            $userinfo[] = $row;
        }
        
        //Writing Data to Excel Sheet

        //Metadata(This is the document details i.e. "Author","Date Created", etc)
        $spreadsheet->getProperties()
        ->setCreator($header[2])
        ->setLastModifiedBy($header[2])
        ->setTitle($header[1])
        ->setSubject($header[1])
        ->setDescription($header[1])
        ->setCategory($header[1]);

        //Rich Text Formatting and Cell Assignment for Headers
        for($i=0; $i<sizeof($header); $i++){
            //Rich Text Formatting and Cell Assignment for Headers
            headers($rowAssign, $header[$i]);
            $rowAssign = $rowAssign + 2;
        }

        //Writing Column Headers
        $spreadsheet->setActiveSheetIndex(0)
        ->fromArray(
            $tableinfo,
            NULL,
            'A'.$rowAssign
        );

        //Writing Table data
        $spreadsheet->setActiveSheetIndex(0)
        ->fromArray(
            $userinfo,
            NULL,
            'A'.($rowAssign+1)
        );

        //Getting Highest Row and Column
        $highestRow = $spreadsheet->getActiveSheet()
        ->getHighestRow();

        $highestColumn = $spreadsheet->getActiveSheet()
        ->getHighestColumn();
        
        //Column Headers Formatting
        $spreadsheet->getActiveSheet()->getStyle('A'.$rowAssign.':'.$highestColumn.''.$rowAssign)
        ->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FF8AB5DF');
        
        $spreadsheet->getActiveSheet()->getStyle('A'.$rowAssign.':'.$highestColumn.''.$rowAssign)
        ->getFont()
        ->setBold(true)
        ->getColor()
        ->setARGB(\PhpOffice\PhpSpreadsheet\Style\Color::COLOR_RED);

        //Cell Alignment
        $spreadsheet->getActiveSheet()
        ->getStyle('A'.$rowAssign.':'.$highestColumn.''.$highestRow)
        ->getAlignment()
        ->setHorizontal(\PhpOffice\PhpSpreadsheet\Style\Alignment::HORIZONTAL_CENTER);

        //Cell Merge
        $rowAssign = $rowAssign-2;
        while($rowAssign>0){
            cellMerging($rowAssign, $highestColumn);
            $rowAssign=$rowAssign-2;
        }
        
        for($i='A'; $i<=$highestColumn; $i++){
            $spreadsheet->getActiveSheet()->getColumnDimension($i)->setAutoSize(true);
        }

        //Renaming Worksheet
        $spreadsheet->getActiveSheet()->setTitle('Worksheet');
        $spreadsheet->getActiveSheet()->setSelectedCell('A1');

        //Header and Footer
        $spreadsheet->getActiveSheet()->getHeaderFooter()
        ->setOddHeader('&L&H'. $spreadsheet->getProperties()->getTitle() . ' &R&D');
        $spreadsheet->getActiveSheet()->getHeaderFooter()
        ->setOddFooter('&RPage &P of &N');

        // Set active sheet index to the first sheet, so Excel opens this as the first sheet
        $spreadsheet->setActiveSheetIndex(0);

        // Redirect output to a client’s web browser (Xlsx)
        header('Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet');
        header('Content-Disposition: attachment;filename="Report.xlsx"');
        header('Cache-Control: max-age=0');
        // If you're serving to IE 9, then the following may be needed
        header('Cache-Control: max-age=1');

        // If you're serving to IE over SSL, then the following may be needed
        header('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
        header('Last-Modified: ' . gmdate('D, d M Y H:i:s') . ' GMT'); // always modified
        header('Cache-Control: cache, must-revalidate'); // HTTP/1.1
        header('Pragma: public'); // HTTP/1.0

        $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
        $writer->save('php://output');
        exit;   
    }

    function headers($colNo, $docHead){
        global $spreadsheet;
        $company = new \PhpOffice\PhpSpreadsheet\RichText\RichText();
        $payable = $company->createTextRun($docHead);
        $payable->getFont()->setBold(true)
        ->setColor( new \PhpOffice\PhpSpreadsheet\Style\Color( \PhpOffice\PhpSpreadsheet\Style\Color::COLOR_BLACK ));

        $spreadsheet->getActiveSheet()->setCellValue('A'.$colNo, $company);
        $spreadsheet->getActiveSheet()->getStyle('A'.$colNo)->getFill()
        ->setFillType(\PhpOffice\PhpSpreadsheet\Style\Fill::FILL_SOLID)
        ->getStartColor()->setARGB('FFD1E8C0');
    }

    function cellMerging($rowNo, $maxColumn){
        global $spreadsheet;
        if($maxColumn=='A'||$maxColumn=='B'||$maxColumn=='C'){
            $spreadsheet->getActiveSheet()->mergeCells('A'.$rowNo.':Z'.$rowNo);
        }else{
            $spreadsheet->getActiveSheet()->mergeCells('A'.$rowNo.':'.$maxColumn.''.$rowNo);
        }
    }
?>