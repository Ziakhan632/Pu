<!DOCTYPE html>
<html lang="en">
<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Small Business & Enterprise Solutions | Comcast Business</title>

    <link rel="stylesheet" href="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/css/bootstrap.min.css">
    <link rel="stylesheet" href="https://stackpath.bootstrapcdn.com/font-awesome/4.7.0/css/font-awesome.min.css" integrity="sha384-wvfXpqpZZVQGK6TAh5PVlGOfQNHSoD2xbE+QkPxCAFlNEevoEH3Sl0sibVcOQVnN" crossorigin="anonymous">

    <!-- jQuery library -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>

    <!-- Latest compiled JavaScript -->
    <script src="https://maxcdn.bootstrapcdn.com/bootstrap/3.4.1/js/bootstrap.min.js"></script>
    <style>
        .top-nav { 
            background-color: black;
            width: 100%;
            height: 22px;
        }
        .main-nav {
            background-color: #191919;
            height: 72px;
        }
        .top-nav-list {
            list-style-type: none;
            color: white;
        }
        .list-item {
            display: inline-block;
            font-size: 11px;
            padding-right: 15px;
            color: #a5afb8;
            font-weight: 400;
            text-decoration: none;
            text-transform: uppercase;
        }
        .menu-item {
            letter-spacing: 1px;
            display: -webkit-flex;
            font-size: 1.075em;
            font-weight: 700;
            text-decoration: none;
            text-transform: uppercase;
            color: #fff;
            display: inline-block;
            margin-top: 24px;
            padding-right: 35px;
        }
        .footer-item {
            letter-spacing: 1px;
            list-style-type: none;
            font-size: 1.075em;
            font-weight: 700;
            text-decoration: none;
            text-transform: uppercase;
            color: #fff;
            margin-top: 20px;
            
            float: left;
            /* padding-right: 35px; */
            cursor: pointer;
        }

        .brand {
            cursor: pointer;
            /* font-weight: bold; */
            font-size: 20px;
        }   

        .nav-list{
            color: white;
        }
        a{
            color: white;
            text-decoration: none;
        }
        a:hover {
            text-decoration: none;
            color: gray;
        }
    </style>
</head>

<body>
<?php #include("menu.php");
#include ("menu.php");?>
<main id="main">



    <section class="inner-page">
      <div class="container" style="min-height: 600px">
   <?php

//    require 'vendor/autoload.php';
    use \PhpOffice\PhpSpreadsheet\IOFactory;

    session_start();
    if (isset($_POST["submit"])) {

        $uploads_dir = 'uploads';

        if ($_FILES["excelfile"]["error"] == UPLOAD_ERR_OK) {

            $file_ext = @strtolower(end(explode('.',$_FILES['excelfile']['name'])));
            if ($file_ext != 'xlsx') {
                $_SESSION['success_message'] = "This file type is not supported. Please, only XLSX format.";
            }
            else{
                $tmp_name = $_FILES["excelfile"]["tmp_name"];

                // может быть целесообразным дополнительно проверить имя файла
                $name = basename($_FILES["excelfile"]["name"]);

                $uploadFilePath = "$uploads_dir/$name";
                $z = move_uploaded_file($tmp_name, $uploadFilePath);

                $reader = PhpOffice\PhpSpreadsheet\IOFactory::createReader("Xlsx");

                $reader->setReadDataOnly(true);
                
               // $sheetname = "Sheet1";
                //$reader->setLoadSheetsOnly($sheetname);
                ini_set('memory_limit', '-1');
                $spreadsheet = $reader->load('./'.$uploadFilePath);
                $sheetCount = $spreadsheet->getSheetCount();

$dataRow = array();
for ($i = 0; $i < $sheetCount; $i++) {
               
                //$worksheet = $spreadsheet->getActiveSheet();
                $worksheet = $spreadsheet->getSheet($i);
                $rowIndex = 0;

                foreach ($worksheet->getRowIterator() as $row) {

                    $cellIterator = $row->getCellIterator();
                    $cellIterator->setIterateOnlyExistingCells(false);

                    $cell_value_first_column = '';
                    foreach ($cellIterator as $cell) {
   
                        $cell_value = $cell->getValue();
                        if (($col = $cell->getColumn()) == "A"){
                            $cell_value_first_column = $cell_value;
                        }


                        if ($cell->getColumn() != "A") {

                            switch ($cell_value_first_column){
                                case 'Title':
                                    $dataRow['Title'][] = $cell_value;
                                    break;
                                case 'Start Date':
                                    $dataRow['Start_Date'][] = $cell_value;
                                    break;
                                case 'End Date':
                                    $dataRow['End_Date'][] = $cell_value;
                                    break;
                                case 'Division':
                                    $dataRow['Division'][] = $cell_value;
                                    break;
                                case 'Region':
                                    $dataRow['Region'][] = $cell_value;
                                    break;
                                case 'Contract Term Required':
                                    $dataRow['Contract_Term_Required'][] = $cell_value;
                                    break;
                                case 'Customer Order Type':
                                    $dataRow['Customer_Order_Type'][] = $cell_value;
                                    break;
                                case 'Is AutoPay/Paperless Billing Eligible?':
                                    $dataRow['Is_AutoPay_Paperless_Billing_Eligible'][] = $cell_value;
                                    break;
                                case 'Value Add':
                                    $dataRow['Value_Add'][] = $cell_value;
                                    break;
                                case 'Package Name':
                                    $dataRow['Package_Name'][] = $cell_value;
                                    break;
                                case 'Common Code':
                                    $dataRow['Common_Code'][] = $cell_value;
                                    break;
                                case 'Package Product IDs':
                                    $dataRow['Package_Product_IDs'][] = $cell_value;
                                    break;
                                case 'Total Package Months 1-12 Price':
                                    $dataRow['Total_Package_Months_1_12_Price'][] = $cell_value;
                                    break;
                                case 'Total Package Months 13-24 Price':
                                    $dataRow['Total_Package_Months_13_24_Price'][] = $cell_value;
                                    break;
                                case 'Total Package Months 25-36 Price':
                                    $dataRow['Total_Package_Months_25_36_Price'][] = $cell_value;
                                    break;
                                case 'Total Package EDP retail rate':
                                    $dataRow['Total_Package_EDP_retail_rate'][] = $cell_value;
                                    break;
                                case 'Business Internet Tier':
                                    $dataRow['Business_Internet_Tier'][] = $cell_value;
                                    break;
                                case 'BI Rate Card':
                                    $dataRow['BI_Rate_Card'][] = $cell_value;
                                    break;
                                case 'BI Months 1-12 Promo':
                                    $dataRow['BI_Months_1_12_Promo'][] = $cell_value;
                                    break;
                                case 'Start BI Months 13-24 Promo':
                                    $dataRow['BI_Months_13_24_Promo'][] = $cell_value;
                                    break;
                                case 'BI Months 25-36 Promo':
                                    $dataRow['BI_Months_25_36_Promo'][] = $cell_value;
                                    break;
                                case 'BI Equipment':
                                    $dataRow['BI_Equipment'][] = $cell_value;
                                    break;
                                case 'Static IP Options':
                                    $dataRow['Static_IP_Options'][] = $cell_value;
                                    break;
                                case 'Wifi Options':
                                    $dataRow['Wifi_Options'][] = $cell_value;
                                    break;
                                case 'Connection Pro Options':
                                    $dataRow['Connection_Pro_Options'][] = $cell_value;
                                    break;
                                case 'SecurityEdge Options':
                                    $dataRow['SecurityEdge_Options'][] = $cell_value;
                                    break;
                                case 'Required Line 1 Type':
                                    $dataRow['Required_Line_1_Type'][] = $cell_value;
                                    break;
                                case 'Required Line 1 Type: Rate Card':
                                    $dataRow['Required_Line_1_Type_Rate_Card'][] = $cell_value;
                                    break;
                                case 'Required Line 1 Type: Months 1-12 Promo':
                                    $dataRow['Required_Line_1_Type_Months_1_12_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 1 Type: Months 13-24 Promo':
                                    $dataRow['Required_Line_1_Type_Months_13_24_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 1 Type: Months 25-36 Promo':
                                    $dataRow['Required_Line_1_Type_Months_25_36_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 2 Type':
                                    $dataRow['Required_Line_2_Type'][] = $cell_value;
                                    break;
                                case 'Required Line 2 Type: Rate Card':
                                    $dataRow['Required_Line_2_Type_Rate_Card'][] = $cell_value;
                                    break;
                                case 'Required Line 2 Type: Months 1-12 Promo':
                                    $dataRow['Required_Line_2_Type_Months_1_12_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 2 Type: Months 13-24 Promo':
                                    $dataRow['Required_Line_2_Type_Months_13_24_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 2 Type: Months 25-36 Promo':
                                    $dataRow['Required_Line_2_Type_Months_25_36_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 3 Type':
                                    $dataRow['Required_Line_3_Type'][] = $cell_value;
                                    break;
                                case 'Required Line 3 Type: Rate Card':
                                    $dataRow['Required_Line_3_Type_Rate_Card'][] = $cell_value;
                                    break;
                                case 'Required Line 3 Type: Months 1-12 Promo':
                                    $dataRow['Required_Line_3_Type_Months_1_12_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 3 Type: Title':
                                    $dataRow['Required_Line_3_Type_Title'][] = $cell_value;
                                    break;
                                case 'Required Line 3 Type: Months 13-24 Promo':
                                    $dataRow['Required_Line_3_Type_Months_13_24_Promo'][] = $cell_value;
                                    break;
                                case 'Required Line 3 Type: Months 25-36 Promo':
                                    $dataRow['Required_Line_3_Type_Months_25_36_Promo'][] = $cell_value;
                                    break;
                                case 'Optional Line Type':
                                    $dataRow['Optional_Line_Type'][] = $cell_value;
                                    break;
                                case 'Optional Line Type: Rate Card':
                                    $dataRow['Optional_Line_Type_Rate_Card'][] = $cell_value;
                                    break;
                                case 'Optional Line Type: Months 1-12 Promo':
                                    $dataRow['Optional_Line_Type_Months_1_12_Promo'][] = $cell_value;
                                    break;
                                case 'Optional Line Type: Months 13-24 Promo':
                                    $dataRow['Optional_Line_Type_Months_13_24_Promo'][] = $cell_value;
                                    break;
                                case 'Optional Line Type: Months 25-36 Promo':
                                    $dataRow['Optional_Line_Type_Months_25_36_Promo'][] = $cell_value;
                                    break;
                                case 'Optional Line Type: BV Equipment':
                                    $dataRow['Optional_Line_Type_BV_Equipment'][] = $cell_value;
                                    break;
                                case 'Optional Line Type: BV Bolt Ons':
                                    $dataRow['Optional_Line_Type_BV_Bolt_Ons'][] = $cell_value;
                                    break;
                                case 'Business TV Tier':
                                    $dataRow['Business_TV_Tier'][] = $cell_value;
                                    break;
                                case 'Business TV Rate Card':
                                    $dataRow['Business_TV_Rate_Card'][] = $cell_value;
                                    break;
                                case 'Business TV Months 1-12 Promo':
                                    $dataRow['Business_TV_Months_1_12_Promo'][] = $cell_value;
                                    break;
                                case 'Business TV Months 13-24 Promo':
                                    $dataRow['Business_TV_Months_13_24_Promo'][] = $cell_value;
                                    break;
                                case 'Business TV Months 25-36 Promo':
                                    $dataRow['Business_TV_Months_25_36_Promo'][] = $cell_value;
                                    break;
                                case 'Business TV Equipment':
                                    $dataRow['Business_TV_Equipment'][] = $cell_value;
                                    break;
                                case 'Install Fee':
                                    $dataRow['Install_Fee'][] = $cell_value;
                                    break;
                                case 'eCom Promo Code':
                                    $dataRow['eCom_Promo_Code'][] = $cell_value;
                                    break;
                                case 'Promo Description':
                                    $dataRow['Promo_Description'][] = $cell_value;
									$dataRow['Details_and_Restrictions'][] = $cell_value;
                                    break;
                                case "BB Order Entry Package Code":
                                    $dataRow['BB_Order_Entry_Package_Code'][] = $cell_value;
                                    break;
                                case 'BB Order Entry Promo Code':
                                    $dataRow['BB_Order_Entry_Promo_Code'][] = $cell_value;
                                    break;
									
                            }
                        }
                    }
                      
                }
             
            }

                /* database connection */
                $host = "localhost";
                $user = "root";
                $pass = "comcast";
                //$pass = '';
				$db_name = "crawl_summary";
                $connection = mysqli_connect($host, $user, $pass, $db_name);
                /* database connection */

                if (mysqli_connect_errno()) {
                    die("connection failed: "
                        . mysqli_connect_error()
                        . " (" . mysqli_connect_errno()
                        . ")");
                }

                $sqlcreate = "CREATE TABLE IF NOT EXISTS crawl_summary.offers_from_excel (Offer_Id INT PRIMARY KEY AUTO_INCREMENT, Title VARCHAR(255), Start_Date VARCHAR(255), End_Date VARCHAR(255), Division VARCHAR(255), Region VARCHAR(255), Contract_Term_Required VARCHAR(255), Customer_Order_Type VARCHAR(255), Is_AutoPay_Paperless_Billing_Eligible VARCHAR(255), Value_Add VARCHAR(255), Package_Name VARCHAR(255), Common_Code VARCHAR(255), Package_Product_IDs VARCHAR(255), Total_Package_Months_1_12_Price VARCHAR(255), Total_Package_Months_13_24_Price VARCHAR(255), Total_Package_Months_25_36_Price VARCHAR(255), Total_Package_EDP_retail_rate VARCHAR(255), Business_Internet_Tier VARCHAR(255), BI_Rate_Card VARCHAR(255), BI_Months_1_12_Promo VARCHAR(255), BI_Months_13_24_Promo VARCHAR(255), BI_Months_25_36_Promo VARCHAR(255), BI_Equipment VARCHAR(255), Static_IP_Options VARCHAR(255), Wifi_Options VARCHAR(255), Connection_Pro_Options VARCHAR(255), SecurityEdge_Options VARCHAR(255), Required_Line_1_Type VARCHAR(255), Required_Line_1_Type_Rate_Card VARCHAR(255), Required_Line_1_Type_Months_1_12_Promo VARCHAR(255), Required_Line_1_Type_Months_13_24_Promo VARCHAR(255), Required_Line_1_Type_Months_25_36_Promo VARCHAR(255), Required_Line_2_Type VARCHAR(255), Required_Line_2_Type_Rate_Card VARCHAR(255), Required_Line_2_Type_Months_1_12_Promo VARCHAR(255), Required_Line_2_Type_Months_13_24_Promo VARCHAR(255), Required_Line_2_Type_Months_25_36_Promo VARCHAR(255), Required_Line_3_Type VARCHAR(255), Required_Line_3_Type_Rate_Card VARCHAR(255), Required_Line_3_Type_Months_1_12_Promo VARCHAR(255), Required_Line_3_Type_Title VARCHAR(255), Required_Line_3_Type_Months_13_24_Promo VARCHAR(255), Required_Line_3_Type_Months_25_36_Promo VARCHAR(255), Optional_Line_Type VARCHAR(255), Optional_Line_Type_Rate_Card VARCHAR(255), Optional_Line_Type_Months_1_12_Promo VARCHAR(255), Optional_Line_Type_Months_13_24_Promo VARCHAR(255), Optional_Line_Type_Months_25_36_Promo VARCHAR(255), Optional_Line_Type_BV_Equipment VARCHAR(255), Optional_Line_Type_BV_Bolt_Ons VARCHAR(255), Business_TV_Tier VARCHAR(255), Business_TV_Rate_Card VARCHAR(255), Business_TV_Months_1_12_Promo VARCHAR(255), Business_TV_Months_13_24_Promo VARCHAR(255), Business_TV_Months_25_36_Promo VARCHAR(255), Business_TV_Equipment VARCHAR(255), Install_Fee VARCHAR(255), eCom_Promo_Code VARCHAR(255), Promo_Description TEXT, Details_and_Restrictions TEXT, BB_Order_Entry_Package_Code VARCHAR(255), BB_Order_Entry_Promo_Code VARCHAR(255));";

                if (!$connection->query($sqlcreate) === TRUE) {
                    echo "Error: " . $sqlcreate . "<br>" . $connection->error;
                }

//                $sql = "TRUNCATE TABLE crawl_summary.offers_from_excel;";
//                if (!$connection->query($sql) === TRUE) {
//                    echo "Error: " . $sql . "<br>" . $connection->error;
//                }

                $offersData = array();

                foreach($dataRow as $key => $value)
                {
                    $i = 0;
                    foreach ($value as $val)
                    {
                        if ($val == null) continue;
                        $offersData[$i][$key] = $val;
                        $i++;
                    }

                }
                if(count($offersData)){
                    foreach ($offersData as $offer){

                    $sql = "INSERT INTO crawl_summary.offers_from_excel (";
                    $values = "VALUES (";

                    foreach($offer as $key => $val){
                        if ($val == null) continue;
                        $sql .= $key.', ';
                        $escapestr = $connection->real_escape_string(htmlspecialchars($val, ENT_QUOTES));
                        $values .= "'$escapestr'".', ';
                    }
                    $sql = substr($sql, 0, strlen($sql) - 2).') ';
                    $values = substr($values, 0, strlen($values) - 2).');';

                    $sql .= $values;
                    if (!$connection->query($sql) === TRUE) {
                        echo "Error: " . $sql . "<br>" . $connection->error;
                    }
                }
                $_SESSION['success_message'] = "File uploaded successfully.";
                }else{
                     $_SESSION['error_message'] = "Something went wrong,File Data is not Uploaded,Please try again.";
                }
                
               
            }

        }

    }




   ?>
   </br>
   <form method="POST" action="upload_file_new.php" enctype="multipart/form-data">
    <div class="upload-wrapper">
      <h4>Upload file to database</h4></br>
	  <?php
        if(isset($_SESSION['success_message']) && $_SESSION['success_message'] != '')
		{
			echo '<p style="color:green">'.$_SESSION['success_message'].'</p>';
			//session_destroy();
		}			
	  ?>
      <?php
        if(isset($_SESSION['error_message']) && $_SESSION['error_message'] != '')
        {
            echo '<p style="color:red">'.$_SESSION['error_message'].'</p>';
            //session_destroy();
        }           
      ?>
      <label for="file-upload">Choose File<span>*</span> </br><input type="file" id="file-upload" name="excelfile" required></label>
    </div>
 
    <input type="submit" class='btn-primary' name="submit" value="Submit" />
  </form>
      </div>
    </section>

  </main><!-- End #main -->
  

    <div class="col-lg-12" style="background-color: black; width: 100%; height: 600px; color: white;">
        <div class="col-lg-1">

        </div>
        <div class="col-lg-10" style="height: 80%;">
            <div class="col-lg-2" style="height: 100%; padding-top: 60px;">
                <img src="./logo_black_bg.jpg"/>
            </div>

            <div class="col-lg-2" style="height: 100%;padding-top: 72px;">
                <h4 style="text-align: center;font-weight: 700;"> 
                    Business
                    <hr >
                </h4>
                <ul>
                    <li class="footer-item" >
                        <a href="business.php">
                            Offers
                        </a>
                    </li>
                    <li class="footer-item">
                        
                        <a href="business.php">
                        Configure
                        </a>
                    </li>
                    <li class="footer-item" >
                        
                        <a href="business.php">
                        Checkout
                        </a>
                    </li>
                </ul>
                
            </div>

            <div class="col-lg-2" style="height: 100%;padding-top: 72px;">
                <h4 style="text-align: center;font-weight: 700;"> 
                    Preview
                    <hr>
                </h4>
                <ul>
                    <li class="footer-item" >
                        
                        <a href="preview.php">
                        Main Page
                        </a>
                    </li>
                    <li class="footer-item">
                        
                        <a href="preview.php">
                            Second Page
                        </a>
                    </li>
                </ul>


            </div>

            <div class="col-lg-2" style="height: 100%;padding-top: 72px;">
                <h4 style="text-align: center;font-weight: 700;"> 
                    Healthcheck
                    <hr>
                </h4>
                <ul>
                    <li class="footer-item" >
                        
                        <a href="healthcheck.php">
                            Business Page
                        </a>
                    </li>
                    <li class="footer-item">
                        <a href="healthcheck.php">
                        
                        Preview Page
                        </a>
                    </li>
                </ul>
            </div>

            <div class="col-lg-2" style="height: 100%;padding-top: 72px;">

            </div>
            
            <div class="col-lg-2" style="height: 100%;padding-top: 72px;">

            </div>

        </div>
        <div class="col-lg-1">

        </div>
        <hr width="83%">

        <span style="margin-left: 550px; font-weight: 700;font-size: 15px;">
            ©2020 Comcast Corporation
        </span>
    </div>
    
<script>




$(document).ready(function() {
    $('#btnOutlook').click(function() {
        var x = "http://hqswl-c051213:8080/comcast/index.php";
        
        var today = new Date();
        var dd = String(today.getDate()).padStart(2, '0');
        var mm = String(today.getMonth() + 1).padStart(2, '0'); //January is 0!
        // var yyyy = today.getFullYear();
        
        var b  = "Offers Section by Address and Date "+ mm + '/'+dd;
        x = "Here is the link: "+ x;

        window.open('mailto:test@example.com?subject='+b+'&body='+x);
    });
});
    </script>

</body>
</html>
