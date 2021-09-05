<?php

	/*-------------------------
	Authors:
    JOEL DAVID DE LOS RÍOS GOEZ
    DEISY ALEJANDRA CORREA VANEGAS
	FERNEY LONDOÑO QUINTERO
	JOHAN SEBASTIÁN HOLGUÍN CANO

    ---------------------------*/
    if (PHP_SAPI == 'cli')
	    die('Este ejemplo sólo se puede ejecutar desde un navegador Web');

    require_once dirname(__FILE__) . '/classes/PHPExcel.php';   /** Incluye PHPExcel */
    require_once ("config/db.php");//Contiene las variables de configuracion para conectar a la base de datos
    require_once ("config/conexion.php");//Contiene funcion que conecta a la base de datos

    // Crear nuevo objeto PHPExcel
    $objPHPExcel = new PHPExcel();

    // Propiedades del documento
        $objPHPExcel->getProperties()->setCreator("CarPoint");
        $objPHPExcel->getProperties()->setDescription("Reporte Clientes para Office 2010 XLSX,
         generado usando clases de PHP.")
        ;


	// Establecer índice de hoja activa a la primera hoja , por lo que Excel abre esto como la primera hoja
    $objPHPExcel->setActiveSheetIndex(0);

    // Cambiar el nombre de hoja de cálculo
    $objPHPExcel->getActiveSheet()->setTitle("Reporte");

    // Combino las celdas desde A1 hasta G1
    $objPHPExcel->getActiveSheet()->mergeCells('A1:G1');

    //Títulos de las celdas
    $objPHPExcel->getActiveSheet()->setCellValue('A1', 'REPORTE CLIENTES');
    $objPHPExcel->getActiveSheet()->setCellValue('A2', 'ID_Cliente');
    $objPHPExcel->getActiveSheet()->setCellValue('B2', 'Nombre');
    $objPHPExcel->getActiveSheet()->setCellValue('C2', 'Telefono');
    $objPHPExcel->getActiveSheet()->setCellValue('D2', 'Email');
	$objPHPExcel->getActiveSheet()->setCellValue('E2', 'Direccion');
    $objPHPExcel->getActiveSheet()->setCellValue('F2', 'Estado');
    $objPHPExcel->getActiveSheet()->setCellValue('G2', 'Agregado');

    // Fuente de la primera fila en negrita
    $boldArray = array('font' => array('bold' => true,),'alignment' => array('horizontal' => PHPExcel_Style_Alignment::HORIZONTAL_CENTER));
    $objPHPExcel->getActiveSheet()->getStyle('A1:G2')->applyFromArray($boldArray);	

//Ancho de las columnas
$objPHPExcel->getActiveSheet()->getColumnDimension('A')->setWidth(12);	
$objPHPExcel->getActiveSheet()->getColumnDimension('B')->setWidth(15);	
$objPHPExcel->getActiveSheet()->getColumnDimension('C')->setWidth(11);	
$objPHPExcel->getActiveSheet()->getColumnDimension('D')->setWidth(30);	
$objPHPExcel->getActiveSheet()->getColumnDimension('E')->setWidth(15);	
$objPHPExcel->getActiveSheet()->getColumnDimension('F')->setWidth(9);	
$objPHPExcel->getActiveSheet()->getColumnDimension('G')->setWidth(15);	


/*Extraer datos de MYSQL*/

    $sql="SELECT * FROM clientes";
    $query=mysqli_query($con,$sql);
    $cel=3;//Numero de fila donde empezara a crear  el reporte
	while ($row=mysqli_fetch_array($query)){
		$id_cliente=$row['id_cliente'];
		$nombre_cliente=$row['nombre_cliente'];
		$telefono_cliente=$row['telefono_cliente'];
		$email_cliente=$row['email_cliente'];
        $direccion_cliente=$row['direccion_cliente'];
        $status_cliente=$row['status_cliente'];
        $date_added=$row['date_added'];
		
			$a="A".$cel;
			$b="B".$cel;
			$c="C".$cel;
			$d="D".$cel;
            $e="E".$cel;
            $f="F".$cel;
            $g="G".$cel;
            

			// Agregar datos
			$objPHPExcel->getActiveSheet()->setCellValue($a, $id_cliente);
            $objPHPExcel->getActiveSheet()->setCellValue($b, $nombre_cliente);
            $objPHPExcel->getActiveSheet()->setCellValue($c, $telefono_cliente);
            $objPHPExcel->getActiveSheet()->setCellValue($d, $email_cliente);
            $objPHPExcel->getActiveSheet()->setCellValue($e, $direccion_cliente);
            $objPHPExcel->getActiveSheet()->setCellValue($f, $status_cliente);
            $objPHPExcel->getActiveSheet()->setCellValue($g, $date_added);
			
	$cel++;
    }
    
    /*Fin extracion de datos MYSQL*/

    
    $rango="A2:$g";
    $styleArray = array('font' => array( 'name' => 'Arial','size' => 10),
    'borders'=>array('allborders'=>array('style'=> PHPExcel_Style_Border::BORDER_THIN,'color'=>array('argb' => 'FFF')))
    );
    $objPHPExcel->getActiveSheet()->getStyle($rango)->applyFromArray($styleArray);
    


    // Redirigir la salida al navegador web de un cliente ( Excel2007 )
	header("Content-Type: application/vnd.openxmlformats-officedocument.spreadsheetml.sheet");
	header('Content-Disposition: attachment;filename=reporte.xlsx');
    header('Cache-Control: max-age=0');
 
    // Si usted está sirviendo a IE 9 , a continuación, puede ser necesaria la siguiente
    header('Cache-Control: max-age=1');
    
    // Si usted está sirviendo a IE a través de SSL , a continuación, puede ser necesaria la siguiente
    header ('Expires: Mon, 26 Jul 1997 05:00:00 GMT'); // Date in the past
    header ('Last-Modified: '.gmdate('D, d M Y H:i:s').' GMT'); // always modified
    header ('Cache-Control: cache, must-revalidate'); // HTTP/1.1
    header ('Pragma: public'); // HTTP/1.0
   

    $objWriter = PHPExcel_IOFactory::createWriter($objPHPExcel, 'Excel2007');
	$objWriter->save('php://output'); 

?>