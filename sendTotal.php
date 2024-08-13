<?php
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require 'vendor/autoload.php';

//$trama = "1,8";
//genset
$trama = "2,4393";
//madurador
//$trama ="2,257"
$cadenaTrama = explode(",",$trama);
$mail = new PHPMailer(true);
$mensaje = "";
require_once 'api_correo.php';
$api = new ApiModel();
//echo $cadenaTrama[0]."<br>";
//echo $cadenaTrama[1];
$fechaActual = date("Y-m-d");
$fechaExcel =  date("Y-m-d",strtotime($fechaActual."- 1 days"));
$fechaExcel1 = strtotime($fechaActual."- 1 days");
$fechaInicio =date("Y-m-d",strtotime($fechaActual."- 1 days"))." 00:00:00";
$fechaFin =date("Y-m-d H:i:s");
$correoEnvio = "devpablito2023@gmail.com";
$telemetria_id =$cadenaTrama[1];

$cadenaTiempo =explode("-",$fechaExcel);
$fechaValida =$cadenaTiempo[2]."/".$cadenaTiempo[1]."/".$cadenaTiempo[0];

switch ($cadenaTrama[0]) {
    case 1:
        //Estamos en el tipos de reefer
        $R = $api->reefer($telemetria_id);
        $Completo = $api->listaReeferFecha($telemetria_id ,$fechaInicio , $fechaFin);
        $reporte = $api->reporteReefer($telemetria_id ,$fechaInicio , $fechaFin);
        $empresa =$api->empresa($R['empresa_id']);
        $usuario_empresa = $api->usuario_id($R['empresa_id']);   
        $usuario_seleccionado = $api->usuario_enviar($usuario_empresa['usuario_id']);
        //creacion de Excel para adjuntar a correo
        $tituloExcel =" Reporte de rutina para Reefer";
        $documento = new Spreadsheet();
        $documento
        ->getProperties()
        ->setCreator("Luis Pablo Marcelo Perea")
        ->setLastModifiedBy('Helios Tec')
        ->setTitle($tituloExcel)
        ->setDescription('Detalles del comportamiento del reefer');
        $hojaDeProductos = $documento->getActiveSheet();
        $hojaDeProductos->setTitle("Reefer");
        # Encabezado de los productos
        $nombreContenedor = $R['nombre_contenedor'];
        $dispositivo = ["Reefer :" ,$nombreContenedor];
        $encabezado = ["Reception Date", "Set Point", "Temp Supply", "Return Air", "Evaporation Coil","Ambient Air","Relative Humidity","Alarm Present","Alarm Number","Controlling Mode","Power State","Defrost Term Temp","Defrost Interval","Latitude","Length"];
        # El último argumento es por defecto A1
        $hojaDeProductos->fromArray($dispositivo, null, 'A1');
        $hojaDeProductos->fromArray($encabezado, null, 'A2');
        # Comenzamos en la fila 3
        $numeroDeFila = 3;
        foreach($reporte as $fila){
            # Escribir registros en el documento
            $hojaDeProductos->setCellValueByColumnAndRow(1, $numeroDeFila, $fila['created_at']);
            $hojaDeProductos->setCellValueByColumnAndRow(2, $numeroDeFila, $fila['set_point']);
            $hojaDeProductos->setCellValueByColumnAndRow(3, $numeroDeFila, $fila['temp_supply']);
            $hojaDeProductos->setCellValueByColumnAndRow(4, $numeroDeFila, $fila['return_air']);
            $hojaDeProductos->setCellValueByColumnAndRow(5, $numeroDeFila, $fila['evaporation_coil']);
            $hojaDeProductos->setCellValueByColumnAndRow(6, $numeroDeFila, $fila['ambient_air']);
            $hojaDeProductos->setCellValueByColumnAndRow(7, $numeroDeFila, $fila['relative_humidity']);
            $hojaDeProductos->setCellValueByColumnAndRow(8, $numeroDeFila, $fila['alarm_present']);
            $hojaDeProductos->setCellValueByColumnAndRow(9, $numeroDeFila, $fila['alarm_number']);
            $hojaDeProductos->setCellValueByColumnAndRow(10, $numeroDeFila, $fila['controlling_mode']);
            $hojaDeProductos->setCellValueByColumnAndRow(11, $numeroDeFila, $fila['power_state']);
            $hojaDeProductos->setCellValueByColumnAndRow(12, $numeroDeFila, $fila['defrost_term_temp']);
            $hojaDeProductos->setCellValueByColumnAndRow(13, $numeroDeFila, $fila['defrost_interval']);
            $hojaDeProductos->setCellValueByColumnAndRow(14, $numeroDeFila, $fila['latitud']);
            $hojaDeProductos->setCellValueByColumnAndRow(15, $numeroDeFila, $fila['longitud']);
            $numeroDeFila++;
        }
        # Crear un "escritor"
        $writer = new Xlsx($documento);
        # Le pasamos la ruta de guardado
        $writer->save('./excel/'.$nombreContenedor.'_'.$fechaExcel.'.xlsx');
        $mensaje .="
             <head>
                <style>
                    table {
                        border-collapse: collapse;
                        margin: 25px 0; 
                        font-size: 1em;
                        font-family: sans-serif;
                        min-width: 1450px; 
                        box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
                    }
                    thead tr {
                        background-color: #1a2c4e;
                        color: #ffffff;
                        text-align: middle;
                    }
                    th, td {
                        padding: 12px 15px;
                    }      
                    tbody tr{
                        border-bottom: 1px solid #1a2c4e;
                    }      
                </style>   
            </head>
        ";
        //Aspectos para el envio de correo
        $asunto = "Reporte de rutina , Reefer: ".$R['nombre_contenedor']."  dia : ".$fechaValida;
        $mensaje .= "<h2> Dear : ".$usuario_seleccionado['nombres']." ".$usuario_seleccionado['apellidos']."</h2>";
        $mensaje .= "<h2> Reefer : ".$R['nombre_contenedor']."</h2>";
        $mensaje .= "<h2> Message : "." Reporte de Rutina "."</h2>";
        $mensaje .="<body><table  style='border:1px solid #1a2c4e' ><thead><tr ><th width='130'>Reception Date</th><th width='60'>Set Point </th><th > Temp Supply </th><th>Return Air </th><th>Evaporation Coil </th><th>Ambient Air </th><th>Relative Humidity</th><th>Alarm Present</th><th>";
        $mensaje .="Alarm Number </th><th>Controlling Mode</th><th> Power State </th><th> Defrost Term Temp </th><th>Defrost Interval</th><th>Latitude </th><th>Length </th></tr></thead><tbody>";
        foreach($Completo as $fila){
            $mensaje .="<tr align='center' ><td width='130'><strong>".$fila['created_at']."</strong></td><td>".$fila['set_point']."</td></td>".$fila['temp_supply']."</td><td>".$fila['return_air']."</td></td>".$fila['evaporation_coil']."</td><td>".$fila['ambient_air']."</td></td>".$fila['relative_humidity']." %</td><td>".$fila['alarm_present'];
            $mensaje .="</td></td>".$fila['alarm_number']."</td><td>".$fila['controlling_mode']."</td></td>".$fila['power_state']."</td><td>".$fila['defrost_term_temp']."</td></td>".$fila['defrost_interval']."</td><td>".$fila['latitud']."</td></td>".$fila['longitud']."</td></tr>";
        }
        $mensaje .="</tbody></table></body>";
        try {
            //Server settings
            $mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
            $mail->isSMTP();    
            //$mail->From = "ztrack@zgroup.com.pe"; 
            $mail->From = "desarrollo@zgroup.com.pe";                                   //Send using SMTP
            $mail->Host       = "smtp.gmail.com";                   //Set the SMTP server to send through
            $mail->SMTPAuth   = true;                                   //Enable SMTP authentication
            $mail->Username   = 'desarrollo@zgroup.com.pe';                     //SMTP username
            $mail->Password   = 'Des5090100';                               //SMTP password
            //$mail->Username   = 'ztrack@zgroup.com.pe';                     //SMTP username
            //$mail->Password   = 'Proyectoztrack2023!';                               //SMTP password
            $mail->SMTPSecure = 'tls';            //Enable implicit TLS encryption
            $mail->Port       = 587;                                    //TCP port to connect to; use 587 if you have set `SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS`
            //Agregar destinatario
            $mail->AddAddress($correoEnvio);
            $mail->Subject = $asunto;
            $mail->Body =$mensaje;
            $mail->isHTML(true);
            $mail->AddAttachment('./excel/'.$nombreContenedor.'_'.$fechaExcel.'.xlsx', $nombreContenedor.'_'.$fechaExcel.'.xlsx');
            //Avisar si fue enviado o no y dirigir al index
            if ($mail->Send()) {
                echo'<script type="text/javascript">alert("Enviado Correctamente");</script>';  
            } else {
               echo'<script type="text/javascript">alert("NO ENVIADO, intentar de nuevo");</script>';
            }    
        }catch (Exception $e) {
            echo "Se ha producido un mensaje de error . Mailer Error: {$mail->ErrorInfo}"; 
        }
    break;
    case 3:
        //Estamos en el tipos de Genset
        $R = $api->generador($telemetria_id);
        $Completo = $api->listaGeneradorFecha($telemetria_id ,$fechaInicio , $fechaFin);
        $reporte = $api->reporteGenerador($telemetria_id ,$fechaInicio , $fechaFin);
        $empresa =$api->empresa($R['empresa_id']);
        $usuario_empresa = $api->usuario_id($R['empresa_id']);   
        $usuario_seleccionado = $api->usuario_enviar($usuario_empresa['usuario_id']);
        //creacion de Excel para adjuntar a correo
        $tituloExcel =" Reporte de rutina para Genset";
        $documento = new Spreadsheet();
        $documento
        ->getProperties()
        ->setCreator("Luis Pablo Marcelo Perea")
        ->setLastModifiedBy('Helios Tec')
        ->setTitle($tituloExcel)
        ->setDescription('Detalles del comportamiento del Genset');
        $hojaDeProductos = $documento->getActiveSheet();
        $hojaDeProductos->setTitle("Genset");
        # Encabezado de los productos
        $nombreContenedor = $R['nombre_generador'];
        $dispositivo = ["Genset :" ,$nombreContenedor];
        $encabezado = ["Reception Date", "Battery Voltage", "Water Temp", "Running Frequency", "Fuel Level","Voltage Measure","Rotor Current","Fiel Current","Speed","Eco Power","RPM","Unit Mode","Horometro","Model","Latitude","Length","Alarm","Event","Reefer Conected","Set Point","Temp Supply","Return Air"];
        # El último argumento es por defecto A1
        $hojaDeProductos->fromArray($dispositivo, null, 'A1');
        $hojaDeProductos->fromArray($encabezado, null, 'A2');
        # Comenzamos en la fila 3
        $numeroDeFila = 3;
        foreach($reporte as $fila){
            # Escribir registros en el documento
            $hojaDeProductos->setCellValueByColumnAndRow(1, $numeroDeFila, $fila['created_at']);
            $hojaDeProductos->setCellValueByColumnAndRow(2, $numeroDeFila, $fila['battery_voltage']);
            $hojaDeProductos->setCellValueByColumnAndRow(3, $numeroDeFila, $fila['water_temp']);
            $hojaDeProductos->setCellValueByColumnAndRow(4, $numeroDeFila, $fila['running_frequency']);
            $hojaDeProductos->setCellValueByColumnAndRow(5, $numeroDeFila, $fila['fuel_level']);
            $hojaDeProductos->setCellValueByColumnAndRow(6, $numeroDeFila, $fila['voltage_measure']);
            $hojaDeProductos->setCellValueByColumnAndRow(7, $numeroDeFila, $fila['rotor_current']);
            $hojaDeProductos->setCellValueByColumnAndRow(8, $numeroDeFila, $fila['fiel_current']);
            $hojaDeProductos->setCellValueByColumnAndRow(9, $numeroDeFila, $fila['speed']);
            $hojaDeProductos->setCellValueByColumnAndRow(10, $numeroDeFila, $fila['eco_power']);
            $hojaDeProductos->setCellValueByColumnAndRow(11, $numeroDeFila, $fila['rpm']);
            $hojaDeProductos->setCellValueByColumnAndRow(12, $numeroDeFila, $fila['unit_mode']);
            $hojaDeProductos->setCellValueByColumnAndRow(13, $numeroDeFila, $fila['horometro']);
            $hojaDeProductos->setCellValueByColumnAndRow(14, $numeroDeFila, $fila['modelo']);
            $hojaDeProductos->setCellValueByColumnAndRow(15, $numeroDeFila, $fila['latitud']);
            $hojaDeProductos->setCellValueByColumnAndRow(16, $numeroDeFila, $fila['longitud']);
            $hojaDeProductos->setCellValueByColumnAndRow(17, $numeroDeFila, $fila['alarma_id']);
            $hojaDeProductos->setCellValueByColumnAndRow(18, $numeroDeFila, $fila['evento_id']);
            $hojaDeProductos->setCellValueByColumnAndRow(19, $numeroDeFila, $fila['reefer_conected']);
            $hojaDeProductos->setCellValueByColumnAndRow(20, $numeroDeFila, $fila['set_point']);
            $hojaDeProductos->setCellValueByColumnAndRow(21, $numeroDeFila, $fila['temp_supply']);
            $hojaDeProductos->setCellValueByColumnAndRow(22, $numeroDeFila, $fila['return_air']);
            $numeroDeFila++;
        }
        # Crear un "escritor"
        $writer = new Xlsx($documento);
        # Le pasamos la ruta de guardado
        $writer->save('./excel/'.$nombreContenedor.'_'.$fechaExcel.'.xlsx');
        $mensaje .="
             <head>
                <style>
                    table {
                        border-collapse: collapse;
                        margin: 25px 0; 
                        font-size: 1em;
                        font-family: sans-serif;
                        min-width: 1450px; 
                        box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
                    }
                    thead tr {
                        background-color: #1a2c4e;
                        color: #ffffff;
                        text-align: middle;
                    }
                    th, td {
                        padding: 12px 15px;
                    }      
                    tbody tr{
                        border-bottom: 1px solid #1a2c4e;
                    }      
                </style>   
            </head>
        ";
        //Aspectos para el envio de correo
        $asunto = "Reporte de rutina , Genset : ".$R['nombre_generador']."  dia : ".$fechaValida;
        $mensaje .= "<h2> Dear : ".$usuario_seleccionado['nombres']." ".$usuario_seleccionado['apellidos']."</h2>";
        $mensaje .= "<h2> Genset : ".$R['nombre_generador']."</h2>";
        $mensaje .= "<h2> Message : "." Reporte de Rutina "."</h2>";
        $mensaje .="<body><table  style='border:1px solid #1a2c4e' ><thead><tr ><th width='140'>Reception Date</th><th width='60'>Battery Voltage </th><th > Water Temp</th><th>Running Frequency</th><th>Fuel Level </th><th>Voltage Measure </th><th>Rotor Current</th><th>Fiel Current</th><th>";
        $mensaje .="Speed </th><th>Eco Power</th><th> RPM </th><th> Unit Mode</th><th>Horometro</th><th>Model </th><th>Latitude </th></th>";
        $mensaje .="Length </th><th>Alarm</th><th>Event</th><th>Reefer Conected</th><th>Set Point</th><th>Temp Supply </th><th>Return Air </th></tr></thead><tbody>";
        foreach($Completo as $fila){
            $mensaje .="<tr align='center' ><td><strong>".$fila['created_at']."</strong></td><td>".$fila['battery_voltage']."</td></td>".$fila['water_temp']."</td><td>".$fila['running_frequency']."</td></td>".$fila['fuel_level']."</td><td>".$fila['voltage_measure']."</td></td>".$fila['rotor_current']." </td><td>".$fila['fiel_current'];
            $mensaje .="</td></td>".$fila['speed']."</td><td>".$fila['eco_power']."</td></td>".$fila['rpm']."</td><td>".$fila['unit_mode']."</td></td>".$fila['horometro']."</td><td>".$fila['modelo']."</td></td>".$fila['latitud']."</td><td>";
            $mensaje .=$fila['longitud']."</td><td>".$fila['alarma_id']."</td></td>".$fila['evento_id']."</td><td>".$fila['reefer_conected']."</td></td>".$fila['set_point']."</td><td>".$fila['temp_supply']."</td></td>".$fila['return_air']."</td></tr>";
        }
        $mensaje .="</tbody></table></body>";
        try {
            //Server settings
            $mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
            $mail->isSMTP();    
            $mail->From = "zgroup.telemetria@gmail.com";                                    //Send using SMTP
            $mail->Host       = "smtp.gmail.com";                   //Set the SMTP server to send through
            $mail->SMTPAuth   = true;                                   //Enable SMTP authentication
            $mail->Username   = 'zgroup.telemetria@gmail.com';                     //SMTP username
            $mail->Password   = 'cydmtmwcyyavlato';                               //SMTP password
            $mail->SMTPSecure = 'tls';            //Enable implicit TLS encryption
            $mail->Port       = 587;                                    //TCP port to connect to; use 587 if you have set `SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS`
            //Agregar destinatario
            $mail->AddAddress($correoEnvio);
            $mail->Subject = $asunto;
            $mail->Body =$mensaje;
            $mail->isHTML(true);
            $mail->AddAttachment('./excel/'.$nombreContenedor.'_'.$fechaExcel.'.xlsx', $nombreContenedor.'_'.$fechaExcel.'.xlsx');
            //Avisar si fue enviado o no y dirigir al index
            if ($mail->Send()) {
                echo'<script type="text/javascript">alert("Enviado Correctamente");</script>';  
            } else {
               echo'<script type="text/javascript">alert("NO ENVIADO, intentar de nuevo");</script>';
            }    
        }catch (Exception $e) {
            echo "Se ha producido un mensaje de error . Mailer Error: {$mail->ErrorInfo}"; 
        }
    break;
        case 2:
            //Estamos en el tipos de reefer
            $R = $api->madurador($telemetria_id);
            $Completo = $api->listaMaduradorFecha($telemetria_id ,$fechaInicio , $fechaFin);
            $reporte = $api->reporteMadurador($telemetria_id ,$fechaInicio , $fechaFin);
            $empresa =$api->empresa($R['empresa_id']);
            $usuario_empresa = $api->usuario_id($R['empresa_id']);   
            $usuario_seleccionado = $api->usuario_enviar($usuario_empresa['usuario_id']);
            //creacion de Excel para adjuntar a correo
            $tituloExcel =" Reporte de rutina para Madurador";
            $documento = new Spreadsheet();
            $documento
            ->getProperties()
            ->setCreator("Luis Pablo Marcelo Perea")
            ->setLastModifiedBy('Helios Tec')
            ->setTitle($tituloExcel)
            ->setDescription('Detalles del comportamiento del Madurador');
            $hojaDeProductos = $documento->getActiveSheet();
            $hojaDeProductos->setTitle("Madurador");
            # Encabezado de los productos
            $nombreContenedor = $R['nombre_contenedor'];
            $dispositivo = ["Ripener :" ,$nombreContenedor];
            $encabezado = ["Reception Date", "Set Point", "Temp Supply", "Return Air", "Evaporation Coil","Ambient Air","Relative Humidity","Controlling Mode","Sp Ethylene","Ethylene ", " AVL", "Power State","Compress","Current ph1","Current ph2","Current ph3","Co2_reading","O2_reading","Set_point_o2","Set_point_co2","Voltage","Defrost Term Temp","Defrost Interval","Latitude","Length"];
            # El último argumento es por defecto A1
            $hojaDeProductos->fromArray($dispositivo, null, 'A1');
            $hojaDeProductos->fromArray($encabezado, null, 'A2');
            # Comenzamos en la fila 3
            $numeroDeFila = 3;
            foreach($reporte as $fila){
                # Escribir registros en el documento
                $hojaDeProductos->setCellValueByColumnAndRow(1, $numeroDeFila, $fila['created_at']);
                $hojaDeProductos->setCellValueByColumnAndRow(2, $numeroDeFila, $fila['set_point']);
                $hojaDeProductos->setCellValueByColumnAndRow(3, $numeroDeFila, $fila['temp_supply_1']);
                $hojaDeProductos->setCellValueByColumnAndRow(4, $numeroDeFila, $fila['return_air']);
                $hojaDeProductos->setCellValueByColumnAndRow(5, $numeroDeFila, $fila['evaporation_coil']);
                $hojaDeProductos->setCellValueByColumnAndRow(6, $numeroDeFila, $fila['ambient_air']);
                $hojaDeProductos->setCellValueByColumnAndRow(7, $numeroDeFila, $fila['relative_humidity']);
                $hojaDeProductos->setCellValueByColumnAndRow(8, $numeroDeFila, $fila['controlling_mode']);
                $hojaDeProductos->setCellValueByColumnAndRow(9, $numeroDeFila, $fila['sp_ethyleno']);
                $hojaDeProductos->setCellValueByColumnAndRow(10, $numeroDeFila, $fila['ethylene']);
                $hojaDeProductos->setCellValueByColumnAndRow(11, $numeroDeFila, $fila['avl']);
                $hojaDeProductos->setCellValueByColumnAndRow(12, $numeroDeFila, $fila['power_state']);
                $hojaDeProductos->setCellValueByColumnAndRow(13, $numeroDeFila, $fila['compress_coil_1']);
                $hojaDeProductos->setCellValueByColumnAndRow(14, $numeroDeFila, $fila['consumption_ph_1']);
                $hojaDeProductos->setCellValueByColumnAndRow(15, $numeroDeFila, $fila['consumption_ph_2']);
                $hojaDeProductos->setCellValueByColumnAndRow(16, $numeroDeFila, $fila['consumption_ph_3']);
                $hojaDeProductos->setCellValueByColumnAndRow(17, $numeroDeFila, $fila['co2_reading']);
                $hojaDeProductos->setCellValueByColumnAndRow(18, $numeroDeFila, $fila['o2_reading']);
                $hojaDeProductos->setCellValueByColumnAndRow(19, $numeroDeFila, $fila['set_point_o2']);
                $hojaDeProductos->setCellValueByColumnAndRow(20, $numeroDeFila, $fila['set_point_co2']);
                $hojaDeProductos->setCellValueByColumnAndRow(21, $numeroDeFila, $fila['line_voltage']);
                $hojaDeProductos->setCellValueByColumnAndRow(22, $numeroDeFila, $fila['defrost_term_temp']);
                $hojaDeProductos->setCellValueByColumnAndRow(23, $numeroDeFila, $fila['defrost_interval']);
                $hojaDeProductos->setCellValueByColumnAndRow(24, $numeroDeFila, $fila['latitud']);
                $hojaDeProductos->setCellValueByColumnAndRow(25, $numeroDeFila, $fila['longitud']);
                $numeroDeFila++;
            }
            # Crear un "escritor"
            $writer = new Xlsx($documento);
            # Le pasamos la ruta de guardado
            $writer->save('./excel/'.$nombreContenedor.'_'.$fechaExcel.'.xlsx');
            $mensaje .="
                 <head>
                    <style>
                        table {
                            border-collapse: collapse;
                            margin: 25px 0; 
                            font-size: 1em;
                            font-family: sans-serif;
                            min-width: 1450px; 
                            box-shadow: 0 0 20px rgba(0, 0, 0, 0.15);
                        }
                        thead tr {
                            background-color: #1a2c4e;
                            color: #ffffff;
                            text-align: middle;
                        }
                        th, td {
                            padding: 12px 15px;
                        }      
                        tbody tr{
                            border-bottom: 1px solid #1a2c4e;
                        }      
                    </style>   
                </head>
            ";
            //Aspectos para el envio de correo
            $asunto = "Reporte de rutina , Madurador 2: ".$R['nombre_contenedor']."  dia : ".$fechaValida;
            $mensaje .= "<h2> Dear : ".$usuario_seleccionado['nombres']." ".$usuario_seleccionado['apellidos']."</h2>";
            $mensaje .= "<h2> Ripener : ".$R['nombre_contenedor']."</h2>";
            $mensaje .= "<h2> Message : "." Reporte de Rutina "."</h2>";
            $mensaje .="<body><table  style='border:1px solid #1a2c4e' ><thead><tr ><th width='180'>Reception Date</th><th width='60'>Set Point </th><th > Temp Supply </th><th>Return Air </th><th>Evaporation Coil </th><th>Ambient Air </th><th>Relative Humidity</th><th>Controlling Mode</th><th>";
            $mensaje .="Sp Ethylene </th><th>Ethylene</th><th> AVL </th><th> Power State </th><th>Compress</th><th>Current Ph1 </th><th>Current Ph2 </th><th>Current Ph3 </th><th>Co2_reading </th><th>O2_reading </th><th>Set Point O2 </th><th>Set Point co2 </th><th>Voltage </th>";
            $mensaje .="<th>Defrost Term Temp </th><th>Defrost Interval </th><th>Latitude </th><th>Length </th></tr></thead><tbody>";
            foreach($Completo as $fila){
                $mensaje .="<tr align='center' ><td width='180'><strong>".$fila['created_at']."</strong></td><td>".$fila['set_point']."</td></td>".$fila['temp_supply_1']."</td><td>".$fila['return_air']."</td></td>".$fila['evaporation_coil']."</td><td>".$fila['ambient_air']."</td></td>".$fila['relative_humidity']." %</td><td>".$fila['controlling_mode'];
                $mensaje .="</td></td>".$fila['sp_ethyleno']."</td><td>".$fila['ethylene']."</td></td>".$fila['avl']."</td><td>".$fila['power_state']."</td></td>".$fila['compress_coil_1']."</td><td>".$fila['consumption_ph_1']."</td></td>".$fila['consumption_ph_2']."</td><td>";
                $mensaje .=$fila['consumption_ph_3']."</td><td>".$fila['co2_reading']."</td></td>".$fila['o2_reading']."</td><td>".$fila['set_point_o2']."</td></td>".$fila['set_point_co2']."</td><td>".$fila['line_voltage']."</td></td>".$fila['defrost_term_temp']."</td><td>";
                $mensaje .=$fila['defrost_interval']."</td><td>".$fila['latitud']."</td></td>".$fila['longitud']."</td><tr>";
            }
            $mensaje .="</tbody></table></body>";
            try {
                //Server settings
                $mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
                $mail->isSMTP();    
                $mail->From = "ztrack@zgroup.com.pe";  
                //$mail->From = "desarrollo@zgroup.com.pe";                                     //Send using SMTP
                $mail->Host       = "smtp.gmail.com";                   //Set the SMTP server to send through
                $mail->SMTPAuth   = true;                                   //Enable SMTP authentication
                //$mail->Username   = 'desarrollo@zgroup.com.pe';                     //SMTP username
                //$mail->Password   = 'Des5090100';  
                $mail->Username   = 'ztrack@zgroup.com.pe';                     //SMTP username
                $mail->Password   = 'Proyectoztrack2023!';                               //SMTP password
                $mail->SMTPSecure = 'tls';            //Enable implicit TLS encryption
                $mail->Port       = 587;                                    //TCP port to connect to; use 587 if you have set `SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS`
                //Agregar destinatario
                $mail->AddAddress($correoEnvio);
                $mail->Subject = $asunto;
                $mail->Body =$mensaje;
                $mail->isHTML(true);
                $mail->AddAttachment('./excel/'.$nombreContenedor.'_'.$fechaExcel.'.xlsx', $nombreContenedor.'_'.$fechaExcel.'.xlsx');
                //Avisar si fue enviado o no y dirigir al index
                if ($mail->Send()) {
                    echo'<script type="text/javascript">alert("Enviado Correctamente");</script>';  
                } else {
                   echo'<script type="text/javascript">alert("NO ENVIADO, intentar de nuevo");</script>';
                }    
            } catch (Exception $e) {
                echo "Se ha producido un mensaje de error . Mailer Error: {$mail->ErrorInfo}"; 
            }
        break;
}

?>