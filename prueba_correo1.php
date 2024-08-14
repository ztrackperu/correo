<?php
        require '../ztrack3/vendor/autoload.php';
        //use Exception;
        use MongoDB\Client;
        use MongoDB\Driver\ServerApi;
        use MongoDB\BSON\UTCDateTime ;

use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
require 'vendor/autoload.php';
//aqui haremos la consulta en mongo db 
require_once 'config.php';
require_once 'conexion.php';

class CorreoModel{

    private $pdo, $con;
    public function __construct() {
        $this->con = new Conexion();
        $this->pdo = $this->con->conectar();
    }
    public function consultarDispositivo($telemetria_id )
    {
        $consult = $this->pdo->prepare("SELECT nombre_contenedor,descripcionC,empresa_id,set_point,temp_supply_1,return_air, evaporation_coil,ambient_air,relative_humidity, power_state ,ultima_fecha,updated_at ,extra_1 FROM contenedores WHERE telemetria_id = ? AND estado = 1");
        $consult->execute([$telemetria_id]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }
    public function ver_detalle_alarma($alarma )
    {
        $consult = $this->pdo->prepare("SELECT * FROM nombre_alarma WHERE codigo = ? ");
        $consult->execute([$alarma]);
        return $consult->fetch(PDO::FETCH_ASSOC);

    }
    public function estado_dispositivo($telemetria_id )
    {
        $consult = $this->pdo->prepare("SELECT * FROM control_dispositivos WHERE estado_control=1 and telemetria_id = ? ");
        $consult->execute([$telemetria_id]);
        return $consult->fetch(PDO::FETCH_ASSOC);

    }
    public function estado_dispositivo_salog($telemetria_id )
    {
        $consult = $this->pdo->prepare("SELECT * FROM control_dispositivos WHERE estado_control=2 and telemetria_id = ? ");
        $consult->execute([$telemetria_id]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }

    public function excel_1($fechaFin,$telemetria_id)
    {
        //$consult = $this->pdo->prepare("SELECT id_permiso FROM detalle_permisos WHERE id_permiso = ? AND id_usuario = ?");
        //$consult->execute([$permiso, $id_usuario]);
        //return $consult->fetch(PDO::FETCH_ASSOC);


        $uri = 'mongodb://localhost:27017';
        // Specify Stable API version 1
        $apiVersion = new ServerApi(ServerApi::V1);
        // Create a new client and connect to the server
        $client = new MongoDB\Client($uri, [], ['serverApi' => $apiVersion]);
        $fechaaInicio = strtotime($fechaFin);
        $fechaaInicio1 = strtotime("-12 hours",$fechaaInicio);
        $fechaaInicio2 = date("Y-m-d H:i:s",$fechaaInicio1);

        //problemas con fecha 5 horas menos debe ser UTC-5
        $puntoA = strtotime($fechaaInicio2);
        $puntoA1 = strtotime("-5 hours",$puntoA)*1000;
        $puntoB = strtotime($fechaFin)  ;
        $puntoB1 = strtotime("-5 hours" ,$puntoB)*1000  ;
        // se selcciona los campos y las fechas 
        
        $cursor  = $client->ztrack_ja->madurador->find(array('$and' =>array( ['created_at'=>array('$gte'=>new MongoDB\BSON\UTCDateTime($puntoA1),'$lte'=>new MongoDB\BSON\UTCDateTime($puntoB1)),'telemetria_id'=>intval($telemetria_id)] )),
        array('projection' => array('_id' => 0,'trama'=> 1, 'created_at' => 1,'stateProcess' => 1,'set_point' => 1,'temp_supply_1' => 1,'return_air' => 1,'evaporation_coil' => 1,'ambient_air' => 1,'relative_humidity' => 1,'controlling_mode' => 1,'sp_ethyleno' => 1,
        'ethylene' => 1,'avl' => 1,'power_state' => 1,'compress_coil_1' => 1,'consumption_ph_1' => 1,'consumption_ph_2' => 1,'consumption_ph_3' => 1,'co2_reading' => 1,'o2_reading' => 1,'set_point_o2' => 1,'set_point_co2' => 1,'line_voltage' => 1,
        'defrost_term_temp' => 1,'defrost_interval' => 1,'inyeccion_pwm' => 1,'inyeccion_hora' => 1,'latitud' => 1,'longitud' => 1,'fresh_air_ex_mode' => 1,'telemetria_id'=>1,'cargo_1_temp' =>1,'cargo_2_temp' =>1,'cargo_3_temp' =>1,'cargo_4_temp' =>1,'id'=>1 ,'power_kwh' =>1),'sort'=>array('id'=>1)));
       

       
       
        $total['madurador'] = [];
        foreach ($cursor as $document) {
            array_unshift($total['madurador'],$document);

        }

        return $total['madurador'];

    }
    public function envio_correo2($fechaFin,$telemetria_id,$codigo_alarma)
    {
           
        $mail = new PHPMailer(true);
        $mensaje = "";
       // $mail = new PHPMailer(true);
        $fechaZ =date("Y-m-d_H-i-s");  

        //$correoEnvio = "devpablito2023@gmail.com";
        $correoEnvio ="ztrack@zgroup.com.pe";
        $correoEnvio1 = "atencionalcliente@zgroup.com.pe";
        //$correoEnvio5 = "atencionalcliente2@zgroup.com.pe";
        $correoEnvio5 = "desarrollozgroup@gmail.com";

        $correoEnvio2 = "ingenieria@zgroup.com.pe";
        $correoEnvio3 = "customer@zgroup.com.pe";
        $correoEnvio4 = "informes@zgroup.com.pe";
        //$correoEnvio = "atencionalcliente@zgroup.com.pe";
        //Estamos en el tipos de reefer
        $detalle_alarma =$this->ver_detalle_alarma($codigo_alarma);
        $detalle_estado =$this->estado_dispositivo($telemetria_id);
        $dataDispositivo =$this->consultarDispositivo($telemetria_id);
        
        $nombreContenedor = $dataDispositivo['nombre_contenedor'];

        $tituloExcel =$nombreContenedor." Últimas 12 horas ";
       // $reporte = $this->excel_1($fechaFin,$telemetria_id);

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
        
        $dispositivo = ["Reefer :" ,$nombreContenedor];
        $encabezado = ["Reception Date", "Set Point", "Temp Supply", "Return Air", "Evaporation Coil","Ambient Air","Relative Humidity","Power State","Defrost Term Temp","Defrost Interval"];
        # El último argumento es por defecto A1
        $hojaDeProductos->fromArray($dispositivo, null, 'A1');
        $hojaDeProductos->fromArray($encabezado, null, 'A2');
        # Comenzamos en la fila 3
        $numeroDeFila = 3;
        /*
        foreach($reporte as $fila){
            # Escribir registros en el documento
            $evaluadorApagado = $fila['power_state'];
            if($evaluadorApagado==1){
               $on_off= "ON";
            }else{
                $on_off= "OFF";

            }
            $datoFecha = strval($fila['created_at']);
            
            $datoFecha1 = intval($datoFecha)/1000;
            
            $hora_total =date("Y-m-d H:i:s" , $datoFecha1);

            //$hora_total =date("Y-m-d H:i:s" , $datoFecha1);
            $hora_total3 = strtotime($hora_total);
            $hora_total1 = strtotime("+5 hours" ,$hora_total3);
            $hora_total2 =date("Y-m-d H:i:s" , $hora_total1);



            $hojaDeProductos->setCellValueByColumnAndRow(1, $numeroDeFila,$hora_total2);
            $hojaDeProductos->setCellValueByColumnAndRow(2, $numeroDeFila, $fila['set_point']);
            $hojaDeProductos->setCellValueByColumnAndRow(3, $numeroDeFila, $fila['temp_supply_1']);
            $hojaDeProductos->setCellValueByColumnAndRow(4, $numeroDeFila, $fila['return_air']);
            $hojaDeProductos->setCellValueByColumnAndRow(5, $numeroDeFila, $fila['evaporation_coil']);
            $hojaDeProductos->setCellValueByColumnAndRow(6, $numeroDeFila, $fila['ambient_air']);
            $hojaDeProductos->setCellValueByColumnAndRow(7, $numeroDeFila, $fila['relative_humidity']);
            $hojaDeProductos->setCellValueByColumnAndRow(8, $numeroDeFila, $on_off);
            $hojaDeProductos->setCellValueByColumnAndRow(9, $numeroDeFila, $fila['defrost_term_temp']);
            $hojaDeProductos->setCellValueByColumnAndRow(10, $numeroDeFila, $fila['defrost_interval']);

            $numeroDeFila++;
        }
        # Crear un "escritor"
        $writer = new Xlsx($documento);
        # Le pasamos la ruta de guardado
        $writer->save('./excel/'.$nombreContenedor.'_'.$fechaZ.'.xlsx');
        */

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
        $cliente = $dataDispositivo['descripcionC'];
        $mensaje_alarma= $detalle_alarma['mensaje'];

        $concidencia1 = strpos($mensaje_alarma,"CONEXION");
        $concidencia2 = strpos($mensaje_alarma,"APAGADO");

        if($concidencia1==true){
            $inicioApagado = $detalle_estado['ultimo_dato'];

        }elseif ($concidencia2==true){
            $inicioApagado = $detalle_estado['estado_on'];
        }else{
            $inicioApagado = $detalle_estado['temp_ok'];
        }




        
        $asunto = "URGENTE ,Alarma  en  ".$nombreContenedor." - " .$cliente."  dia : ".$fechaFin;
        $mensaje .= "<h2> Señores : ZGROUP CENTRAL </h2>";
        $mensaje .= "<h2> Dispositivo : ".$nombreContenedor."</h2>";
        $mensaje .= "<h2> Cliente : ".$cliente."</h2>";
        $mensaje .= "<h2> Mensaje : ".$mensaje_alarma."</h2>";
        $mensaje .= "<h2> Desde : ".$inicioApagado."</h2>";
        $mensaje .= "<h2>* Último estado recibido  : </h2>";
        $mensaje .="<body><table  style='border:1px solid #1a2c4e' ><thead><tr ><th width='130'>Reception Date</th><th width='60'>Set Point </th><th > Temp Supply </th><th>Return Air </th><th>Evaporation Coil </th>";
        $mensaje .=" </tr></thead><tbody>";
        
        $mensaje .="<tr align='center' ><td width='130'><strong>".$dataDispositivo['ultima_fecha']."</strong></td><td>".$dataDispositivo['set_point']." C°</td></td>".$dataDispositivo['temp_supply_1']." C°</td><td>".$dataDispositivo['return_air']." C°</td></td>".$dataDispositivo['evaporation_coil']." C°</td>";
        $mensaje .="</tr>";
           
        $mensaje .="</tbody></table></body>";
        $mensaje .= "<h3>Temperatura Ambiente : ".$dataDispositivo['ambient_air']." C°</h3>";
        $mensaje .= "<h3>Humedad Relativa: ".$dataDispositivo['relative_humidity']." %</h3>";
        if($dataDispositivo['power_state']==1){
            $epa="ENCENDIDO";
        }else{
            $epa="APAGADO";
        }
        $mensaje .= "<h3>Estado : ".$epa."</h3>";

    


        try {
            //Server settings
            $mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
            $mail->isSMTP();    
           // $mail->From = "ztrack@zgroup.com.pe"; 
            $mail->From = "devpablito2023@gmail.com";                                   //Send using SMTP
            $mail->Host       = "smtp.gmail.com";                   //Set the SMTP server to send through
            $mail->SMTPAuth   = true;                                   //Enable SMTP authentication
            $mail->Username   = 'devpablito2023@gmail.com';                     //SMTP username
            $mail->Password   = 'fdcjahqtohijkkoc';                               //SMTP password
            //$mail->Username   = 'ztrack@zgroup.com.pe';                     //SMTP username
            //$mail->Password   = 'Proyectoztrack2023!';
                               //SMTP password
            $mail->SMTPSecure = 'tls';            //Enable implicit TLS encryption
            $mail->Port       = 587;                                    //TCP port to connect to; use 587 if you have set `SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS`
            //Agregar destinatario
            $mail->AddAddress($correoEnvio);
            $mail->AddAddress($correoEnvio1);
            $mail->AddAddress($correoEnvio2);
            $mail->AddAddress($correoEnvio3);
            $mail->AddAddress($correoEnvio4);
            $mail->AddAddress($correoEnvio5);
            $mail->Subject = utf8_decode($asunto);
            $mail->Body =utf8_decode($mensaje);
            $mail->isHTML(true);
            //$mail->AddAttachment('./excel/'.$nombreContenedor.'_'.$fechaZ.'.xlsx', $nombreContenedor.'_'.$fechaZ.'.xlsx');
            //Avisar si fue enviado o no y dirigir al index
            if ($mail->Send()) {
                echo'<script type="text/javascript">alert("Enviado Correctamente");</script>';  
            } else {
               echo'<script type="text/javascript">alert("NO ENVIADO, intentar de nuevo");</script>';
            }    
        }catch (Exception $e) {
            echo "Se ha producido un mensaje de error . Mailer Error: {$mail->ErrorInfo}"; 
        }


    }





    public function envio_correo_salog($fechaFin,$telemetria_id,$codigo_alarma)
    {     
        $mail = new PHPMailer(true);
        $mensaje = "";
        $fechaZ =date("Y-m-d_H-i-s"); 
        $correoEnvio ="ztrack@zgroup.com.pe"; 
        //lista de correos
        $correoEnvio1 ="fjorge@salog.pe";
        $correoEnvio2 ="mdiestra@salog.com.pe";
        $correoEnvio3 ="mceron@salog.com.pe";
        $correoEnvio4 ="rcaceres@salog.com.pe";
        $correoEnvio5 ="rconde@salutare.pe";
        $correoEnvio6 ="kvera@salutare.com.pe";
        $correoEnvio7 ="pvasquez@salog.pe";
        $correoEnvio8 ="centrocontrol@salog.com.pe";
        $correoEnvio9 ="jtaquila@salutare.com.pe";
        $correoEnvio10 ="ti.sitrad@salog.com.pe";
        $correoEnvio11 ="alarrauri@salutare.com.pe";
        $correoEnvio12 ="jzevallos@salutare.pe";
        $correoEnvio13 ="alazaro@salutare.pe";
        $correoEnvio14 ="ylopez@salutare.pe";
        $correoEnvio15 ="jtinoco@salutare.pe";
        //Estamos en el tipos de reefer
        $detalle_alarma =$this->ver_detalle_alarma($codigo_alarma);
        $detalle_estado =$this->estado_dispositivo_salog($telemetria_id);
        $dataDispositivo =$this->consultarDispositivo($telemetria_id);
        
        $nombreContenedor = $dataDispositivo['nombre_contenedor'];

        $tituloExcel =$nombreContenedor." Últimas 12 horas ";
       // $reporte = $this->excel_1($fechaFin,$telemetria_id);

        $documento = new Spreadsheet();
        $documento
        ->getProperties()
        ->setCreator("Luis Pablo Marcelo Perea")
        ->setLastModifiedBy('ZGROUP')
        ->setTitle($tituloExcel)
        ->setDescription('Detalles del comportamiento del reefer');
        $hojaDeProductos = $documento->getActiveSheet();
        $hojaDeProductos->setTitle("Reefer");
        # Encabezado de los productos
        
        $dispositivo = ["Reefer :" ,$nombreContenedor];
        $encabezado = ["Reception Date", "Set Point", "Temp Supply", "Return Air", "Evaporation Coil","Ambient Air","Relative Humidity","Power State","Defrost Term Temp","Defrost Interval"];
        # El último argumento es por defecto A1
        $hojaDeProductos->fromArray($dispositivo, null, 'A1');
        $hojaDeProductos->fromArray($encabezado, null, 'A2');
        # Comenzamos en la fila 3
        $numeroDeFila = 3;
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
        $cliente = $dataDispositivo['descripcionC'];
        $mensaje_alarma= $detalle_alarma['mensaje'];

        $concidencia1 = strpos($mensaje_alarma,"CONEXION");
        $concidencia2 = strpos($mensaje_alarma,"APAGADO");

        if($concidencia1==true){
            $inicioApagado = $detalle_estado['ultimo_dato'];

        }elseif ($concidencia2==true){
            $inicioApagado = $detalle_estado['estado_on'];
        }else{
            $inicioApagado = $detalle_estado['temp_ok'];
        }

        $asunto = "URGENTE ,Alarma  en  ".$nombreContenedor." - " .$cliente."  dia : ".$fechaFin;
        $mensaje .= "<h2> Señores : SALOG CENTRAL </h2>";
        $mensaje .= "<h2> Dispositivo : ".$nombreContenedor."</h2>";
        $mensaje .= "<h2> Referencia : ".$cliente."</h2>";
        $mensaje .= "<h2> Mensaje : ".$mensaje_alarma."</h2>";
        $mensaje .= "<h2> Desde : ".$inicioApagado."</h2>";
        $mensaje .= "<h2>* Último estado recibido  : </h2>";
        $mensaje .="<body><table  style='border:1px solid #1a2c4e' ><thead><tr ><th width='130'>Reception Date</th><th width='60'>Set Point </th><th > Temp Supply </th><th>Return Air </th><th>Evaporation Coil </th>";
        $mensaje .=" </tr></thead><tbody>";
        
        $mensaje .="<tr align='center' ><td width='130'><strong>".$dataDispositivo['ultima_fecha']."</strong></td><td>".$dataDispositivo['set_point']." C°</td></td>".$dataDispositivo['temp_supply_1']." C°</td><td>".$dataDispositivo['return_air']." C°</td></td>".$dataDispositivo['evaporation_coil']." C°</td>";
        $mensaje .="</tr>";
           
        $mensaje .="</tbody></table></body>";
        $mensaje .= "<h3>Temperatura Ambiente : ".$dataDispositivo['ambient_air']." C°</h3>";
        $mensaje .= "<h3>Humedad Relativa: ".$dataDispositivo['relative_humidity']." %</h3>";
        if($dataDispositivo['power_state']==1){
            $epa="ENCENDIDO";
        }else{
            $epa="APAGADO";
        }
        $mensaje .= "<h3>Estado : ".$epa."</h3>";
        try {
            //Server settings
            $mail->SMTPDebug = SMTP::DEBUG_SERVER;                      //Enable verbose debug output
            $mail->isSMTP();    
           // $mail->From = "ztrack@zgroup.com.pe"; 
            $mail->From = "devpablito2023@gmail.com";                                   //Send using SMTP
            $mail->Host       = "smtp.gmail.com";                   //Set the SMTP server to send through
            $mail->SMTPAuth   = true;                                   //Enable SMTP authentication
            $mail->Username   = 'devpablito2023@gmail.com';                     //SMTP username
            $mail->Password   = 'fdcjahqtohijkkoc';                               //SMTP password
            //$mail->Username   = 'ztrack@zgroup.com.pe';                     //SMTP username
            //$mail->Password   = 'Proyectoztrack2023!';
                               //SMTP password
            $mail->SMTPSecure = 'tls';            //Enable implicit TLS encryption
            $mail->Port       = 587;                                    //TCP port to connect to; use 587 if you have set `SMTPSecure = PHPMailer::ENCRYPTION_STARTTLS`
            //Agregar destinatario
            $mail->AddAddress($correoEnvio);
            $mail->AddAddress($correoEnvio1);
            $mail->AddAddress($correoEnvio2);
            $mail->AddAddress($correoEnvio3);
            $mail->AddAddress($correoEnvio4);
            $mail->AddAddress($correoEnvio5);
            $mail->AddAddress($correoEnvio6);
            $mail->AddAddress($correoEnvio7);
            $mail->AddAddress($correoEnvio8);
            $mail->AddAddress($correoEnvio9);
            $mail->AddAddress($correoEnvio10);
            $mail->AddAddress($correoEnvio11);
            $mail->AddAddress($correoEnvio12);
            $mail->AddAddress($correoEnvio13);
            $mail->AddAddress($correoEnvio14);
            $mail->AddAddress($correoEnvio15);
            $mail->Subject = utf8_decode($asunto);
            $mail->Body =utf8_decode($mensaje);
            $mail->isHTML(true);
            //$mail->AddAttachment('./excel/'.$nombreContenedor.'_'.$fechaZ.'.xlsx', $nombreContenedor.'_'.$fechaZ.'.xlsx');
            //Avisar si fue enviado o no y dirigir al index
            if ($mail->Send()) {
                echo'<script type="text/javascript">alert("Enviado Correctamente");</script>';  
            } else {
               echo'<script type="text/javascript">alert("NO ENVIADO, intentar de nuevo");</script>';
            }    
        }catch (Exception $e) {
            echo "Se ha producido un mensaje de error . Mailer Error: {$mail->ErrorInfo}"; 
        }


    }
  
}

?>


