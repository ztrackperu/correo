

<?php
 
use PHPMailer\PHPMailer\PHPMailer;
use PHPMailer\PHPMailer\SMTP;
use PHPMailer\PHPMailer\Exception;

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
//Create a new PHPMailer instance
//$mail =new PHPMailer(true);
require 'vendor/autoload.php';
$mail = new PHPMailer(true);
$mensaje = "";
require_once 'api_correo.php';
$api = new ApiModel();
$telemetria_id =4;
$fechaActual = date("Y-m-d");
$fechaInicio =date("Y-m-d",strtotime($fechaActual."- 1 days"))." 00:00:00";
$fechaFin =date("Y-m-d H:i:s");
$R = $api->reefer($telemetria_id);
$Completo = $api->listaReeferFecha($telemetria_id ,$fechaInicio , $fechaFin);
$reporte = $api->reporteReefer($telemetria_id ,$fechaInicio , $fechaFin);
$empresa =$api->empresa($R['empresa_id']);
$usuario_empresa = $api->usuario_id($R['empresa_id']);
$usuario_seleccionado = $api->usuario_enviar($usuario_empresa['usuario_id']);
//echo var_dump($empresa);
//echo var_dump($usuario_empresa);
//echo var_dump($usuario_seleccionado);

$tituloExcel =" Reporte de rutina A";
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
$dispositivo = ["Reefer :" ,"ZGRU102020"];
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
$writer->save('./excel/reporte1.xlsx');
$mensaje .="
   <head>
     <style>
     table {
        border-collapse: collapse;
         margin: 25px 0; 
         font-size: 1em;
         font-family: sans-serif;
          min-width: 450px; 
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
$Hoy = date("d/m/Y");
$hoy1= date("d/m/Y",strtotime($Hoy."- 1 days"));
$asunto = "Reporte de rutina , Reefer1: ".$R['nombre_contenedor']."  dia : ".$hoy1;
$correoEnvio = "desarrollozgroup@gmail.com";
//echo var_dump($R);
//echo $R['nombre_contenedor'];
$mensaje .= "<h2> Dear : ".$usuario_seleccionado['nombres']." ".$usuario_seleccionado['apellidos']."</h2>";
$mensaje .= "<h2> Reefer : ".$R['nombre_contenedor']."</h2>";
$mensaje .= "<h2> Message : "." Reporte de Rutina "."</h2>";
$mensaje .="<body><table  style='border:1px solid #1a2c4e' ><thead><tr ><th width='140'>Reception Date</th><th width='60'>Set Point </th><th > Temp Supply </th><th>Return Air </th><th>Evaporation Coil </th><th>Ambient Air </th><th>Relative Humidity</th><th>Alarm Present</th><th>";
$mensaje .="Alarm Number </th><th>Controlling Mode</th><th> Power State </th><th> Defrost Term Temp </th><th>Defrost Interval</th><th>Latitude </th><th>Length </th></tr></thead><tbody>";
foreach($Completo as $fila){
    $mensaje .="<tr align='center' ><td><strong>".$fila['created_at']."</strong></td><td>".$fila['set_point']."</td></td>".$fila['temp_supply']."</td><td>".$fila['return_air']."</td></td>".$fila['evaporation_coil']."</td><td>".$fila['ambient_air']."</td></td>".$fila['relative_humidity']." %</td><td>".$fila['alarm_present'];
    $mensaje .="</td></td>".$fila['alarm_number']."</td><td>".$fila['controlling_mode']."</td></td>".$fila['power_state']."</td><td>".$fila['defrost_term_temp']."</td></td>".$fila['defrost_interval']."</td><td>".$fila['latitud']."</td></td>".$fila['longitud']."</td></tr>";
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
    //$mail->AddAddress($_POST['email']);
    $mail->AddAddress($correoEnvio);
    $mail->Subject = $asunto;
    //$mail->Subject = $_POST['subject'];
    $mail->Body =$mensaje;
    $mail->isHTML(true);

    $mail->AddAttachment('./excel/reporte1.xlsx' , 'reporte1.xlsx');

    //Avisar si fue enviado o no y dirigir al index
if ($mail->Send()) {
    echo'<script type="text/javascript">
           alert("Enviado Correctamente");
        </script>';
} else {
    echo'<script type="text/javascript">
           alert("NO ENVIADO, intentar de nuevo");
        </script>';
}
    
} catch (Exception $e) {
    //echo "Message could not be sent. Mailer Error: {$mail->ErrorInfo}";
}

?>