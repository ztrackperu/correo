<?php
require_once 'config.php';
require_once 'conexion.php';
class ApiModel{
    private $pdo, $con;
    public function __construct() {
        $this->con = new Conexion();
        $this->pdo = $this->con->conectar();
    }
    public function reefer($telemetria){
        $valor =intval($telemetria);
        $consult = $this->pdo->prepare("SELECT *  FROM contenedores WHERE telemetria_id = ? and tipo ='Reefer'");
        $consult->execute([$valor]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }
    public function listaReeferFecha($id ,$fechaInicio , $fechaFin)
    {
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT * FROM registro_reefers WHERE telemetria_id = ? and created_at >= ? and created_at <= ? ORDER BY id desc limit 20");
        $consult->execute([$valor,$fechaInicio , $fechaFin]);
        return $consult->fetchAll(PDO::FETCH_ASSOC);
    }
    public function reporteReefer($id ,$fechaInicio , $fechaFin)
    {
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT * FROM registro_reefers WHERE telemetria_id = ? and created_at >= ? and created_at <= ? ORDER BY id desc ");
        $consult->execute([$valor,$fechaInicio , $fechaFin]);
        return $consult->fetchAll(PDO::FETCH_ASSOC);
    }
    public function empresa($id){
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT *  FROM empresas WHERE id =? ");
        $consult->execute([$valor]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }
    public function usuario_id($id){
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT *  FROM usuario_empresa WHERE empresa_id =? ");
        $consult->execute([$valor]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }
    public function usuario_enviar($id){
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT *  FROM usuarios WHERE id =? ");
        $consult->execute([$valor]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }
    
    //consultas para maduradres 
    public function madurador($telemetria){
        $valor =intval($telemetria);
        $consult = $this->pdo->prepare("SELECT *  FROM contenedores WHERE telemetria_id = ? and tipo ='Madurador'");
        $consult->execute([$valor]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }
    public function listaMaduradorFecha($id ,$fechaInicio , $fechaFin)
    {
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT * FROM registro_madurador WHERE telemetria_id = ? and created_at >= ? and created_at <= ? ORDER BY id desc limit 20");
        $consult->execute([$valor,$fechaInicio , $fechaFin]);
        return $consult->fetchAll(PDO::FETCH_ASSOC);
    }
    public function reporteMadurador($id ,$fechaInicio , $fechaFin)
    {
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT * FROM registro_madurador WHERE telemetria_id = ? and created_at >= ? and created_at <= ? ORDER BY id desc limit 30 ");
        $consult->execute([$valor,$fechaInicio , $fechaFin]);
        return $consult->fetchAll(PDO::FETCH_ASSOC);
    }
    //consultas para generadres 
    public function generador($telemetria){
        $valor =intval($telemetria);
        $consult = $this->pdo->prepare("SELECT *  FROM generadores WHERE telemetria_id = ? ");
        $consult->execute([$valor]);
        return $consult->fetch(PDO::FETCH_ASSOC);
    }
    public function listaGeneradorFecha($id ,$fechaInicio , $fechaFin)
    {
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT * FROM registro_generador WHERE telemetria_id = ? and created_at >= ? and created_at <= ? ORDER BY id desc limit 20");
        $consult->execute([$valor,$fechaInicio , $fechaFin]);
        return $consult->fetchAll(PDO::FETCH_ASSOC);
    }
    public function reporteGenerador($id ,$fechaInicio , $fechaFin)
    {
        $valor =intval($id);
        $consult = $this->pdo->prepare("SELECT * FROM registro_generador WHERE telemetria_id = ? and created_at >= ? and created_at <= ? ORDER BY id desc ");
        $consult->execute([$valor,$fechaInicio , $fechaFin]);
        return $consult->fetchAll(PDO::FETCH_ASSOC);
    }
    

}