<?php
error_reporting(E_ALL & ~E_DEPRECATED & ~E_STRICT & ~E_NOTICE & ~E_WARNING); 
require_once('config.inc.php');

header("Access-Control-Allow-Origin: *");
header("Access-Control-Allow-Methods: GET, POST, OPTIONS");
header("Access-Control-Allow-Headers: Content-Type, cache-control, Authorization, X-Requested-With, satoken, Token");
header("Content-type: application/json");
header('Cache-Control: no-cache');

// Handle preflight requests
if ($_SERVER['REQUEST_METHOD'] === 'OPTIONS') {
    exit(0);
}

if($_GET['action']=='getConfig')  {
    $RS = [];
    $RS['status']       = 'ok';
    $RS['API_URL']      = $API_URL;
    $RS['API_MODE']     = $API_MODE;
    print json_encode($RS);
    exit;
}

$_POST = json_decode(file_get_contents("php://input"), true);

if($_POST['aiApiUrl'] && $_POST['aiModel'] && $_POST['aiToken']) {

    $redis->set("API_URL", $_POST['aiApiUrl']);
    $redis->set("API_MODE", $_POST['aiModel']);
    $redis->set("API_KEY", $_POST['aiToken']);

    $RS = [];
    $RS['status']   = 'ok';
    $RS['msg']      = '参数配置成功';
    print json_encode($RS);
}
else {
    $RS = [];
    $RS['status']   = 'error';
    $RS['msg']      = '三个参数都需要同时填写';
    print json_encode($RS);
}


?>
