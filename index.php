<?php

if(isset($_POST['req'])){
    
    $valb64 = $_POST['req'];  
    $rand = random_strings(5);
    $archivo = fopen('datos'.$rand.'.txt', 'a');

    fwrite($archivo, base64_decode($valb64) . "\n");

    fclose($archivo);
} else {    
    echo "Error: No se ha proporcionado el parámetro";
}

function random_strings($length_of_string){
    $str_result = '0123456789ABCDEFGHIJKLMNOPQRSTUVWXYZabcdefghijklmnopqrstuvwxyz';
    return md5(substr(str_shuffle($str_result),0, $length_of_string));
}

?>
