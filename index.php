<?php

require_once('./Docx.php');

$filesDir = './';
$files_ar = scandir($filesDir);

echo "<h1>php-docx test</h1>";

foreach($files_ar as $entry){
    // check file type
    if(strtolower(substr($entry, -5)) !== ".docx") continue;

    $aDocx = new Docx($filesDir.$entry);
    echo "<div>";
    echo "<p><h3>Origin file:</h3><span><a href='".$entry."' target=_blank>".$entry."</a></span></p>"; 
    echo "<p><h3>Text:</h3><span>".$aDocx->toTxt()."</span></p>"; 
    echo "<p><h3>HTML:</h3><span>".$aDocx->toHtml()."</span></p>"; 
    echo "</div>";
}

?>
