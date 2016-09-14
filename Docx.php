<?php

Class Docx {
    public $fullpath;
    private $dataFile;

    function __construct($fullpath){
        $this->fullpath = $fullpath;
        $this->dataFile = "word/document.xml";
    }

    /**
    * convert contents to text
    */
    public function toTxt(){
        $ret = "";

        $textNodes = $this->_getTextNodes();

        if ($textNodes !== false) {
            // read text
            foreach ($textNodes as $entry) {
                $ret .=$entry->nodeValue."\n";
            }
        }

        return $ret;
    }


    /**
     * Get contents in html format
     */
    public function toHtml(){
        $ret = "";
        $textNodes = $this->_getTextNodes();

        if ($textNodes !== false) {
            $ret .= "<div class='text-wrapper' class='row'>";

            foreach ($textNodes as $entry) {
                $class = " ";
                $style = " ";

                // parse text style and append as css style
                @$testerNode = $entry->firstChild->lastChild->previousSibling;
                if(isset($testerNode) &&
                    $testerNode->nodeName=="w:jc" &&
                    $testerNode->attributes->getNamedItem("val")->nodeValue=="center"
                ){
                    $class .= "center ";
                    $style .= "text-align:center; ";
                }

                @$testerNodes = $entry->firstChild->lastChild->childNodes;
                if(isset($testerNodes)){
                    foreach($testerNodes as $testerNode){
                        //if($testerNode->nodeName=="w:b") $style .= "font-weight:bold; ";
                        if($testerNode->nodeName=="w:bCs") $style .= "font-weight:bold; border-bottom: solid 1px black; ";
                        //if($testerNode->nodeName=="w:szCs") $style .= "border-bottom: solid 1px black; ";
                        if($testerNode->nodeName=="w:color" && $testerNode->attributes->getNamedItem("val")->nodeValue!="000000"){
                            $style .= "color:#".$testerNode->attributes->getNamedItem("val")->nodeValue.";";
                        }
                    }
                }

                $ret .="<p class='".$class."' style='".$style."'>".$entry->nodeValue."</p>";
            }

            $ret .= "</div>";
        }
        return $ret;
    }

    /**
    * parse docx xml content and return text nodes
    */
    private function  _getTextNodes(){
        $ret = false;

        $zip = new ZipArchive;
        if (true === $zip->open($this->fullpath)) {
            if (($index = $zip->locateName($this->dataFile)) !== false) { // If done, search for the data file in the archive
                $data = $zip->getFromIndex($index); // If found, read it to the string
                $zip->close();

                // Load XML from a string, Skip errors and warnings
                $xml = DOMDocument::loadXML($data, LIBXML_NOENT | LIBXML_XINCLUDE | LIBXML_NOERROR | LIBXML_NOWARNING);
                $xpath = new DomXPath($xml);
                $textNodes = $xpath->query("/w:document/w:body/w:p");
                
                $ret = $textNodes;
            }
        }

        return $ret;
    }

}    

