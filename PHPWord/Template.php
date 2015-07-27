<?php
/**
 * PHPWord
 *
 * Copyright (c) 2011 PHPWord
 *
 * This library is free software; you can redistribute it and/or
 * modify it under the terms of the GNU Lesser General Public
 * License as published by the Free Software Foundation; either
 * version 2.1 of the License, or (at your option) any later version.
 *
 * This library is distributed in the hope that it will be useful,
 * but WITHOUT ANY WARRANTY; without even the implied warranty of
 * MERCHANTABILITY or FITNESS FOR A PARTICULAR PURPOSE.  See the GNU
 * Lesser General Public License for more details.
 *
 * You should have received a copy of the GNU Lesser General Public
 * License along with this library; if not, write to the Free Software
 * Foundation, Inc., 51 Franklin Street, Fifth Floor, Boston, MA  02110-1301  USA
 *
 * @category   PHPWord
 * @package    PHPWord
 * @copyright  Copyright (c) 010 PHPWord
 * @license    http://www.gnu.org/licenses/old-licenses/lgpl-2.1.txt    LGPL
 * @version    Beta 0.6.3, 08.07.2011
 */


/**
 * PHPWord_DocumentProperties
 *
 * @category   PHPWord
 * @package    PHPWord
 * @copyright  Copyright (c) 2009 - 2011 PHPWord (http://www.codeplex.com/PHPWord)
 */
class PHPWord_Template implements PHPWord_Writer_IWriter {
    
    /**
     * ZipArchive
     * 
     * @var ZipArchive
     */
    private $_objZip;
    
    /**
     * Temporary Filename
     * 
     * @var string
     */
    private $_tempFileName;
    
    /**
     * Document XML
     * 
     * @var string
     */
    private $_documentXML;

    /**
     * Element Collection
     * 
     * @var array
     */
    private $_elementCollection = array();
    private $_imageTypes = array();
    private $_objectTypes = array();
    private $_useDiskCaching = false;

    private $_documentRelsXML;
    private $_relationshipsNum = 0;
    private $_firstTempImgAdd = false;

    /**
     * Create a new Template Object
     * 
     * @param string $strFilename
     */
    public function __construct($strFilename) {
        $path = dirname($strFilename);
        $this->_tempFileName = $path.DIRECTORY_SEPARATOR.time(). '_' .rand(100, 999) . '.docx';
        
        copy($strFilename, $this->_tempFileName); // Copy the source File to the temp File

        $this->_objZip = new ZipArchive();
        $this->_objZip->open($this->_tempFileName);
        
        $this->_documentXML = $this->_objZip->getFromName('word/document.xml');

        $this->_documentRelsXML = $this->_objZip->getFromName('word/_rels/document.xml.rels');
    }
    
    /**
     * Set a Template value
     * 
     * @param mixed $search
     * @param mixed $replace
     */
    public function setValue($search, $replace) {
        if(substr($search, 0, 2) !== '${' && substr($search, -1) !== '}') {
            $search = '${'.$search.'}';
        }
        
        // if(!is_array($replace)) {
        //     $replace = utf8_encode($replace);
        // }
        
        $this->_documentXML = str_replace($search, $replace, $this->_documentXML);
    }
    /**
     * preg SET a Template value
     * 
     * @param mixed $pattern
     * @param mixed $replace
     */
    public function pregSetValue($pattern, $replace, $limit = -1) {
        // if(!is_array($replace)) {
        //     $replace = utf8_encode($replace);
        // }

        $this->_documentXML = preg_replace($pattern, $replace, $this->_documentXML, $limit);
    }
    /**
     * preg GET a _relationshipsNum
     * 
     * @param mixed $pattern
     * @param mixed $replace
     */
    public function getRelationshipsNum() {
        // if(!is_array($replace)) {
        //     $replace = utf8_encode($replace);
        // }
        if (!$this->_firstTempImgAdd) {
            $this->_relationshipsNum += substr_count($this->_documentRelsXML, '<Relationship ');
        }
        return $this->_relationshipsNum;
    }    
    /**
     * Save Template
     * 
     * @param string $strFilename
     */
    public function save($strFilename = null) {
        if(file_exists($strFilename)) {
            unlink($strFilename);
        }
        $this->_objZip->addFromString('word/document.xml', $this->_documentXML);  

        if (count($this->_elementCollection)) {
            $sectionElements = array();
            $_secElements = PHPWord_Media::getSectionMediaElements();
            foreach($_secElements as $element) { // loop through section media elements
                if($element['type'] != 'hyperlink') {
                    $this->_addFileToPackage($this->_objZip, $element);
                }
                $sectionElements[] = $element;
            }

            /*************************************************/
            $documentrels = new PHPWord_Writer_Word2007_DocumentRels();
            $documentrels->setParentWriter($this);

            $pattern = '/(.*)(<\/Relationships>)/s';
            $j = 0;
            $replace = '';
            for ($i = count($this->_elementCollection); $i > 0; $i--) {
                $j++;
                $replace .= '<Relationship Id="rId' . ($this->_relationshipsNum - ($i - 1)) .
                '" Type="http://schemas.openxmlformats.org/officeDocument/2006/relationships/image"' .
                ' Target="media/section_image' . $j . '.png"/>';;
            }
            $_documentRelsXML = trim(preg_replace($pattern, '${1}' . $replace . '${2}', $this->_documentRelsXML));
            $this->_objZip->addFromString('word/_rels/document.xml.rels', $_documentRelsXML);
            /*************************************************/
        }

        // Close zip file
        if($this->_objZip->close() === false) {
            throw new Exception('Could not close zip file.');
        }
        
        rename($this->_tempFileName, $strFilename);
    }
    /**
     * Get SomeOne Of Table
     * 
     * @param string $cutoff
     */
    public function getSomeTable($cutoff = 'tbl') {
        $someTemp = '';
        if (strlen($cutoff)) {

            $patternXml = '/(^<\?)(.*?)(\?>)/';
            $notxml = trim(preg_replace($patternXml, '', $this->_documentXML, 1));

            $patternDocument = '/(<w:document)(.*?)(wordml\">)(.*)/';
            $notdoc0 = trim(preg_replace($patternDocument, '${4}', $notxml));

            $patternBody = '/(<w:body>)(.*)/';
            $notbody0 = trim(preg_replace($patternBody, '${2}', $notdoc0));

            $patternDocument = '/(.*)(<\/w:document>)/';
            $notdoc1 = trim(preg_replace($patternDocument, '${1}', $notbody0));

            $patternBody = '/(.*)(<\/w:body>)/';
            $notbody1 = trim(preg_replace($patternBody, '${1}', $notdoc1));

            $patternBody = '/(.*)(<w:tbl>.*<\/w:tbl>)(.*)/';
            $notbody = trim(preg_replace($patternBody, '${2}', $notbody1));
            
            $someTemp = $notbody;
        }
        if (file_exists($this->_tempFileName)) {
            unlink($this->_tempFileName);
        }
        return $someTemp;
    }

    /**
     * Get SomeOne Of Text
     * 
     * @param string $cutoff
     */
    public function getSomeText($cutoff = 'text') {
        $someTemp = '';
        if (strlen($cutoff)) {

            $patternXml = '/(^<\?)(.*?)(\?>)/';
            $notxml = trim(preg_replace($patternXml, '', $this->_documentXML, 1));

            $patternDocument = '/(<w:document)(.*?)(\">)(.*)/';
            $notdoc0 = trim(preg_replace($patternDocument, '${4}', $notxml));

            $patternBody = '/(<w:body>)(.*)/';
            $notbody0 = trim(preg_replace($patternBody, '${2}', $notdoc0));

            $patternDocument = '/(.*)(<\/w:document>)/';
            $notdoc1 = trim(preg_replace($patternDocument, '${1}', $notbody0));

            $patternBody = '/(.*)(<\/w:body>)/';
            $notbody = trim(preg_replace($patternBody, '${1}', $notdoc1));

            $someTemp = $notbody;
        }
        if (file_exists($this->_tempFileName)) {
            unlink($this->_tempFileName);
        }
        return $someTemp;
    }

    /**
     * Get SomeOne Of Image
     * 
     * @param string $cutoff
     */
    public function getSomeImage($cutoff = 'p') {
        $someTemp = '';
        if (strlen($cutoff)) {

            $patternXml = '/(^<\?)(.*?)(\?>)/';
            $notxml = trim(preg_replace($patternXml, '', $this->_documentXML, 1));

            $patternDocument = '/(<w:document)(.*?)(wordml\">)(.*)/';
            $notdoc0 = trim(preg_replace($patternDocument, '${4}', $notxml));

            $patternBody = '/(<w:body>)(.*)/';
            $notbody0 = trim(preg_replace($patternBody, '${2}', $notdoc0));

            $patternDocument = '/(.*)(<\/w:document>)/';
            $notdoc1 = trim(preg_replace($patternDocument, '${1}', $notbody0));

            $patternBody = '/(.*)(<\/w:body>)/';
            $notbody1 = trim(preg_replace($patternBody, '${1}', $notdoc1));

            $patternBody = '/(.*)(<w:sectPr.*)/s';
            $notbody = trim(preg_replace($patternBody, '${1}', $notbody1));

            $someTemp = $notbody;
        }
        if (file_exists($this->_tempFileName)) {
            unlink($this->_tempFileName);
        }
        return $someTemp;
    }

    /**
     * Add a Image Element
     * 
     * @param string $src
     * @param mixed $style
     * @return PHPWord_Section_Image
     */
    public function addImage($src, $style = null) {
        $image = new PHPWord_Section_Image($src, $style);
        if(!is_null($image->getSource())) {
            /***********************************/
            if (!$this->_firstTempImgAdd) {
                $this->_firstTempImgAdd = true;
            }
            $this->_relationshipsNum++;
            /***********************************/
            $rID = PHPWord_Media::addSectionMediaElement($src, 'image');
            $image->setRelationId($rID);

            $this->_elementCollection[] = $image;

            return $image;
        } else {
            trigger_error('Source does not exist or unsupported image type.');
        }
    }


    private function _chkContentTypes($src) {
        $srcInfo   = pathinfo($src);
        $extension = strtolower($srcInfo['extension']);
        if(substr($extension, 0, 3) == 'php') {
            $extension = 'php';
        }
        $_supportedImageTypes = array('jpg', 'jpeg', 'gif', 'png', 'bmp', 'tif', 'tiff', 'php');

        if(in_array($extension, $_supportedImageTypes)) {
            $imagedata = getimagesize($src);
            $imagetype = image_type_to_mime_type($imagedata[2]);
            $imageext = image_type_to_extension($imagedata[2]);
            $imageext = str_replace('.', '', $imageext);
            if($imageext == 'jpeg') $imageext = 'jpg';

            if(!in_array($imagetype, $this->_imageTypes)) {
                $this->_imageTypes[$imageext] = $imagetype;
            }
        } else {
            if(!in_array($extension, $this->_objectTypes)) {
                $this->_objectTypes[] = $extension;
            }
        }
    }

    private function _addFileToPackage($objZip, $element) {
        if(isset($element['isMemImage']) && $element['isMemImage']) {
            $image = call_user_func($element['createfunction'], $element['source']);
            ob_start();
            call_user_func($element['imagefunction'], $image);
            $imageContents = ob_get_contents();
            ob_end_clean();
            $objZip->addFromString('word/'.$element['target'], $imageContents);
            imagedestroy($image);

            $this->_chkContentTypes($element['source']);
        } else {
            $objZip->addFile($element['source'], 'word/'.$element['target']);
            $this->_chkContentTypes($element['source']);
        }
    }

    public function getUseDiskCaching() {
        return $this->_useDiskCaching;
    }

    ///////////////////append by wade 2015-02-27///////////////////////
    public function delLineBreak($xml) {
        $lineBreakPattern = '/\s(?=\s)/';
        return trim(preg_replace($lineBreakPattern, '', $xml));
    }
    public function getBody() {
        $xmlVersionPattern = '/(^<\?)(.*?)(\?>)/';
        $xml = trim(preg_replace($xmlVersionPattern, '', $this->_documentXML, 1));

        $xml = $this->delLineBreak($xml);

        $bodyPattern = '/(.*<w:body>)(.*?)(<\/w:body>.*)/';
        $body = trim(preg_replace($bodyPattern, '${2}', $xml));

        if (file_exists($this->_tempFileName)) {
            unlink($this->_tempFileName);
        }
        return $body;
    }

    // public function getSomeCell($flag = '') {
    //     $cellPattern = '/(.*)(<w:tc>.*?' . $flag . '.*?<\/w:tc>)(.*)/';

    //     return trim(preg_replace($cellPattern, '${2}', $this->_documentXML));
    // }

    public function addCells($flag, $col) {
        $this->_documentXML = $this->delLineBreak($this->_documentXML);

        $cellPattern = '/(.*)(<w:tc>.*?)(' . $flag . ')(.*?<\/w:r>)' .
            '(<w:bookmarkStart.*?\/><w:bookmarkEnd.*?\/>|)(.*?<\/w:tc>)(.*)/';

        $replace = '${2}${3}${4}${5}${6}';
        if (count($col)) {
            $replace = '';
            for ($i = 0; $i < count($col); $i++) {
                if ($i === count($col) - 1) {
                    $replace .= '${2}' . $col[$i] . '${4}${5}${6}';
                } else {
                    $replace .= '${2}' . $col[$i] . '${4}${6}';
                }
            }
        }
        $replace = '${1}' . $replace . '${7}';

        $this->_documentXML = preg_replace($cellPattern, $replace, $this->_documentXML);
    }
    // <w:p...>...</w:p>
    public function overallSubstitutionOfFlagP($flag, $replace = '') {
        $this->_documentXML = $this->delLineBreak($this->_documentXML);

        $pPattern = '/(.*)(<w:p .*?' . $flag . '.*?<\/w:r>)' .
            '(<w:bookmarkStart.*?\/><w:bookmarkEnd.*?\/>|)(.*?<\/w:p>)(.*)/';
        $replace = '${1}' . $replace . '${5}';

        $this->_documentXML = trim(preg_replace($pPattern, $replace, $this->_documentXML));
    }

    // public function getGrid($tbl = '') {
    //     $gridColPattern = '/.*<w:tblGrid>.*(<w:gridCol w:w=\"[0-9]+\"\/>)<\/w:tblGrid>.*/';

    //     return trim(preg_replace($gridColPattern, '${1}', $tbl));
    // }

    public function addGrid($col) {
        $this->_documentXML = $this->delLineBreak($this->_documentXML);

        $gridColPattern = '/(.*<w:tblGrid>.*)(<w:gridCol w:w=\"[0-9]+\"\/>)(<\/w:tblGrid>.*)/';
        $replace = '${2}';
        $width = 0;
        if (count($col)) {
            $replace = '';
            for ($i = 0; $i < count($col); $i++) {
                $replace .= '${2}';
            }
            $gridP = '/<w:tblGrid>.*<w:gridCol w:w=\"([0-9]+)\"\/><\/w:tblGrid>/';
            preg_match($gridP, $this->_documentXML, $matches, PREG_OFFSET_CAPTURE);
            if (isset($matches[1][0])) {
                $width = (int)$matches[1][0] * (count($col) - 1);
                unset($matches);
            }
        }
        $replace = '${1}' . $replace . '${3}';
        $this->_documentXML = trim(preg_replace($gridColPattern, $replace, $this->_documentXML));

        $tblWP = '/<w:tblW w:w=\"([0-9]+)/';
        preg_match($tblWP, $this->_documentXML, $matches, PREG_OFFSET_CAPTURE);
        if (isset($matches[1][0])) {
            $width += (int)$matches[1][0];
            unset($matches);
        }

        $tblWpattern = '/(.*<w:tblW w:w=\")([0-9]+)(\".*)/';
        $this->_documentXML = trim(preg_replace($tblWpattern, '${1}' . $width . '${3}', $this->_documentXML));
    }

}
?>
