<?php

namespace App\Utils;

use PhpOffice\PhpWord\Escaper\RegExp;
use PhpOffice\PhpWord\Escaper\Xml;
use PhpOffice\PhpWord\Settings;
use PhpOffice\PhpWord\TemplateProcessor;

class TemplateProcessor2 extends TemplateProcessor {
        /**
     * @param mixed $search
     * @param mixed $replace
     */
    public function replaceBookmark($search, $replace)
    {
        if (is_array($replace)) {
            foreach ($replace as &$item) {
                $item = self::ensureUtf8Encoded($item);
            }
        } else {
            $replace = self::ensureUtf8Encoded($replace);
        }

        if (Settings::isOutputEscapingEnabled()) {
            $xmlEscaper = new Xml();
            $replace = $xmlEscaper->escape($replace);
        }

        foreach ($this->tempDocumentHeaders as $index => $xml) {
            $xml = $this->setBookmarkForPart($search, $replace, $xml);
        }
        $this->tempDocumentMainPart = $this->setBookmarkForPart($search, $replace, $this->tempDocumentMainPart);
        foreach ($this->tempDocumentFooters as $index => $xml) {
            $xml = $this->setBookmarkForPart($search, $replace, $xml);
        }

    }


   /**
     * Find and replace bookmarks in the given XML section.
     *
     * @param mixed $search
     * @param mixed $replace
     * @param string $documentPartXML
     *
     * @return string
     */
    protected function setBookmarkForPart($search, $replace, $documentPartXML)
    {
        $regExpEscaper = new RegExp();
        $pattern = '~<w:bookmarkStart\s+w:id="(\d*)"\s+w:name="'.$search.'"\s*\/>()~mU';
        $searchstatus = preg_match($pattern, $documentPartXML, $matches, PREG_OFFSET_CAPTURE);
        if($searchstatus){
            $startbookmark = $matches[2][1];
            $pattern = '~(<w:bookmarkEnd\s+w:id="'.$matches[1][0].'"\s*\/>)~mU';
            $searchstatus = preg_match($pattern, $documentPartXML, $matches, PREG_OFFSET_CAPTURE, $startbookmark);
            if($searchstatus){
                $endbookmark = $matches[1][1];
                $count = 0;
                $startpos = $startbookmark;
                $pattern = '~(<w:t[\s\S]*>)([\s\S]*)(<\/w:t>)~mU';
                do{
                    $searchstatus = preg_match($pattern, $documentPartXML, $matches, PREG_OFFSET_CAPTURE, $startpos);
                    if($searchstatus){
                        if($count == 0){
                            $startpos = $matches[2][1];
                            $endpos = $matches[3][1];
                        }else{
                            $startpos = $matches[1][1];
                            $endpos = $matches[3][1] + 6;
                        }
                        if($endpos > $endbookmark){
                            break;
                        }

                        $documentPartXML = substr($documentPartXML, 0, $startpos) . ($count == 0 ? $replace : '') . substr($documentPartXML, $endpos);
                        $endbookmark = $endbookmark - ($endpos - $startpos);

                        $count ++;
                    }

                }while($searchstatus);

            }
        }

        return $documentPartXML;
    }
}

