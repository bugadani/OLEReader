<?php

namespace OLEReader;

class OLEFile extends BaseFile
{
    private $contents;
    private $contentsRead = false;

    public function getContents()
    {
        if (!$this->contentsRead) {
            $this->readContents();
        }

        return $this->contents;
    }

    private function readContents()
    {
        $this->contentsRead = true;
        $this->contents = $this->owner->readStream($this->firstSector, $this->size);
    }
}