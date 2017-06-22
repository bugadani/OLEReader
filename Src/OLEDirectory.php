<?php

namespace OLEReader;

class OLEDirectory extends BaseFile
{
    protected $children = [];
    protected $childNameMap;

    protected function __construct(OLEReader $owner, $sid, $name, $entry)
    {
        parent::__construct($owner, $sid, $name, $entry);

        $isRoot           = $this->entryType == OLEReader::STGTY_ROOT;
        $isFirstDirectory = $sid === 0;

        if ($isRoot && !$isFirstDirectory) {
            throw new \InvalidArgumentException("Directory {$sid} must not be root");
        } else if (!$isRoot && $isFirstDirectory) {
            throw new \InvalidArgumentException("Directory {$sid} should be root, it is {$this->entryType}");
        }

        if ($this->child != OLEReader::NOSTREAM) {
            $this->appendChild($this->child);
        }
    }

    private function appendChild($sid)
    {
        if ($sid === null || $sid === OLEReader::NOSTREAM) {
            return;
        }

        if ($sid < 0) {
            throw new \InvalidArgumentException("Invalid directory id: {$sid}");
        }

        $child = $this->owner->loadDirEntry($sid);

        $this->appendChild($child->left);
        $this->children[] = $child;
        if (isset($this->childNameMap[ strtolower($child->name) ])) {
            throw new \InvalidArgumentException("Duplicate filename '{$child->name}'");
        }
        $this->childNameMap[ strtolower($child->name) ] = $child;

        $this->appendChild($child->right);
    }

    public function getChildren()
    {
        return $this->children;
    }

    /**
     * @param $name
     * @return int
     */
    public function getChild($name)
    {
        if (!isset($this->childNameMap[ strtolower($name) ])) {
            throw new \OutOfBoundsException("File not found: {$name}");
        }

        return $this->childNameMap[ strtolower($name) ];
    }

    public function getSize()
    {
        if ($this->entryType == OLEReader::STGTY_ROOT) {
            return parent::getSize();
        }

        return array_reduce($this->children, function ($sum, BaseFile $item) {
            return $sum + $item->getSize();
        }, 0);
    }
}