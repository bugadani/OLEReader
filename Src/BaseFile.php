<?php

namespace OLEReader;

class BaseFile
{
    private static $entryTypeMap = [];

    public static function registerEntryType($type, $class)
    {
        self::$entryTypeMap[ $type ] = $class;
    }

    public static function create(OLEReader $owner, $sid, $entry)
    {
        $nameRaw    = substr($entry, self::FILE_ENTRY_NAME_OFFSET, 64);
        $nameLength = Utils::getUint2d($entry, self::FILE_ENTRY_NAME_LENGTH_OFFSET);
        $entryType  = ord($entry[ self::FILE_ENTRY_TYPE_OFFSET ]);

        // Validate entry type
        //[OLEReader::STGTY_EMPTY, OLEReader::STGTY_STORAGE, OLEReader::STGTY_STREAM, OLEReader::STGTY_ROOT]
        if (!array_key_exists($entryType, self::$entryTypeMap)) {
            throw new \InvalidArgumentException("Invalid entry type for directory: {$entryType}");
        }

        // Set name
        if ($nameLength > 64) {
            $nameLength = 64;
        }

        $name = str_replace("\x00", '', substr($nameRaw, 0, $nameLength - 2));

        return new self::$entryTypeMap[ $entryType ]($owner, $sid, $name, $entry);
    }

    // property storage offsets (directory offsets)
    const FILE_ENTRY_NAME_OFFSET = 0;
    const FILE_ENTRY_NAME_LENGTH_OFFSET = 0x40;
    const FILE_ENTRY_TYPE_OFFSET = 0x42;
    const FILE_ENTRY_COLOR_OFFSET = 0x43;
    const FILE_ENTRY_SID_LEFT_OFFSET = 0x44;
    const FILE_ENTRY_SID_RIGHT_OFFSET = 0x48;
    const FILE_ENTRY_SID_CHILD_OFFSET = 0x4C;
    const FILE_ENTRY_CLSID_OFFSET = 0x50;
    const FILE_ENTRY_USER_FLAGS_OFFSET = 0x60;
    const FILE_ENTRY_CTIME_OFFSET = 0x64;
    const FILE_ENTRY_MTIME_OFFSET = 0x6C;
    const FILE_ENTRY_ISECT_START_OFFSET = 0x74;
    const FILE_ENTRY_SIZE_LOW_OFFSET = 0x78;
    const FILE_ENTRY_SIZE_HIGH_OFFSET = 0x7C;

    protected $size;
    protected $entryType;
    protected $isMinifat;
    protected $color;
    protected $left;
    protected $right;
    protected $child;
    protected $clsid;
    protected $flags;
    protected $ctime;
    protected $mtime;
    protected $firstSector;
    /**
     * @var OLEReader
     */
    protected $owner;
    /**
     * @var
     */
    protected $sid;

    /**
     * @var OLEDirectory
     */
    private $parent;

    protected function __construct(OLEReader $owner, $sid, $name, $entry)
    {
        $this->owner = $owner;
        $this->sid   = $sid;
        $this->name  = $name;

        $this->entryType   = ord($entry[ self::FILE_ENTRY_TYPE_OFFSET ]);
        $this->color       = ord($entry[ self::FILE_ENTRY_COLOR_OFFSET ]);
        $this->left        = Utils::getInt4d($entry, self::FILE_ENTRY_SID_LEFT_OFFSET);
        $this->right       = Utils::getInt4d($entry, self::FILE_ENTRY_SID_RIGHT_OFFSET);
        $this->child       = Utils::getInt4d($entry, self::FILE_ENTRY_SID_CHILD_OFFSET);
        $this->clsid       = substr($entry, self::FILE_ENTRY_CLSID_OFFSET, 16);
        $this->flags       = Utils::getUint4d($entry, self::FILE_ENTRY_USER_FLAGS_OFFSET);
        $this->ctime       = Utils::getTimestamp($entry, self::FILE_ENTRY_CTIME_OFFSET);
        $this->mtime       = Utils::getTimestamp($entry, self::FILE_ENTRY_MTIME_OFFSET);
        $this->firstSector = Utils::getUint4d($entry, self::FILE_ENTRY_ISECT_START_OFFSET);
        $this->lowSize     = Utils::getUint4d($entry, self::FILE_ENTRY_SIZE_LOW_OFFSET);
        $this->highSize    = Utils::getUint4d($entry, self::FILE_ENTRY_SIZE_HIGH_OFFSET);

        // Set size
        if ($owner->getHeader()->getSectorSize() == 512) {
            //TODO validate highSize
            $this->size = $this->lowSize;
        } else {
            $this->size = $this->lowSize + $this->highSize << 32;
        }

        if ($this->entryType == OLEReader::STGTY_ROOT || $this->entryType == OLEReader::STGTY_STREAM) {
            if ($this->size > 0) {
                $this->isMinifat = ($this->size < $owner->getHeader()->getMiniStreamCutoffSize());
            }
        }
    }

    /**
     * @return string
     */
    public function getName()
    {
        return $this->name;
    }

    protected function setParent(OLEDirectory $parent)
    {
        $this->parent = $parent;
    }

    /**
     * @return OLEDirectory
     */
    public function getParent()
    {
        return $this->parent;
    }

    public function getPath()
    {
        if (isset($this->parent)) {
            return $this->parent->getPath() . '/' . $this->getName();
        } else {
            return $this->getName();
        }
    }

    /**
     * @return int
     */
    public function getSize()
    {
        return $this->size;
    }

    /**
     * @return mixed
     */
    public function getFirstSector()
    {
        return $this->firstSector;
    }
}
