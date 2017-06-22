<?php

namespace OLEReader;

class OLEReader
{
    const UNKNOWN_SIZE = -1;
    const NOSTREAM = -1;

    // Size of a directory entry always = 128 bytes
    const DIRECTORY_ENTRY_SIZE = 128;

    // Minimum size of a standard stream = 4096 bytes, streams smaller than this are stored as short streams
    const SMALL_BLOCK_THRESHOLD = 0x1000;

    const MAXREGSECT = 0xFFFFFFFA;
    const DIFSECT = 0xFFFFFFFC;
    const FATSECT = 0xFFFFFFFD;
    const END_OF_CHAIN = -2;
    const FREE_SECTOR = -1;

    const STGTY_EMPTY = 0;// empty directory entry
    const STGTY_STORAGE = 1;// element is a storage object
    const STGTY_STREAM = 2;// element is a stream object
    const STGTY_LOCKBYTES = 3;// element is an ILockBytes object
    const STGTY_PROPERTY = 4;// element is an IPropertyStorage object;//
    const STGTY_ROOT = 5;// element is a root storage

    const VT_EMPTY = 0;
    const VT_NULL = 1;
    const VT_I2 = 2;
    const VT_I4 = 3;
    const VT_R4 = 4;
    const VT_R8 = 5;
    const VT_CY = 6;
    const VT_DATE = 7;
    const VT_BSTR = 8;
    const VT_DISPATCH = 9;
    const VT_ERROR = 10;
    const VT_BOOL = 11;
    const VT_VARIANT = 12;
    const VT_UNKNOWN = 13;
    const VT_DECIMAL = 14;
    const VT_I1 = 16;
    const VT_UI1 = 17;
    const VT_UI2 = 18;
    const VT_UI4 = 19;
    const VT_I8 = 20;
    const VT_UI8 = 21;
    const VT_INT = 22;
    const VT_UINT = 23;
    const VT_VOID = 24;
    const VT_HRESULT = 25;
    const VT_PTR = 26;
    const VT_SAFEARRAY = 27;
    const VT_CARRAY = 28;
    const VT_USERDEFINED = 29;
    const VT_LPSTR = 30;
    const VT_LPWSTR = 31;
    const VT_FILETIME = 64;
    const VT_BLOB = 65;
    const VT_STREAM = 66;
    const VT_STORAGE = 67;
    const VT_STREAMED_OBJECT = 68;
    const VT_STORED_OBJECT = 69;
    const VT_BLOB_OBJECT = 70;
    const VT_CF = 71;
    const VT_CLSID = 72;
    const VT_VECTOR = 0x1000;

    private $data = '';

    public $workbook = null;
    public $summaryInformation = null;
    public $documentSummaryInformation = null;
    public $filename;

    /**
     * @var OLEHeader
     */
    private $header;
    public $isFileRead = false;
    private $minifatLoaded;
    private $ministream;
    private $directoryEntires;
    private $directoryStream;

    /**
     * @var OLEDirectory
     */
    private $rootDirectory;

    private $fatSectors = [];
    private $minifatSectors = [];

    /**
     * @param $pFilename string Filename
     *
     * @throws \Exception
     */
    public function __construct($pFilename)
    {
        $this->filename = $pFilename;
        $this->read();
    }

    public static function registerTypes()
    {
        BaseFile::registerEntryType(OLEReader::STGTY_ROOT, OLEDirectory::class);
        BaseFile::registerEntryType(OLEReader::STGTY_STORAGE, OLEDirectory::class);
        BaseFile::registerEntryType(OLEReader::STGTY_EMPTY, OLEFile::class);
        BaseFile::registerEntryType(OLEReader::STGTY_STREAM, OLEFile::class);
        BaseFile::registerEntryType(OLEReader::STGTY_PROPERTY, OLEPropertyStorage::class);
    }

    public function getHeader()
    {
        return $this->header;
    }

    /**
     * Read the file.
     */
    private function read()
    {
        // Get the file data
        $this->data = file_get_contents($this->filename);

        try {
            $this->header = new OLEHeader($this->data);
        } catch (\InvalidArgumentException $e) {
            throw new \InvalidArgumentException("File {$this->filename} is not a valid OLE file");
        }

        $this->loadFat();
        $this->loadDirectory($this->header->getFirstDirectorySector());
        $this->isFileRead = true;
    }

    private function loadDirectory($directorySector)
    {
        $this->openDirectory($directorySector);

        $this->rootDirectory = $this->loadDirEntry(0);
    }

    public function getRootDirectory()
    {
        if (!$this->isFileRead) {
            $this->read();
        }

        return $this->rootDirectory;
    }

    public function getFile($string)
    {
        if (!$this->isFileRead) {
            $this->read();
        }

        $path = explode("/", $string);
        /** @var OLEDirectory $dir */
        $dir = $this->rootDirectory;
        foreach ($path as $part) {
            $dir = $dir->getChild($part);
        }

        return $dir;
    }

    public function readStream($startSector, $size = self::UNKNOWN_SIZE, $forceFat = false)
    {
        if ($size !== self::UNKNOWN_SIZE && $size < $this->header->getMiniStreamCutoffSize() && !$forceFat) {
            if (!$this->minifatLoaded) {
                $this->loadMinifat();

                $miniStreamSize = $this->rootDirectory->getSize();
                $this->ministream = $this->readStream($this->rootDirectory->getFirstSector(), $miniStreamSize, true);
            }

            return self::readStreamInternal($this->ministream, $startSector,0, $size, $this->minifatSectors, $this->header->getMiniSectorSize());
        } else {
            return self::readStreamInternal($this->data, $startSector,$this->header->getSectorSize(), $size, $this->fatSectors, $this->header->getSectorSize());
        }
    }

    public function loadDirEntry($sid)
    {
        if ($sid < 0) {
            throw new \OutOfBoundsException("\$sid must be >= 0, currently {$sid}");
        }

        if (!isset($this->directoryEntires[ $sid ])) {
            $entry = substr($this->directoryStream, $sid * self::DIRECTORY_ENTRY_SIZE, self::DIRECTORY_ENTRY_SIZE);

            $this->directoryEntires[ $sid ] = BaseFile::create($this, $sid, $entry);
        }

        return $this->directoryEntires[ $sid ];
    }

    private function openDirectory($directorySector)
    {
        $this->directoryStream = $this->readStream($directorySector, self::UNKNOWN_SIZE, false);
    }

    private static function readStreamInternal($data, $startSector, $offset, $size, $fatSectors, $sectorSize)
    {
        //max size
        if ($size == self::UNKNOWN_SIZE) {
            $size          = count($fatSectors) * $sectorSize;
            $isUnknownSize = true;
        } else {
            $isUnknownSize = false;
        }
        $nbSectors = floor(($size + $sectorSize - 1) / $sectorSize);

        if ($nbSectors > count($fatSectors)) {
            throw new \InvalidArgumentException("Stream too large");
        }

        $result = '';
        if ($size === 0 && $startSector != self::END_OF_CHAIN) {
            throw new \InvalidArgumentException("Incorrect sector index for empty stream");
        }

        $sect   = $startSector;
        for ($i = 0; $i < $nbSectors; $i++) {
            if ($sect === self::END_OF_CHAIN) {
                break;
            }

            if ($sect < 0 || $sect > count($fatSectors)) {
                throw new \InvalidArgumentException("Incorrect sector index {$sect}");
            }

            $sectorData = substr($data, $offset + $sectorSize * $sect, $sectorSize);
            $result     .= $sectorData;
            $sect       = $fatSectors[ $sect ];
        }

        return substr($result, 0, $size);
    }

    private function loadMinifat()
    {
        $this->minifatLoaded = true;

        $sectorSize     = $this->header->getSectorSize();
        $streamSize     = $this->header->getNumMiniFatSectors() * $sectorSize;
        $miniSectorSize = $this->header->getMiniSectorSize();
        $nbMiniSectors  = floor(($this->rootDirectory->getSize() + $miniSectorSize - 1) / $miniSectorSize);
        //1723
        $usedSize = $nbMiniSectors * 4;
        if ($usedSize > $streamSize) {
            throw new \UnexpectedValueException("OLE MiniStream is larger than MiniFAT");
        }

        $s = $this->readStream($this->header->getFirstMiniFatSector(), $streamSize, true);

        $this->minifatSectors = array_slice(Utils::getInt4dArray($s), 0, $nbMiniSectors);
    }

    private function loadFat()
    {
        foreach ($this->header->getFirstFatSectors() as $s) {
            if ($s == OLEReader::END_OF_CHAIN || $s == OLEReader::FREE_SECTOR) {
                break;
            }

            $offset = ($s + 1) * $this->header->getSectorSize();
            $sector = substr($this->data, $offset, $this->header->getSectorSize());

            foreach (Utils::getInt4dArray($sector, 0) as $sc) {
                $this->fatSectors[] = $sc;
            }
        }

        $maxSectorCount = floor((strlen($this->data) + $this->header->getSectorSize() - 1) / $this->header->getSectorSize()) - 1;
        if ($this->header->getNumDifatSectors() > 0) {
            if ($this->header->getNumFatSectors() < 109) {
                throw new \UnexpectedValueException("DIFAT is used but number of sectors is < 109");
            }

            //TODO
        }

        if (count($this->fatSectors) > $maxSectorCount) {
            $this->fatSectors = array_slice($this->fatSectors, 0, $maxSectorCount);
        }
    }
}