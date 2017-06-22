<?php

namespace OLEReader;

class OLEHeader
{
    const OLE_IDENTIFIER = "\xd0\xcf\x11\xe0\xa1\xb1\x1a\xe1";

    // header offsets
    const SIGNATURE_OFFSET = 0x00;
    const CLSID_OFFSET = 0x08;
    const MINOR_VERSION_OFFSET = 0x18;
    const DLL_VERSION_OFFSET = 0x1A;
    const BYTE_ORDER_OFFSET = 0x1C;
    const SECTOR_SHIFT_OFFSET = 0x1E;
    const MINI_SECTOR_SHIFT_OFFSET = 0x20;
    const RESERVED_OFFSET = 0x22;
    const RESERVED1_OFFSET = 0x24;
    const SECT_DIR_NUM_OFFSET = 0x28;
    const SECT_FAT_NUM_OFFSET = 0x2C;
    const FIRST_DIR_SECT_START_OFFSET = 0x30;
    const TRANSACTION_SIGNATURE_OFFSET = 0x34;
    const MINI_SECTOR_CUTOFF_OFFSET = 0x38;
    const FIRST_MINI_FAT_SECT_OFFSET = 0x3c;
    const MINI_FAT_SECT_NUMBER_OFFSET = 0x40;
    const FIRST_DIFAT_SECTOR_OFFSET = 0x44;
    const DIFAT_SECTOR_NUMBER_OFFSET = 0x48;
    const FAT_SECTORS_OFFSET = 0x4c;

    private $signature;
    private $clsid;
    private $minorVersion;
    private $dllVersion;
    private $byteOrder;
    private $sectorSize;
    private $miniSectorSize;
    private $numDirSectors;
    private $numFatSectors;
    private $firstDirectorySector;
    private $transactionSignatureNumber;
    private $miniStreamCutoffSize;
    private $firstMiniFatSector;
    private $numMiniFatSectors;
    private $firstDifatSector;
    private $numDifatSectors;
    private $firstFatSectors;

    /**
     * OLEHeader constructor.
     */
    public function __construct($data)
    {
        if (strlen($data) < 512) {
            throw new \InvalidArgumentException("Data must be at least 512 bytes long");
        }

        // Check OLE identifier
        $this->signature = substr($data, self::SIGNATURE_OFFSET, 8);
        if ($this->signature != self::OLE_IDENTIFIER) {
            throw new \Exception("The data is not recognised as an OLE file");
        }

        $this->clsid = substr($data, self::CLSID_OFFSET, 16);

        $this->minorVersion   = Utils::getUint2d($data, self::MINOR_VERSION_OFFSET);
        $this->dllVersion     = Utils::getUint2d($data, self::DLL_VERSION_OFFSET);
        $this->byteOrder      = Utils::getUint2d($data, self::BYTE_ORDER_OFFSET);
        $this->sectorSize     = 1 << Utils::getUint2d($data, self::SECTOR_SHIFT_OFFSET);
        $this->miniSectorSize = 1 << Utils::getUint2d($data, self::MINI_SECTOR_SHIFT_OFFSET);
        //Utils::getInt2d($data, self::RESERVED_OFFSET);
        //Utils::getInt4d($data, self::RESERVED_OFFSET);
        $this->numDirSectors              = Utils::getInt4d($data, self::SECT_DIR_NUM_OFFSET);
        $this->numFatSectors              = Utils::getInt4d($data, self::SECT_FAT_NUM_OFFSET);
        $this->firstDirectorySector       = Utils::getInt4d($data, self::FIRST_DIR_SECT_START_OFFSET);
        $this->transactionSignatureNumber = Utils::getInt4d($data, self::TRANSACTION_SIGNATURE_OFFSET);
        $this->miniStreamCutoffSize       = Utils::getInt4d($data, self::MINI_SECTOR_CUTOFF_OFFSET);
        $this->firstMiniFatSector         = Utils::getInt4d($data, self::FIRST_MINI_FAT_SECT_OFFSET);
        $this->numMiniFatSectors          = Utils::getInt4d($data, self::MINI_FAT_SECT_NUMBER_OFFSET);

        $this->firstDifatSector = Utils::getInt4d($data, self::FIRST_DIFAT_SECTOR_OFFSET);
        $this->numDifatSectors  = Utils::getInt4d($data, self::DIFAT_SECTOR_NUMBER_OFFSET);

        $this->firstFatSectors = Utils::getInt4dArray($data, self::FAT_SECTORS_OFFSET, 109);
    }

    public function __toString()
    {
        return "";
    }

    /**
     * @return bool|string
     */
    public function getSignature()
    {
        return $this->signature;
    }

    /**
     * @return bool|string
     */
    public function getClsid()
    {
        return $this->clsid;
    }

    /**
     * @return int
     */
    public function getMinorVersion()
    {
        return $this->minorVersion;
    }

    /**
     * @return int
     */
    public function getDllVersion()
    {
        return $this->dllVersion;
    }

    /**
     * @return int
     */
    public function getByteOrder()
    {
        return $this->byteOrder;
    }

    /**
     * @return int
     */
    public function getSectorSize()
    {
        return $this->sectorSize;
    }

    /**
     * @return int
     */
    public function getMiniSectorSize()
    {
        return $this->miniSectorSize;
    }

    /**
     * @return int
     */
    public function getNumDirSectors()
    {
        return $this->numDirSectors;
    }

    /**
     * @return int
     */
    public function getNumFatSectors()
    {
        return $this->numFatSectors;
    }

    /**
     * @return int
     */
    public function getFirstDirectorySector()
    {
        return $this->firstDirectorySector;
    }

    /**
     * @return int
     */
    public function getTransactionSignatureNumber()
    {
        return $this->transactionSignatureNumber;
    }

    /**
     * @return int
     */
    public function getMiniStreamCutoffSize()
    {
        return $this->miniStreamCutoffSize;
    }

    /**
     * @return int
     */
    public function getFirstMiniFatSector()
    {
        return $this->firstMiniFatSector;
    }

    /**
     * @return int
     */
    public function getNumMiniFatSectors()
    {
        return $this->numMiniFatSectors;
    }

    /**
     * @return int
     */
    public function getFirstDifatSector()
    {
        return $this->firstDifatSector;
    }

    /**
     * @return int
     */
    public function getNumDifatSectors()
    {
        return $this->numDifatSectors;
    }

    /**
     * @return array
     */
    public function getFirstFatSectors()
    {
        return $this->firstFatSectors;
    }
}