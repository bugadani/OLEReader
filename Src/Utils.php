<?php

namespace OLEReader;

class Utils
{
    /**
     * Read 2 bytes of data at specified position.
     *
     * @param string $data
     * @param int $pos
     * @return int
     * @throws \Exception
     */
    public static function getInt2d($data, $pos)
    {
        if ($pos + 2 > strlen($data)) {
            throw new \InvalidArgumentException("Read after end");
        }
        list(, $tmp) = unpack('s', substr($data, $pos, 2));

        return $tmp;
    }

    /**
     * Read 4 bytes of data at specified position.
     *
     * @param string $data
     * @param int $pos
     * @return int
     * @throws \Exception
     */
    public static function getInt4d($data, $pos)
    {
        if ($pos + 4 > strlen($data)) {
            throw new \InvalidArgumentException("Read after end");
        }
        list(, $tmp) = unpack('l', substr($data, $pos, 4));

        return $tmp;
    }

    public static function getInt4dArray($data, $offset = 0, $size = null)
    {
        if ($size === null) {
            $size = (strlen($data) - $offset) / 4;
        }
        $array = [];
        for ($i = 0; $i < $size; $i++) {
            $array[] = Utils::getInt4d($data, $offset + 4 * $i);
        }

        return $array;
    }

    public static function getUint2d($data, $pos)
    {
        if ($pos + 2 > strlen($data)) {
            throw new \InvalidArgumentException("Read after end");
        }
        list(, $tmp) = unpack('S', substr($data, $pos, 2));

        return $tmp;
    }

    public static function getUint4d($data, $pos)
    {
        if ($pos + 4 > strlen($data)) {
            throw new \InvalidArgumentException("Read after end");
        }
        list(, $tmp) = unpack('L', substr($data, $pos, 4));

        return $tmp;
    }

    public static function getTimestamp($data, $pos)
    {
        //TODO
        return 0;
    }
}