<?php

namespace OLEReader;

class Stream
{
    private $data;
    private $pos;
    private $length;

    public function __construct($data)
    {
        $this->data   = $data;
        $this->length = strlen($data);
    }

    public function read($length)
    {
        if ($this->pos + $length > $this->length) {
            throw new \InvalidArgumentException("Cannot read {$length} bytes");
        }
        $data      = substr($this->data, $this->pos, $length);
        $this->pos += $length;

        return $data;
    }

    public function __toString()
    {
        return $this->data;
    }
}