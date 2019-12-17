<?php

const OF_ERROR_400 = 4000;
const OF_ERROR_401 = 4010;
const OF_ERROR_403 = 4030;
const OF_ERROR_404 = 4040;
const OF_ERROR = 5000;

class CustomException extends Exception
{

  public function __construct($message, $code, Throwable $previous = null)
  {
    parent::__construct($message, $code, $previous);
  }

}