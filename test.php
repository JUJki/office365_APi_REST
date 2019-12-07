<?php
error_reporting(-1);
ini_set('display_errors', 1);
require __DIR__ . '/vendor/autoload.php';
include 'office365Interface.php';
include 'UsageExample.php';

$office365Interface = new office365Interface();
$accessToken = $office365Interface->getAccessTokenByCredential();
$useage = new UsageExample();
$useage->run($accessToken);