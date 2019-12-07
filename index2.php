<?php

use Dotenv\Dotenv;

require __DIR__ . '/vendor/autoload.php';
$dotenv = Dotenv::createImmutable(__DIR__);
$dotenv->load();

error_reporting(($_ENV['ENVIRONEMENT'] === 'development') ? -1 : 0);
ini_set('display_errors', ($_ENV['ENVIRONEMENT'] === 'development') ? 1 : 0);

include_once 'office365Interface.php';
$clientOffice = new office365Interface();
$accesToken = $clientOffice->getAccessTokenByCredential();
$users = $clientOffice->getInfoUsers($accesToken);
foreach ($users as $user) {
  echo 'name: ' . $user->getGivenName() . '<br>';
  echo 'display name: ' . $user->getDisplayName() . '<br>';
  echo 'job: ' . $user->getJobTitle() . '<br>';
  echo 'mail: ' . $user->getMail() . '<br>';
  echo 'phone: ' . $user->getMobilePhone() . '<br>';
  echo 'business phone: ' . $user->getBusinessPhones()[0] . '<br>';
  echo 'location: ' . $user->getOfficeLocation() . '<br>';
  echo 'langue prefere: ' . $user->getPreferredLanguage() . '<br>';
  echo 'sur nom: ' . $user->getSurname() . '<br>';
  echo 'userPrincipalName: ' . $user->getUserPrincipalName() . '<br>';
  echo 'id: ' . $user->getId() . '<br>';
}

?>
