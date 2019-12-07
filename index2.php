<?php

require __DIR__ . '/vendor/autoload.php';

error_reporting(-1);
ini_set('display_errors', 1);
if (session_status() == PHP_SESSION_NONE) {
  session_start();
}

include_once 'office365Interface.php';
$clientOffice = new office365Interface();
$accesToken = $clientOffice->getAccessTokenByCredential();
$users = $clientOffice->getInfoUsers($accesToken);
foreach ($users as $user) {
  echo 'name: ' . $user->getGivenName().'<br>';
  echo 'display name: ' . $user->getDisplayName().'<br>';
  echo 'job: ' . $user->getJobTitle().'<br>';
  echo 'mail: ' . $user->getMail().'<br>';
  echo 'phone: ' . $user->getMobilePhone().'<br>';
  echo 'business phone: ' . $user->getBusinessPhones()[0].'<br>';
  echo 'location: ' . $user->getOfficeLocation().'<br>';
  echo 'langue prefere: ' . $user->getPreferredLanguage().'<br>';
  echo 'sur nom: ' . $user->getSurname().'<br>';
  echo 'userPrincipalName: ' . $user->getUserPrincipalName().'<br>';
  echo 'id: ' . $user->getId().'<br>';
}

?>
