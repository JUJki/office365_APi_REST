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
$body = [
  'enable'=> true,
  'name'=> 'audrey',
  'mailNickname'=> 'audrey',
  'password'=> '~JceB]ggg9Ns@|>H',
  'userType'=> 'Membre',
  'firstname'=> 'audrey',
  'lastname'=> 'audrey',
  'mobilePhone'=> '0663341706',
  'job'=> 'developpeur',
  'otherMails' => ['audrey.deevd.paris@gmail.com'],
  'userType' => 'Guest'

];
$clientOffice->createOneUser($accesToken, $body);
$users = $clientOffice->getInfoUsers($accesToken);
foreach ($users as $user) {
  $businessPhone = (count($user->getBusinessPhones())> 0) ? $user->getBusinessPhones()[0] : '';
  echo 'name: ' . $user->getGivenName() . '<br>';
  echo 'display name: ' . $user->getDisplayName() . '<br>';
  echo 'job: ' . $user->getJobTitle() . '<br>';
  echo 'mail: ' . $user->getMail() . '<br>';
  echo 'phone: ' . $user->getMobilePhone() . '<br>';
  echo 'business phone: ' . $businessPhone. '<br>';
  echo 'location: ' . $user->getOfficeLocation() . '<br>';
  echo 'langue prefere: ' . $user->getPreferredLanguage() . '<br>';
  echo 'sur nom: ' . $user->getSurname() . '<br>';
  echo 'userPrincipalName: ' . $user->getUserPrincipalName() . '<br>';
  echo 'id: ' . $user->getId() . '<br>';
}

?>
