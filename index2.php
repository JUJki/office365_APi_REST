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
//$clientOffice->createOneUser($accesToken, $body);
//$clientOffice->deleteOneUserById($accesToken, '839c5346-5ddf-4d19-90bf-8bcddbacc023');
//$clientOffice->deleteOneUserByPrincipalName($accesToken,'upnvalue@testDevJulien.onmicrosoft.com ');
$users = $clientOffice->getInfoUsers($accesToken);
/*foreach ($users as $user) {
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
}*/

/*$user = $clientOffice->getOneUserById($accesToken, '02fcd2a7-9781-4d4e-a716-e54f767aaf60');
echo 'name: ' . $user->getGivenName() . '<br>';
echo 'display name: ' . $user->getDisplayName() . '<br>';
echo 'job: ' . $user->getJobTitle() . '<br>';
echo 'mail: ' . $user->getMail() . '<br>';
echo 'phone: ' . $user->getMobilePhone() . '<br>';
echo 'location: ' . $user->getOfficeLocation() . '<br>';
echo 'langue prefere: ' . $user->getPreferredLanguage() . '<br>';
echo 'sur nom: ' . $user->getSurname() . '<br>';
echo 'userPrincipalName: ' . $user->getUserPrincipalName() . '<br>';
echo 'id: ' . $user->getId() . '<br>';
$user = $clientOffice->getOneUserByPrincipalName($accesToken, 'audrey@testDevJulien.onmicrosoft.com');
echo 'rrr';
echo 'name: ' . $user->getGivenName() . '<br>';
echo 'display name: ' . $user->getDisplayName() . '<br>';
echo 'job: ' . $user->getJobTitle() . '<br>';
echo 'mail: ' . $user->getMail() . '<br>';
echo 'phone: ' . $user->getMobilePhone() . '<br>';
echo 'location: ' . $user->getOfficeLocation() . '<br>';
echo 'langue prefere: ' . $user->getPreferredLanguage() . '<br>';
echo 'sur nom: ' . $user->getSurname() . '<br>';
echo 'userPrincipalName: ' . $user->getUserPrincipalName() . '<br>';
echo 'id: ' . $user->getId() . '<br>';*/


$clientOffice->updateOneUserById($accesToken, '02fcd2a7-9781-4d4e-a716-e54f767aaf60',['name'=>'audrey1']);
$clientOffice->updateOneUserByPrincipalName($accesToken, 'audrey@testDevJulien.onmicrosoft.com',['mailNickname'=>'audrey12'])


?>
