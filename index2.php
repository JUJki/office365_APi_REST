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


$organisations = $clientOffice->getOrganization($accesToken);?>
<h1>Organisation</h1>
<?php
foreach($organisations as $organisation) {
  echo $organisation->getDisplayName().'</br>';
  echo '<pre>';
  var_dump($organisation->getProperties());
  echo '</pre>';
  echo '<br>';
  $verifiesDomains = $organisation->getVerifiedDomains();
  foreach ($verifiesDomains as $domain) {
    $domainOrganisationEmail = $domain['name'];
  }
}?>

<h1>List des users </h1>
<?
$users = $clientOffice->getInfoUsers($accesToken); // ou $users = $clientOffice->getDeltaUsers($accesToken);
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
  echo 'id: ' . $user->getId() . '<br><br>';


  //$clientOffice->updateOneUserById($accesToken, $user->getId(),['name'=>'zzzzz']);
  //$clientOffice->updateOneUserByPrincipalName($accesToken, $user->getUserPrincipalName(),['mailNickname'=>'zzzzzz'])
  //$clientOffice->deleteOneUserById($accesToken, $user->getId());
  //$clientOffice->deleteOneUserByPrincipalName($accesToken, $user->getUserPrincipalName());

  /*$dataAddContactUser = [
    'assistantName' => 'zzz',
    'givenName' => 'zzzz',
    'companyName' => 'zzzz',
    'displayName' => 'zzzz',
    'email' => 'zzzzzz@testDevJulien.onmicrosoft.com'
  ];*/
  //$contact = $clientOffice->addContactUserById($accesToken,$user->getId(),$dataAddContactUser);
  //$contact = $clientOffice->addContactUserByUserPrincipalName($accesToken, $user->getUserPrincipalName(), $dataAddContactUser);
  //$user = $clientOffice->getOneUserById($accesToken, $user->getId());
  //$clientOffice->getOneUserByPrincipalName($accesToken, $user->getUserPrincipalName());

}


// creation d'un utilisateur
/*$bodyCreateUser = [
  'enable'=> true,
  'name'=> 'xxx',
  'mailNickname'=> 'xxxx',
  'password'=> '~JceB]sssfpejn@|>H',
  'firstname'=> 'xxxx',
  'lastname'=> 'xxx',
  'mobilePhone'=> '0663342436',
  'job'=> 'analyste',
  'otherMails' => ['juju.pxxx@gmail.com'],
  'userType' => 'Member'
];
$clientOffice->addOneUser($accesToken, $bodyCreateUser);*/


//$photo = $clientOffice->getOneUsersPhoto($accesToken,'02fcd2a7-9781-4d4e-a716-e54f767aaf60');

?>
