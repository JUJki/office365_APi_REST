<?php

use Dotenv\Dotenv;

require __DIR__ . '/vendor/autoload.php';
$dotenv = Dotenv::createImmutable(__DIR__);
$dotenv->load();

error_reporting(($_ENV['ENVIRONEMENT'] === 'development') ? -1 : 0);
ini_set('display_errors', ($_ENV['ENVIRONEMENT'] === 'development') ? 1 : 0);
if (session_status() == PHP_SESSION_NONE) {
  session_start();
}

function storageAndDecodeToken($accessToken)
{
  $_SESSION['access_token'] = $accessToken->getToken();
  $_SESSION['refresh_token'] = $accessToken->getRefreshToken();
  $_SESSION['expire_token'] = $accessToken->getExpires();
  // The id token is a JWT token that contains information about the user
  // It's a base64 coded string that has a header, payload and signature
  $idToken = $accessToken->getValues()['id_token'];
  $decodedAccessTokenPayload = base64_decode(
    explode('.', $idToken)[1]
  );
  $jsonAccessTokenPayload = json_decode($decodedAccessTokenPayload, true);

  // The following user properties are needed in the next page
  $_SESSION['preferred_username'] = $jsonAccessTokenPayload['preferred_username'];
  $_SESSION['given_name'] = $jsonAccessTokenPayload['name'];
}

include_once 'office365Interface.php';
$clientOffice = new office365Interface();

if (isset($_SESSION['access_token'])) {
  storageAndDecodeToken($clientOffice->refreshAccessToken($_SESSION['refresh_token']));
} else {
  if ($_SERVER['REQUEST_METHOD'] === 'GET' && !isset($_GET['code'])) {
    $authorizationUrl = $clientOffice->getAuthorizationUrl();
    $_SESSION['state'] = $clientOffice->getState();
    header('Location: ' . $authorizationUrl);
    exit();
  } elseif ($_SERVER['REQUEST_METHOD'] === 'GET' && isset($_GET['code'])) {
    // Validate the OAuth state parameter
    if (empty($_GET['state']) || ($_GET['state'] !== $_SESSION['state'])) {
      unset($_SESSION['state']);
      exit('State value does not match the one initially sent');
    }
    storageAndDecodeToken($clientOffice->getAccessTokenByCode($_GET['code']));
  }
}


$user = $clientOffice->getInfoUserConnected($_SESSION['access_token']);
echo 'User connected :<br>';
echo '<pre>';
var_dump($user->getProperties());
echo '</pre>';

$peoples = $clientOffice->getPeopleUserConnected($_SESSION['access_token']);
foreach ($peoples as $people) {
  echo '<pre>';
  var_dump($people->getProperties());
  echo '</pre>';
}

$domainOrganisationEmail = '';
$organisations = $clientOffice->getOrganization($_SESSION['access_token']);
echo 'Organisations :</br>';
foreach ($organisations as $organisation) {
  echo $organisation->getDisplayName() . '</br>';
  echo '<pre>';
  var_dump($organisation->getProperties());
  echo '</pre>';
  $verifiesDomains = $organisation->getVerifiedDomains();
  foreach ($verifiesDomains as $domain) {
    $domainOrganisationEmail = $domain['name'];
  }
}

$pictureProfil = $clientOffice->getPhotoUserConnected($_SESSION['access_token']);
if ($pictureProfil) {
  var_dump($pictureProfil->getProperties());
}


// Pour Update une organisation
/* foreach($organisations as $organisation) {
  if($organisation->getId() === 'fdsffsfsdfsdf') {
    $updateOrganisation = ['notificationMarketingEmail'=> 'janneaujulien@xxxxx.fr'];
    $clientOffice->updateOneOrganization($_SESSION['access_token'], $organisation->getId(), $updateOrganisation);
  }
}*/

// Pour creer un utilisateur

/*$bodyCreateUser = [
  'enable'=> true,
  'name'=> 'julien',
  'mailNickname'=> 'julien',
  'password'=> '~JceBdd]sssfpejn@|>H',
  'firstname'=> 'julien',
  'lastname'=> 'janneau',
  'mobilePhone'=> '0663342226',
  'job'=> 'Developpeur',
  'otherMails' => ['juju.par@gmail.com'],
  'userType' => 'Guest', // Guest ou Member
];*/
//$newUser = $clientOffice->addOneUser($_SESSION['access_token'],$bodyCreateUser);

$users = $clientOffice->getInfoUsers($_SESSION['access_token']);
echo 'Liste de tous les utilisateurs :</br>';
foreach ($users as $user) {
  echo '<pre>';
  var_dump($user->getProperties());
  echo '</pre>';
}

// suppression d'un tulisateurs depuis un compte connecté
/*foreach($users as $user) {
  if($user->getSurname() === 'janneau') {
    $clientOffice->deleteContactUserConnected($_SESSION['access_token'], $user->getId());
  }
}*/


// ajout de contacts
/*$dataAddContact = [
  'assistantName' => 'XXXXX',
  'givenName' => 'XXXXXX',
  'companyName' => 'XXXXXX',
  'displayName' => 'mmmms',
  'email' => 'XXXXXX@XXXXXX.XX'
];*/
// 3 methodes possibles
//$contact = $clientOffice->addContactUserConnected($_SESSION['access_token'], $dataAddContact);
//$contact = $clientOffice->addContactUserById($_SESSION['access_token'], $user->getId(),$dataAddContact);
//$contact = $clientOffice->addContactUserByUserPrincipalName($_SESSION['access_token'], $user->getUserPrincipalName(),$dataAddContact);

// listes des contacts connecté
echo 'Liste de tous les contacts :</br>';
$contactsUser = $clientOffice->getContactUserConnected($_SESSION['access_token']);
foreach ($contactsUser as $contact) {
  echo '<pre>';
  var_dump($contact->getProperties());
  echo '</pre>';
  //$contact=$clientOffice->getOneContactUserById($_SESSION['access_token'], $user->getId(), $contact->getId());
  //$contact=$clientOffice->getOneContactUserByUserPrincipalName($_SESSION['access_token'], $user->getUserPrincipalName(), $contact->getId());
  //$clientOffice->deleteContactUserConnected($_SESSION['access_token'], $contact->getId());
  //$clientOffice->deleteContactUserById($_SESSION['access_token'], $user->getId(), $contact->getId());
  //$clientOffice->deleteContactUserByUserPrincipalName($_SESSION['access_token'], $user->getUserPrincipalName(), $contact->getId());
  //$clientOffice->updateOneContactUserConnected($_SESSION['access_token'], $contact->getId(), ['assistantName'=> 'ZZZZZ']);
  //$clientOffice->updateOneContactUserById($_SESSION['access_token'],$user->getId(),$contact->getId(), ['assistantName'=> 'ZZZZZ']);
  // $clientOffice->updateOneContactUserByUserPrincipalName($_SESSION['access_token'], $user->getUserPrincipalName(),$contact->getId(), ['assistantName'=> 'ZZZZZ']);
//
}
exit();
?>
