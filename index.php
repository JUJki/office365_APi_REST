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
$dataContact = [
  'assistantName' => 'juju',
  'givenName' => 'julien',
  'companyName' => 'janneautcf',
  'displayName' => 'julienjanneau',
  'email' => 'julisdddssssen@testDevJulien.onmicrosoft.com',
  'id' => 'julidddden@free.fr'
];
$user = $clientOffice->getInfoUser($_SESSION['access_token']);
$contactsUser = $clientOffice->getContactUserConnected($_SESSION['access_token']);
//$contact = $clientOffice->addContactUserConnected($_SESSION['access_token'], $dataContact);
//var_dump($contact);
/*foreach ($photoUser as $user) {
  echo 'name: ' . $user->getDisplayName() . '<br>';
}*/
exit();

?>
