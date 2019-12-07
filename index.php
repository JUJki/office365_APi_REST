<?php

require __DIR__ . '/vendor/autoload.php';

error_reporting(-1);
ini_set('display_errors', 1);
if (session_status() == PHP_SESSION_NONE) {
  session_start();
}

include_once 'office365Interface.php';
$clientOffice = new office365Interface();

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

  $accessToken = $clientOffice->getAccessTokenByCode($_GET['code']);
  $_SESSION['access_token'] = $accessToken->getToken();

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

  $user = $clientOffice->getInfoUsers($accessToken);
  var_dump($user);
  exit();
}
?>
