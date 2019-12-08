<?php

use Microsoft\Graph\Graph;

require_once 'CustomException.php';

class office365Interface
{

  public $clientOAuth;
  private $app_secret;
  private $tenantId;
  private $urlRedirect;
  private $scopes;
  private $domain;

  private $OAUTH_AUTHORITY = 'https://login.microsoftonline.com/';
  private $OAUTH_AUTHORIZE_ENDPOINT = '/oauth2/v2.0/authorize';
  private $OAUTH_TOKEN_ENDPOINT = '/oauth2/v2.0/token';
  private $VERSION_API = '1.0';
  private $OAUTH_TOKEN_ENDPOINTCREDENTIAL = '/oauth2/token';
  private $OAUTH_BASE = 'https://login.microsoftonline.com/';


  /**
   * office365Interface constructor.
   */
  public function __construct()
  {
    $this->app_id = $_ENV['APP_ID'];
    $this->app_secret = $_ENV['APP_PASSWORD'];
    $this->tenantId = $_ENV['TENANT_ID'];
    $this->urlRedirect = $_ENV['REDIRECT_URI'];
    $this->scopes = $_ENV['SCOPES'];
    $this->domain = $_ENV['APP_DOMAIN'];

    $this->clientOAuth = $this->_getOAuthClient();
  }

  public function getAccessTokenByCredential()
  {
    $guzzle = new \GuzzleHttp\Client();
    $token = json_decode($guzzle->post($this->_getUrlTokenPostForCredentialGrantType(), [
      'form_params' => [
        'client_id' => $this->app_id,
        'client_secret' => $this->app_secret,
        'resource' => 'https://graph.microsoft.com/',
        'grant_type' => 'client_credentials',
      ],
    ])->getBody()->getContents());
    return $token->access_token;
  }

  private function _getUrlTokenPostForCredentialGrantType()
  {
    return $this->OAUTH_BASE . $this->tenantId . $this->OAUTH_TOKEN_ENDPOINTCREDENTIAL . '?api-version=' . $this->VERSION_API;

  }


  private function _getOAuthClient()
  {
    return new \League\OAuth2\Client\Provider\GenericProvider([
      'clientId' => $this->app_id,
      'clientSecret' => $this->app_secret,
      'redirectUri' => $this->urlRedirect,
      'urlAuthorize' => $this->OAUTH_AUTHORITY . 'common' . $this->OAUTH_AUTHORIZE_ENDPOINT,
      'urlAccessToken' => $this->OAUTH_AUTHORITY . 'common' . $this->OAUTH_TOKEN_ENDPOINT,
      'urlResourceOwnerDetails' => '',
      'scopes' => $this->scopes
    ]);
  }

  public function getAuthorizationUrl()
  {
    return $this->clientOAuth->getAuthorizationUrl();
  }


  public function getState()
  {
    try {
      return $this->clientOAuth->getState();
    } catch (\League\OAuth2\Client\Provider\Exception\IdentityProviderException $error) {
      $this->interpretationExceptionProvider($error, 'getState');
    }
  }

  public function getAccessTokenByCode($code)
  {
    try {
      return $this->clientOAuth->getAccessToken('authorization_code', [
        'code' => $code
      ]);
    } catch (\League\OAuth2\Client\Provider\Exception\IdentityProviderException $error) {
      $this->interpretationExceptionProvider($error, 'getAccessTokenByCode');
    } catch (Exception $error) {
      $this->interpretationExceptionProvider($error, 'getAccessTokenByCode');
    }
  }

  public function refreshAccessToken($token)
  {
    try {
      return $this->clientOAuth->getAccessToken('refresh_token', [
        'refresh_token' => $token
      ]);
    } catch (\League\OAuth2\Client\Provider\Exception\IdentityProviderException $error) {
      $this->interpretationExceptionProvider($error, 'refreshAccessToken');
    }
  }

  public function getInfoUser($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      $user = $graph->createRequest('GET', '/me')
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
      return $user;
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getInfoUser');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getInfoUser');
    }

  }

  public function getInfoUsers($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      $users = $graph->createRequest('GET', '/users')
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
      return $users;
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getInfoUsers');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getInfoUsers');
    }
  }

  private function formatBodyCreateUser($dataUser)
  {
    $ageGroupAvailalbe = ['null', 'minor', 'notAdult', 'adult'];
    $userTypeAvailalbe = ['Member', 'Guest'];
    $body = [
      'accountEnabled' => $dataUser['enable'],
      'displayName' => $dataUser['name'],
      'mailNickname' => $dataUser['mailNickname'],
      'userPrincipalName' => $dataUser['mailNickname'] . '@' . $this->domain . '.onmicrosoft.com',
      'passwordPolicies' => 'DisablePasswordExpiration, DisableStrongPassword',
      'preferredLanguage' => 'fr-FR',
      'passwordProfile' => [
        'forceChangePasswordNextSignIn' => true,
        'password' => $dataUser['password']],
      // 'createdDateTime' => new DateTime(),
      'userType' => (isset($dataUser['userType']) && in_array($dataUser['userType'], $userTypeAvailalbe)) ?
        $dataUser['userType'] :
       'Guest'
    ];

    if (isset($dataUser['age']) && in_array($dataUser['age'], $ageGroupAvailalbe)) {
      $body['ageGroup'] = $dataUser['age'];
    }
    if (isset($dataUser['birthday'])) {
      $body['birthday'] = $dataUser['birthday'];
    }
    if (isset($dataUser['businessPhones'])) {
      $body['businessPhones'] = $dataUser['businessPhones'];
    }
    if (isset($dataUser['mobilePhone'])) {
      $body['mobilePhone'] = $dataUser['mobilePhone'];
    }
    if (isset($dataUser['website'])) {
      $body['mySite'] = $dataUser['website'];
    }
    if (isset($dataUser['city'])) {
      $body['ville'] = $dataUser['city'];
    }
    if (isset($dataUser['companyName'])) {
      $body['companyName'] = $dataUser['companyName'];
    }
    if (isset($dataUser['country'])) {
      $body['country'] = $dataUser['country'];
    }
    if (isset($dataUser['firstname'])) {
      $body['givenName'] = $dataUser['firstname'];
    }
    if (isset($dataUser['lastname'])) {
      $body['surname'] = $dataUser['lastname'];
    }
    if (isset($dataUser['job'])) {
      $body['jobTitle'] = $dataUser['job'];
    }
    if (isset($dataUser['otherMails'])) {
      $body['otherMails'] = $dataUser['otherMails'];
    }
    return $body;
  }

  public function createOneUser($accessToken, $dataUser)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      $user = $graph->createRequest('POST', '/users')
        ->attachBody($this->formatBodyCreateUser($dataUser))
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
      return $user;
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'createOneUser');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'createOneUser');
    }
  }

  public function getContacts($accessToken)
  {

    $graph = new Graph();
    $graph->setAccessToken($accessToken);

    $user = $graph->createRequest('GET', '/me')
      ->setReturnType(\Microsoft\Graph\Model\User::class)
      ->execute();

    $getContactsUrl = '/me/contacts?' . http_build_query($contactsQueryParams);
    $contacts = $graph->createRequest('GET', $getContactsUrl)
      ->setReturnType(Model\Contact::class)
      ->execute();

    return view('contacts', array(
      'username' => $user->getDisplayName(),
      'contacts' => $contacts
    ));
  }

  private function interpretationExceptionClient(\GuzzleHttp\Exception\ClientException $error, $nameFunction)
  {
    $this->interpretationCodeError($error, $nameFunction);
  }


  private function interpretationExceptionGraph(\Microsoft\Graph\Exception\GraphException $error, $nameFunction)
  {
    $this->interpretationCodeError($error, $nameFunction);
  }

  private function interpretationExceptionProvider(\League\OAuth2\Client\Provider\Exception\IdentityProviderException $error, $nameFunction)
  {
    $this->interpretationCodeError($error, $nameFunction);
  }

  private function interpretationCodeError($error, $nameFunction)
  {
    if ($error->getCode() === 400) {
      throw new CustomException(
        'Office365Interface/' . $nameFunction . '/ [' . $error->getCode() . '] ' . $error->getMessage(),
        OF_ERROR_400
      );
    } else if ($error->getCode() === 401) {
      throw new CustomException(
        'Office365Interface/' . $nameFunction . '/ [' . $error->getCode() . '] ' . $error->getMessage(),
        OF_ERROR_401
      );
    } else if ($error->getCode() === 403) {
      throw new CustomException(
        'Office365Interface/' . $nameFunction . '/ [' . $error->getCode() . '] ' . $error->getMessage(),
        OF_ERROR_403
      );
    } else if ($error->getCode() === 404) {
      throw new CustomException(
        'Office365Interface/' . $nameFunction . '/ [' . $error->getCode() . '] ' . $error->getMessage(),
        OF_ERROR_404
      );
    } else {
      throw new CustomException(
        'Office365Interface/' . $nameFunction . '/ [' . $error->getCode() . '] ' . $error->getMessage(),
        OF_ERROR
      );
    }
  }
}