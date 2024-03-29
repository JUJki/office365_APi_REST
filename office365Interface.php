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

  /**
   * Permet de formatter un contact graph
   * @param array $dataContact
   * @return \Microsoft\Graph\Model\Contact
   */
  private function _formatBodyAddContact($dataContact)
  {
    $contact = new \Microsoft\Graph\Model\Contact();
    if (isset($dataContact['assistantName'])) {
      $contact->setAssistantName($dataContact['assistantName']);
    }
    if (isset($dataContact['givenName'])) {
      $contact->setGivenName($dataContact['givenName']);
    }
    if (isset($dataContact['companyName'])) {
      $contact->setCompanyName($dataContact['companyName']);
    }
    if (isset($dataContact['displayName'])) {
      $contact->setDisplayName($dataContact['displayName']);
    }
    if (isset($dataContact['email']) && $dataContact['email'] !== NULL) {
      $email = new \Microsoft\Graph\Model\EmailAddress();
      $email->setAddress($dataContact['email']);
      $email->setName($dataContact['displayName']);
      $contact->setEmailAddresses([$email->getProperties()]);
    }
    return $contact;
  }

  /**
   * Formatte l'object contact pour une mise à jour
   * @param string $idContact
   * @param array $dataContactUpdate
   * @return \Microsoft\Graph\Model\Contact
   */
  private function _formatBodyUpdateContact($idContact, $dataContactUpdate)
  {
    $contact = new \Microsoft\Graph\Model\Contact();
    $contact->setId($idContact);
    if (isset($dataContactUpdate['assistantName'])) {
      $contact->setAssistantName($dataContactUpdate['assistantName']);
    }
    if (isset($dataContactUpdate['givenName'])) {
      $contact->setGivenName($dataContactUpdate['givenName']);
    }
    if (isset($dataContactUpdate['companyName'])) {
      $contact->setCompanyName($dataContactUpdate['companyName']);
    }
    if (isset($dataContactUpdate['displayName'])) {
      $contact->setDisplayName($dataContactUpdate['displayName']);
    }
    if (isset($dataContactUpdate['email']) && $dataContactUpdate['email'] !== NULL) {
      $email = new \Microsoft\Graph\Model\EmailAddress();
      $email->setAddress($dataContactUpdate['email']);
      $email->setName($dataContactUpdate['displayName']);
      $contact->setEmailAddresses([$email->getProperties()]);
    }
    return $contact;
  }

  /**
   * Retourne l'objet utilisateur modifié avec des valeurs non obligatoires
   * @param \Microsoft\Graph\Model\User $user
   * @param array $dataUser
   * @return \Microsoft\Graph\Model\User
   * @throws Exception
   */
  private function bodyUserNotRequired(\Microsoft\Graph\Model\User $user, $dataUser)
  {
    $ageGroupAvailalbe = ['null', 'minor', 'notAdult', 'adult'];
    if (isset($dataUser['age']) && in_array($dataUser['age'], $ageGroupAvailalbe)) {
      $user->setAgeGroup($dataUser['age']);
    }
    if (isset($dataUser['birthday'])) {
      $user->setBirthday(new DateTime($dataUser['birthday']));
    }
    if (isset($dataUser['businessPhones'])) {
      $user->setBusinessPhones($dataUser['businessPhones']);
    }
    if (isset($dataUser['mobilePhone'])) {
      $user->setMobilePhone($dataUser['mobilePhone']);
    }
    if (isset($dataUser['website'])) {
      $user->setMySite($dataUser['website']);
    }
    if (isset($dataUser['city'])) {
      $user->setCity($dataUser['city']);
    }
    if (isset($dataUser['companyName'])) {
      $user->setCompanyName($dataUser['companyName']);
    }
    if (isset($dataUser['country'])) {
      $user->setCountry($dataUser['country']);
    }
    if (isset($dataUser['firstname'])) {
      $user->setGivenName($dataUser['firstname']);
    }
    if (isset($dataUser['lastname'])) {
      $user->setSurname($dataUser['lastname']);
    }
    if (isset($dataUser['job'])) {
      $user->setJobTitle($dataUser['job']);
    }
    if (isset($dataUser['otherMails'])) {
      $user->setOtherMails($dataUser['otherMails']);
    }
    return $user;
  }

  /**
   * Format un object user à partir de l'identifiant de l'utilisateur à modifier afin qu'il soit modifié
   * @param string $idUser
   * @param array $dataUpdate
   * @return \Microsoft\Graph\Model\User
   * @throws Exception
   */
  private function _formatBodyUpdateUserById($idUser, $dataUpdate)
  {
    $user = new \Microsoft\Graph\Model\User();
    $user->setId($idUser);
    return $this->_formatBodyUpdateUser($user, $dataUpdate);
  }

  /**
   * Format un object user à partir de l'userPrincipalName de l'utilisateur à modifier afin qu'il soit modifié
   * @param string $userPrincipaleName
   * @param array $dataUpdate
   * @return \Microsoft\Graph\Model\User
   * @throws Exception
   */
  private function _formatBodyUpdateUserByUserPrincipalName($userPrincipaleName, $dataUpdate)
  {
    $user = new \Microsoft\Graph\Model\User();
    $user->setUserPrincipalName($userPrincipaleName);
    return $this->_formatBodyUpdateUser($user, $dataUpdate);
  }

  /**
   * Retourne l'object user mis à jour
   * @param \Microsoft\Graph\Model\User $user
   * @param array $dataUser
   * @return \Microsoft\Graph\Model\User
   * @throws Exception
   */
  private function _formatBodyUpdateUser(\Microsoft\Graph\Model\User $user, $dataUser)
  {
    if (isset($dataUser['enable'])) {
      $user->setAccountEnabled($dataUser['enable']);
    }
    if (isset($dataUser['name'])) {
      $user->setDisplayName($dataUser['name']);
    }
    if (isset($dataUser['mailNickname'])) {
      $user->setMailNickname($dataUser['mailNickname']);
      if (isset($dataUser['domainOrganisationEmail'])) {
        $user->setUserPrincipalName($dataUser['mailNickname'] . '@' . $dataUser['domainOrganisationEmail']);
      } else {
        $user->setUserPrincipalName($dataUser['mailNickname'] . '@' . $this->domain . '.onmicrosoft.com');
      }
    }
    if (isset($dataUser['language'])) {
      $user->setPreferredLanguage($dataUser['language']);
    }
    if (isset($dataUser['password'])) {
      $password = new \Microsoft\Graph\Model\PasswordProfile();
      $password->setPassword($dataUser['password']);
      $user->setPasswordProfile($password->getProperties());
    }
    return $this->bodyUserNotRequired($user, $dataUser);
  }

  /**
   * Retourne un object user pour la creation d'un nouveau utilisateur
   * @param array $dataUser
   * @return \Microsoft\Graph\Model\User
   * @throws Exception
   */
  private function formatBodyCreateUser($dataUser)
  {

    $userTypeAvailalbe = ['Member', 'Guest'];
    $enable = (isset($dataUser['enable']) && is_bool($dataUser['enable'])) ?
      $dataUser['enable'] :
      true;
    $userType = (isset($dataUser['userType']) && in_array($dataUser['userType'], $userTypeAvailalbe)) ?
      $dataUser['userType'] :
      'Guest';
    $user = new \Microsoft\Graph\Model\User();
    $user->setAccountEnabled($enable);
    $user->setDisplayName($dataUser['name']);
    $user->setMailNickname($dataUser['mailNickname']);
    if (isset($dataUser['domainOrganisationEmail'])) {
      $user->setUserPrincipalName($dataUser['mailNickname'] . '@' . $dataUser['domainOrganisationEmail']);
    } else {
      $user->setUserPrincipalName($dataUser['mailNickname'] . '@' . $this->domain . '.onmicrosoft.com');
    }
    $user->setPreferredLanguage('fr-FR');
    $user->setPasswordPolicies('DisablePasswordExpiration, DisableStrongPassword');
    $password = new \Microsoft\Graph\Model\PasswordProfile();
    $password->setForceChangePasswordNextSignIn(true);
    $password->setPassword($dataUser['password']);
    $user->setPasswordProfile($password->getProperties());
    $user->setUserType($userType);
    return $this->bodyUserNotRequired($user, $dataUser);
  }

  /**
   * Retourne un object organisation à partir de son identifiant et avec les valeurs à mettre à jour
   * @param string $idOrganisation
   * @param array $dataUpdate
   * @return \Microsoft\Graph\Model\Organization
   */
  private function _formatUpdateOrganisation($idOrganisation, $dataUpdate)
  {
    $organisation = new \Microsoft\Graph\Model\Organization();
    $organisation->setId($idOrganisation);
    if (isset($dataUpdate['notificationMarketingEmail']) && $dataUpdate['notificationMarketingEmail'] !== NULL) {
      if (is_string($dataUpdate['notificationMarketingEmail'])) {
        $dataUpdate['notificationMarketingEmail'] = [$dataUpdate['notificationMarketingEmail']];
      }
      if (is_array($dataUpdate['notificationMarketingEmail'])) {
        $organisation->setMarketingNotificationEmails($dataUpdate['notificationMarketingEmail']);
      }
    }
    if (isset($dataUpdate['notificationTechnicalEmail']) && $dataUpdate['notificationTechnicalEmail'] !== NULL) {
      if (is_string($dataUpdate['notificationTechnicalEmail'])) {
        $dataUpdate['notificationTechnicalEmail'] = [$dataUpdate['notificationTechnicalEmail']];
      }
      if (is_array($dataUpdate['notificationTechnicalEmail'])) {
        $organisation->setTechnicalNotificationMails($dataUpdate['notificationTechnicalEmail']);
      }

    }
    if (isset($dataUpdate['notificationSecurityEmail']) && $dataUpdate['notificationSecurityEmail'] !== NULL) {
      if (is_string($dataUpdate['notificationSecurityEmail'])) {
        $dataUpdate['notificationSecurityEmail'] = [$dataUpdate['notificationSecurityEmail']];
      }
      if (is_array($dataUpdate['notificationSecurityEmail'])) {
        $organisation->setSecurityComplianceNotificationMails($dataUpdate['notificationSecurityEmail']);
      }
    }
    if (isset($dataUpdate['notificationSecurityPhone']) && $dataUpdate['notificationSecurityPhone'] !== NULL) {
      if (is_string($dataUpdate['notificationSecurityPhone'])) {
        $dataUpdate['notificationSecurityPhone'] = [$dataUpdate['notificationSecurityPhone']];
      }
      if (is_array($dataUpdate['notificationSecurityPhone'])) {
        $organisation->setSecurityComplianceNotificationPhones($dataUpdate['notificationSecurityPhone']);
      }
    }
    return $organisation;
  }

  /**
   * Permet de recuperer un access token à partir d'un acces application
   * @return string
   */
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

  /**
   * Permet de retourner l'url d'authentification
   * @return string
   */
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

  /**
   * Permet d'avoir un access token à partir d'un code
   * @param string $code
   * @return \League\OAuth2\Client\Token\AccessTokenInterface
   */
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

  /**
   * Permet de refresh le token
   * @param string $token
   * @return \League\OAuth2\Client\Token\AccessTokenInterface
   */
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

  /**
   * Permet de recuperer les informations des organisations
   * @param string $accessToken
   * @return \Microsoft\Graph\Model\Organization[]
   */
  public function getOrganization($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/organization')
        ->setReturnType(\Microsoft\Graph\Model\Organization::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getOrganization');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getOrganization');
    }
  }

  /**
   * Permet de mettre à jour une organisation à partir de son identifiant
   * @param string $accessToken
   * @param string $idOrganisation
   * @param array $dataUpdate
   * @return \Microsoft\Graph\Model\Organization
   */
  public function updateOneOrganization($accessToken, $idOrganisation, $dataUpdate)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('PATCH', '/organization/' . $idOrganisation)
        ->attachBody($this->_formatUpdateOrganisation($idOrganisation, $dataUpdate))
        ->setReturnType(\Microsoft\Graph\Model\Organization::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getOrganization');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getOrganization');
    }
  }


  /**
   * Retourne les informations d'un utilisateur connecté
   * @param string $accessToken
   * @return \Microsoft\Graph\Model\User
   */
  public function getInfoUserConnected($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/me')
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getInfoUser');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getInfoUser');
    }
  }

  /**
   * Permet de récuperer la list des contacts de l'utilisateur connecté
   * @param string $accessToken
   * @return \Microsoft\Graph\Model\Contact[]
   */
  public function getContactUserConnected($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/me/contacts')
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getContactUserConnected');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getContactUserConnected');
    }
  }

  /**
   * Permet de retourner les personnes classées par petinence pour un utilisateur connecté
   * @param string $accessToken
   * @return \Microsoft\Graph\Model\Person[]
   */
  public function getPeopleUserConnected($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/me/people')
        ->setReturnType(\Microsoft\Graph\Model\Person::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      if ($error->getCode() === 404) {
        return [];
      }
      $this->interpretationExceptionGraph($error, 'getPeopleUserConnected');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      if ($error->getCode() === 404) {
        return [];
      }
      $this->interpretationExceptionClient($error, 'getPeopleUserConnected');
    }
  }

  /**
   * Permet de mettre a jour le contact d'un utilisateur connecté via l'id du contact
   * @param string $accessToken
   * @param string $idContact
   * @param array $dataContact
   * @return \Microsoft\Graph\Model\Contact
   */
  public function updateOneContactUserConnected($accessToken, $idContact, $dataContact)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('PATCH', '/me/contacts/' . $idContact)
        ->attachBody($this->_formatBodyUpdateContact($idContact, $dataContact))
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'updateOneContactUserConnected');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'updateOneContactUserConnected');
    }
  }


  /**
   * Permet d'ajouter un contact à un utilisateur connecté
   * @param string $accessToken
   * @param array $dataContact
   * @return \Microsoft\Graph\Model\Contact
   */
  public function addContactUserConnected($accessToken, $dataContact)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('POST', '/me/contacts')
        ->attachBody($this->_formatBodyAddContact($dataContact))
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'addContactUserConnected');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'addContactUserConnected');
    }
  }

  /**
   * Permet de la suppression d'un contact d'un utilisateur connecté
   * @param string $accessToken
   * @param string $idDeleteContact
   * @return void
   */
  public function deleteContactUserConnected($accessToken, $idDeleteContact)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('DELETE', '/me/contacts/' . $idDeleteContact)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'deleteContactUserConnected');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'deleteContactUserConnected');
    }
  }

  /**
   * Permet de recuperer les informations d'un contact par son id et par l'id de l'utilisateur auquel il appartient
   * @param string $accessToken
   * @param string $id
   * @param string $idContact
   * @return \Microsoft\Graph\Model\Contact
   */
  public function getOneContactUserById($accessToken, $id, $idContact)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/users/' . $id . '/contacts/' . $idContact)
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getOneContactUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getOneContactUserById');
    }
  }

  /**
   * Permet de recuperer les informations d'un contact par son id et par l'userPrincipalName de l'utilisateur auquel il appartient
   * @param string $accessToken
   * @param string $userPrincipalName
   * @param string $idContact
   * @return \Microsoft\Graph\Model\Contact
   */
  public function getOneContactUserByUserPrincipalName($accessToken, $userPrincipalName, $idContact)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/users/' . $userPrincipalName . '/contacts/' . $idContact)
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getOneContactUserByUserPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getOneContactUserByUserPrincipalName');
    }
  }

  /**
   * Permet de mettre à jour les informations d'un contact par son id et par l'id de l'utilisateur auquel il appartient
   * @param string $accessToken
   * @param string $id
   * @param string $idContact
   * @param array $dataUpdate
   * @return \Microsoft\Graph\Model\Contact
   */
  public function updateOneContactUserById($accessToken, $id, $idContact, $dataUpdate)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('PATCH', '/users/' . $id . '/contacts/' . $idContact)
        ->attachBody($this->_formatBodyUpdateContact($idContact, $dataUpdate))
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'updateOneContactUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'updateOneContactUserById');
    }
  }

  /**
   * Permet de mettre à jour les informations d'un contact par son id et par l'userPrincipalName de l'utilisateur auquel il appartient
   * @param string $accessToken
   * @param string $userPrincipalName
   * @param string $idContact
   * @param array $dataUpdate
   * @return \Microsoft\Graph\Model\Contact
   */
  public function updateOneContactUserByUserPrincipalName($accessToken, $userPrincipalName, $idContact, $dataUpdate)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('PATCH', '/users/' . $userPrincipalName . '/contacts/' . $idContact)
        ->attachBody($this->_formatBodyUpdateContact($idContact, $dataUpdate))
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'updateOneContactUserByUserPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'updateOneContactUserByUserPrincipalName');
    }
  }

  /**
   * Permet d'ajouter un contact a un utilisateur en indiquant l'id de l'utilisateur
   * @param string $accessToken
   * @param string $id
   * @param array $dataContact
   * @return \Microsoft\Graph\Model\Contact
   */
  public function addContactUserById($accessToken, $id, $dataContact)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('POST', '/users/' . $id . '/contacts')
        ->attachBody($this->_formatBodyAddContact($dataContact))
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'addContactUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'addContactUserById');
    }
  }

  /**
   * Permet d'ajouter un contact a un utilisateur en indiquant l'userPrinicipalName de l'utilisateur
   * @param string $accessToken
   * @param string $userPrincipalName
   * @param array $dataContact
   * @return \Microsoft\Graph\Model\Contact
   */
  public function addContactUserByUserPrincipalName($accessToken, $userPrincipalName, $dataContact)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('POST', '/users/' . $userPrincipalName . '/contacts')
        ->attachBody($this->_formatBodyAddContact($dataContact))
        ->setReturnType(\Microsoft\Graph\Model\Contact::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'addContactUserByUserPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'addContactUserByUserPrincipalName');
    }
  }

  /**
   * Permet de supprimer un contact d'un utlisateur en indiquant l'id de l'utlisateur
   * @param $accessToken
   * @param $id
   * @param $idContactDelete
   * @return mixed
   */
  public function deleteContactUserById($accessToken, $id, $idContactDelete)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('DELETE', '/users/' . $id . '/contacts/' . $idContactDelete)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'deleteContactUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'deleteContactUserById');
    }
  }

  /**
   * Permet de supprimer un contact d'un utlisateur en indiquant  l'userPrincipalNale de l'utilisateur
   * @param string $accessToken
   * @param string $userPrincipalName
   * @param string $idContactDelete
   * @return void
   */
  public function deleteContactUserByUserPrincipalName($accessToken, $userPrincipalName, $idContactDelete)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('DELETE', '/users/' . $userPrincipalName . '/contacts/' . $idContactDelete)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'deleteContactUserByUserPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'deleteContactUserByUserPrincipalName');
    }
  }

  // non fonctionnel

  /**
   * Permet de recuperer les informations photos d'un profil connecté
   * @param $accessToken
   * @return Microsoft\Graph\Model\ProfilePhoto
   */
  public function getPhotoUserConnected($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/me/photo/$value')
        ->setReturnType(\Microsoft\Graph\Model\ProfilePhoto::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      if ($error->getCode() === 404) {
        return false;
      }
      $this->interpretationExceptionGraph($error, 'getPhotoUserConnected');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      if ($error->getCode() === 404) {
        return false;
      }
      $this->interpretationExceptionClient($error, 'getPhotoUserConnected');
    }
  }

  /**
   * Retourne une liste de tous les utlisateurs
   * @param string $accessToken
   * @return \Microsoft\Graph\Model\User[]
   */
  public function getInfoUsers($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/users')
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getInfoUsers');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getInfoUsers');
    }
  }

  /**
   * Permet de retourner les utilisateurs nouvellement créés, modifiés ou supprimés
   * @param string $accessToken
   * @return \Microsoft\Graph\Model\User[]
   */
  public function getDeltaUsers($accessToken)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/users/delta')
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getDeltaUsers');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getDeltaUsers');
    }
  }

  /**
   * Creation d'un nouvel utilisateur
   * @param string $accessToken
   * @param array $dataUser
   * @return \Microsoft\Graph\Model\User
   */
  public function addOneUser($accessToken, $dataUser)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      return $graph->createRequest('POST', '/users')
        ->attachBody($this->formatBodyCreateUser($dataUser))
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'addOneUser');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'addOneUser');
    }
  }

  /**
   * Recupère un utilisateur à partir de son id
   * @param string $accessToken
   * @param string $idUser
   * @return \Microsoft\Graph\Model\User
   */
  public function getOneUserById($accessToken, $idUser)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      return $graph->createRequest('GET', '/users/' . $idUser)
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getOneUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getOneUserById');
    }
  }

  /**
   * Recupère un utilisateur à partir de son principalName
   * @param string $accessToken
   * @param string $principalName
   * @return \Microsoft\Graph\Model\User
   */
  public function getOneUserByPrincipalName($accessToken, $principalName)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      return $graph->createRequest('GET', '/users/' . $principalName)
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'getOneUserByPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'getOneUserByPrincipalName');
    }
  }

  /**
   * Permet de retourner les personnes classées par petinence pour un utilisateur selon son identifiant
   * @param string $accessToken
   * @param string $idUser
   * @return \Microsoft\Graph\Model\Person[]
   */
  public function getPeopleUserById($accessToken, $idUser)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/users/' . $idUser . '/people')
        ->setReturnType(\Microsoft\Graph\Model\Person::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      if ($error->getCode() === 404) {
        return [];
      }
      $this->interpretationExceptionGraph($error, 'getPeopleUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      if ($error->getCode() === 404) {
        return [];
      }
      $this->interpretationExceptionClient($error, 'getPeopleUserById');
    }
  }

  /**
   * Permet de retourner les personnes classées par petinence pour un utilisateur selon son userPrincipalName
   * @param string $accessToken
   * @param string $userPrincipalName
   * @return \Microsoft\Graph\Model\Person[]
   */
  public function getPeopleUserByUserPrincipalName($accessToken, $userPrincipalName)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);
    try {
      return $graph->createRequest('GET', '/users/' . $userPrincipalName . '/people')
        ->setReturnType(\Microsoft\Graph\Model\Person::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      if ($error->getCode() === 404) {
        return [];
      }
      $this->interpretationExceptionGraph($error, 'getPeopleUserByUserPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      if ($error->getCode() === 404) {
        return [];
      }
      $this->interpretationExceptionClient($error, 'getPeopleUserByUserPrincipalName');
    }
  }

  /**
   * Recupere les photos de profils d'un utilisateurs à partir de son identifiant
   * @param string $accessToken
   * @param string $id
   * @return \Microsoft\Graph\Model\ProfilePhoto
   */
  public function getPhotoByIdUser($accessToken, $id)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      $user = $graph->createRequest('GET', '/users/' . $id . '/photo/$value')
        ->setReturnType(\Microsoft\Graph\Model\ProfilePhoto::class)
        ->execute();
      return $user;
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      if ($error->getCode() === 404) {
        return false;
      }
      $this->interpretationExceptionGraph($error, 'getPhotoByIdUser');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      if ($error->getCode() === 404) {
        return false;
      }
      $this->interpretationExceptionClient($error, 'getPhotoByIdUser');
    }
  }

  /**
   * Recupere les photos de profils d'un utilisateurs à partir de son userPrincipalName
   * @param string $accessToken
   * @param string $userPrincipalName
   * @return \Microsoft\Graph\Model\User
   */
  public function getPhotoByIUserPrincipalName($accessToken, $userPrincipalName)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      $user = $graph->createRequest('GET', '/users/' . $userPrincipalName . '/photo/$value')
        ->setReturnType(\Microsoft\Graph\Model\ProfilePhoto::class)
        ->execute();
      return $user;
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      if ($error->getCode() === 404) {
        return false;
      }
      $this->interpretationExceptionGraph($error, 'getPhotoByIUserPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      if ($error->getCode() === 404) {
        return false;
      }
      $this->interpretationExceptionClient($error, 'getPhotoByIUserPrincipalName');
    }
  }

  /**
   * Modifie un utilisateur à partir de son id
   * @param string $accessToken
   * @param string $idUser
   * @param array $dataUser
   * @return \Microsoft\Graph\Model\User
   */
  public function updateOneUserById($accessToken, $idUser, $dataUser)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      return $graph->createRequest('PATCH', '/users/' . $idUser)
        ->attachBody($this->_formatBodyUpdateUserById($idUser, $dataUser))
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'updateOneUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'updateOneUserById');
    }
  }

  /**
   * Modifie un utilisateur à partir de son principalName
   * @param string $accessToken
   * @param string $principalName
   * @param array $dataUser
   * @return \Microsoft\Graph\Model\User
   */
  public function updateOneUserByPrincipalName($accessToken, $principalName, $dataUser)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      return $graph->createRequest('PATCH', '/users/' . $principalName)
        ->attachBody($this->_formatBodyUpdateUserByUserPrincipalName($principalName, $dataUser))
        ->setReturnType(\Microsoft\Graph\Model\User::class)
        ->execute();
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'updateOneUserByPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'updateOneUserByPrincipalName');
    }
  }

  /**
   * Supprime un utilisateur à partir de son id
   * @param string $accessToken
   * @param string $idUser
   * @return bool
   */
  public function deleteOneUserById($accessToken, $idUser)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      $graph->createRequest('DELETE', '/users/' . $idUser)
        ->execute();
      return true;
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'deleteOneUserById');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'deleteOneUserById');
    }
  }

  /**
   * Supprime un utilisateur à partir de son principalName
   * @param string $accessToken
   * @param string $principalName
   * @return bool
   */
  public function deleteOneUserByPrincipalName($accessToken, $principalName)
  {
    $graph = new Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    try {
      $graph->createRequest('DELETE', '/users/' . $principalName)
        ->execute();
      return true;
    } catch (\Microsoft\Graph\Exception\GraphException $error) {
      $this->interpretationExceptionGraph($error, 'deleteOneUserByPrincipalName');
    } catch (\GuzzleHttp\Exception\ClientException $error) {
      $this->interpretationExceptionClient($error, 'deleteOneUserByPrincipalName');
    }
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