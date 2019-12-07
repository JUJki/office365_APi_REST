<?php
require __DIR__ . '/vendor/autoload.php';

class UsageExample
{
  public function run($accessToken)
  {

    $graph = new \Microsoft\Graph\Graph();
    $graph->setAccessToken($accessToken);

    $user = $graph->createRequest("GET", "/users")
      ->setReturnType(\Microsoft\Graph\Model\User::class)
      ->execute();

    foreach ($user as $item) {
      echo $item->getGivenName();
    }

    echo "Hello, I am  ";
  }
}