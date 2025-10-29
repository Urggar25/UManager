<?php
require __DIR__.'/config.php';

function legacy_password_hash(string $password): string {
  return hash('sha256', 'umanager::'.$password);
}

$raw = file_get_contents('php://input');
$data = json_decode($raw, true) ?? [];

$identifier = trim((string)($data['email'] ?? ''));
$password = (string)($data['password'] ?? '');

if ($identifier === '' || $password === '') {
  http_response_code(400);
  echo json_encode(['error' => 'missing_fields']);
  exit;
}

$isEmail = filter_var($identifier, FILTER_VALIDATE_EMAIL) !== false;
$user = null;

if ($isEmail) {
  $st = $pdo->prepare('SELECT * FROM users WHERE LOWER(email)=? LIMIT 1');
  $st->execute([strtolower($identifier)]);
  $user = $st->fetch();
}

if (!$user) {
  $st = $pdo->prepare('SELECT * FROM users WHERE username=? LIMIT 1');
  $st->execute([$identifier]);
  $user = $st->fetch();
}

if (!$user) {
  http_response_code(401);
  echo json_encode(['error' => 'invalid_credentials']);
  exit;
}

$storedHash = isset($user['pass_hash']) ? (string)$user['pass_hash'] : '';
$authenticated = false;
$needsRehash = false;

if ($storedHash !== '' && password_verify($password, $storedHash)) {
  $authenticated = true;
  $needsRehash = password_needs_rehash($storedHash, PASSWORD_DEFAULT);
}

if (!$authenticated && $storedHash !== '') {
  $legacyHash = legacy_password_hash($password);
  if (hash_equals($storedHash, $legacyHash)) {
    $authenticated = true;
    $needsRehash = true;
  }
}

if (!$authenticated && isset($user['password']) && is_string($user['password'])) {
  if (hash_equals((string)$user['password'], $password)) {
    $authenticated = true;
    $needsRehash = true;
  }
}

if (!$authenticated) {
  http_response_code(401);
  echo json_encode(['error' => 'invalid_credentials']);
  exit;
}

if ($needsRehash || $storedHash === '') {
  try {
    $newHash = password_hash($password, PASSWORD_DEFAULT);
    $upd = $pdo->prepare('UPDATE users SET pass_hash=? WHERE id=?');
    $upd->execute([$newHash, (int)$user['id']]);
    $user['pass_hash'] = $newHash;
  } catch (Throwable $e) {
    // Ignore les erreurs de hachage afin de ne pas bloquer l'authentification.
  }
}

$_SESSION['uid'] = (int)$user['id'];

echo json_encode([
  'id' => (int)$user['id'],
  'username' => $user['username'] ?? null,
  'email' => $user['email'] ?? null,
  'active_team_id' => !empty($user['active_team_id']) ? (int)$user['active_team_id'] : null,
]);
