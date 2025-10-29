<?php
require __DIR__.'/config.php';

$raw = file_get_contents('php://input');
$data = json_decode($raw, true) ?? [];

$email = trim($data['email'] ?? '');
$password = (string)($data['password'] ?? '');

if ($email === '' || $password === '') {
  http_response_code(400);
  echo json_encode(['error' => 'missing_fields']);
  exit;
}

$st = $pdo->prepare('SELECT * FROM users WHERE email=?');
$st->execute([$email]);
$user = $st->fetch();

if (!$user || !password_verify($password, $user['pass_hash'])) {
  http_response_code(401);
  echo json_encode(['error' => 'invalid_credentials']);
  exit;
}

$_SESSION['uid'] = (int)$user['id'];

echo json_encode([
  'id' => (int)$user['id'],
  'username' => $user['username'],
  'active_team_id' => $user['active_team_id'] ? (int)$user['active_team_id'] : null
]);
