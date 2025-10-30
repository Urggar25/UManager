<?php
require __DIR__.'/config.php';

if (($_SERVER['REQUEST_METHOD'] ?? 'GET') !== 'POST') {
  http_response_code(405);
  header('Allow: POST');
  echo json_encode(['error' => 'method_not_allowed']);
  exit;
}

$raw = file_get_contents('php://input');
$data = json_decode($raw, true);

if (!is_array($data)) {
  http_response_code(400);
  echo json_encode(['error' => 'invalid_json']);
  exit;
}

$username = trim((string)($data['username'] ?? ''));
$email = trim((string)($data['email'] ?? ''));
$password = (string)($data['password'] ?? '');

if ($username === '' || $email === '' || $password === '') {
  http_response_code(400);
  echo json_encode(['error' => 'missing_fields']);
  exit;
}

if (strlen($username) < 3 || strlen($username) > 50) {
  http_response_code(400);
  echo json_encode(['error' => 'invalid_username']);
  exit;
}

if (!preg_match('/^[A-Za-z0-9._-]+$/', $username)) {
  http_response_code(400);
  echo json_encode(['error' => 'invalid_username']);
  exit;
}

if (!filter_var($email, FILTER_VALIDATE_EMAIL)) {
  http_response_code(400);
  echo json_encode(['error' => 'invalid_email']);
  exit;
}

if (strlen($password) < 6) {
  http_response_code(400);
  echo json_encode(['error' => 'password_too_short']);
  exit;
}

$lowerEmail = strtolower($email);

$checkEmail = $pdo->prepare('SELECT id FROM users WHERE LOWER(email)=? LIMIT 1');
$checkEmail->execute([$lowerEmail]);
if ($checkEmail->fetch()) {
  http_response_code(409);
  echo json_encode(['error' => 'email_already_used']);
  exit;
}

$checkUsername = $pdo->prepare('SELECT id FROM users WHERE username=? LIMIT 1');
$checkUsername->execute([$username]);
if ($checkUsername->fetch()) {
  http_response_code(409);
  echo json_encode(['error' => 'username_already_used']);
  exit;
}

$passwordHash = password_hash($password, PASSWORD_DEFAULT);

try {
  $insert = $pdo->prepare('INSERT INTO users (username, email, pass_hash) VALUES (?, ?, ?)');
  $insert->execute([$username, $email, $passwordHash]);
  $userId = (int)$pdo->lastInsertId();
} catch (Throwable $e) {
  if ($e instanceof PDOException && $e->getCode() === '23000') {
    $message = strtolower($e->getMessage());
    http_response_code(409);
    if (strpos($message, 'username') !== false) {
      echo json_encode(['error' => 'username_already_used']);
    } elseif (strpos($message, 'email') !== false) {
      echo json_encode(['error' => 'email_already_used']);
    } else {
      echo json_encode(['error' => 'registration_conflict']);
    }
    exit;
  }

  http_response_code(500);
  echo json_encode(['error' => 'registration_failed']);
  exit;
}

$select = $pdo->prepare('SELECT id, username, email, active_team_id FROM users WHERE id=? LIMIT 1');
$select->execute([$userId]);
$user = $select->fetch();

if (!$user) {
  http_response_code(500);
  echo json_encode(['error' => 'registration_failed']);
  exit;
}

session_regenerate_id(true);
$_SESSION['uid'] = (int)$user['id'];

echo json_encode([
  'id' => (int)$user['id'],
  'username' => $user['username'] ?? null,
  'email' => $user['email'] ?? null,
  'active_team_id' => !empty($user['active_team_id']) ? (int)$user['active_team_id'] : null,
]);
