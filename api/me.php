<?php
require __DIR__.'/config.php';

if (empty($_SESSION['uid'])) {
  http_response_code(401);
  echo json_encode(['error'=>'unauthorized']); exit;
}

$st = $pdo->prepare('SELECT id, email, username, active_team_id FROM users WHERE id=?');
$st->execute([$_SESSION['uid']]);
$user = $st->fetch();
if (!$user) { session_destroy(); http_response_code(401); echo json_encode(['error'=>'unauthorized']); exit; }

$st = $pdo->prepare('SELECT t.id, t.name, m.role
                     FROM memberships m JOIN teams t ON t.id=m.team_id
                     WHERE m.user_id=? ORDER BY t.name');
$st->execute([$_SESSION['uid']]);
$teams = $st->fetchAll();

echo json_encode([
  'id' => (int)$user['id'],
  'email' => $user['email'],
  'username' => $user['username'],
  'active_team_id' => $user['active_team_id'] ? (int)$user['active_team_id'] : null,
  'teams' => $teams
]);
