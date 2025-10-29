<?php
require __DIR__.'/config.php';
$user = require_auth($pdo);

$raw = file_get_contents('php://input');
$data = json_decode($raw, true) ?? [];
$teamId = (int)($data['team_id'] ?? 0);

if ($teamId <= 0) {
  http_response_code(400);
  echo json_encode(['error' => 'team_id_required']);
  exit;
}

$st = $pdo->prepare('SELECT 1 FROM memberships WHERE user_id=? AND team_id=?');
$st->execute([$user['id'], $teamId]);
if (!$st->fetchColumn()) {
  http_response_code(403);
  echo json_encode(['error' => 'not_a_member']);
  exit;
}

$pdo->prepare('UPDATE users SET active_team_id=? WHERE id=?')->execute([$teamId, $user['id']]);
echo json_encode(['ok' => true, 'active_team_id' => $teamId]);
