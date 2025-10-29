<?php
require __DIR__.'/config.php';

$currentUser = require_super_admin($pdo);

$method = $_SERVER['REQUEST_METHOD'] ?? 'GET';

if ($method === 'GET') {
  handle_get_admin_users($pdo);
  exit;
}

if ($method === 'DELETE') {
  handle_delete_admin_user($pdo, $currentUser);
  exit;
}

http_response_code(405);
header('Allow: GET, DELETE');
echo json_encode(['error' => 'method_not_allowed']);
exit;

function handle_get_admin_users(PDO $pdo): void {
  $sql = "SELECT 
            u.id,
            u.username,
            u.email,
            u.active_team_id,
            COUNT(m.team_id) AS team_count,
            SUM(CASE WHEN LOWER(m.role) = 'owner' THEN 1 ELSE 0 END) AS owned_team_count
          FROM users u
          LEFT JOIN memberships m ON m.user_id = u.id
          GROUP BY u.id, u.username, u.email, u.active_team_id
          ORDER BY u.username";

  $rows = $pdo->query($sql)->fetchAll();
  $users = array_map('format_admin_user_row', $rows);

  echo json_encode(['users' => $users]);
}

function handle_delete_admin_user(PDO $pdo, array $currentUser): void {
  $userId = (int)($_GET['id'] ?? 0);
  if ($userId <= 0) {
    http_response_code(400);
    echo json_encode(['error' => 'invalid_request']);
    return;
  }

  if ($userId === (int)($currentUser['id'] ?? 0)) {
    http_response_code(409);
    echo json_encode(['error' => 'cannot_delete_self']);
    return;
  }

  $st = $pdo->prepare('SELECT id, email, username FROM users WHERE id=? LIMIT 1');
  $st->execute([$userId]);
  $target = $st->fetch();

  if (!$target) {
    http_response_code(404);
    echo json_encode(['error' => 'user_not_found']);
    return;
  }

  if (is_super_admin_user($target)) {
    http_response_code(409);
    echo json_encode(['error' => 'cannot_delete_super_admin']);
    return;
  }

  $owns = $pdo->prepare("SELECT COUNT(*) FROM memberships WHERE user_id=? AND LOWER(role)='owner'");
  $owns->execute([$userId]);
  $ownedCount = (int)$owns->fetchColumn();

  if ($ownedCount > 0) {
    http_response_code(409);
    echo json_encode(['error' => 'owns_teams', 'owned_team_count' => $ownedCount]);
    return;
  }

  try {
    $pdo->beginTransaction();

    $deleteInvites = $pdo->prepare('DELETE FROM team_invitations WHERE invited_user_id=? OR invited_by=?');
    $deleteInvites->execute([$userId, $userId]);

    $deleteMemberships = $pdo->prepare('DELETE FROM memberships WHERE user_id=?');
    $deleteMemberships->execute([$userId]);

    $deleteUser = $pdo->prepare('DELETE FROM users WHERE id=?');
    $deleteUser->execute([$userId]);

    $pdo->commit();
  } catch (Throwable $e) {
    $pdo->rollBack();
    http_response_code(500);
    echo json_encode(['error' => 'delete_failed']);
    return;
  }

  echo json_encode(['ok' => true]);
}

function format_admin_user_row(array $row): array {
  $id = isset($row['id']) ? (int)$row['id'] : null;
  $teamCount = isset($row['team_count']) ? (int)$row['team_count'] : 0;
  $ownedTeamCount = isset($row['owned_team_count']) ? (int)$row['owned_team_count'] : 0;
  $activeTeamId = isset($row['active_team_id']) && $row['active_team_id'] !== null
    ? (int)$row['active_team_id']
    : null;

  return [
    'id' => $id,
    'username' => $row['username'] ?? null,
    'email' => $row['email'] ?? null,
    'active_team_id' => $activeTeamId,
    'team_count' => $teamCount,
    'owned_team_count' => $ownedTeamCount,
    'is_super_admin' => is_super_admin_user($row),
  ];
}
