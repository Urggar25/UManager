<?php
require __DIR__.'/config.php';

$user = require_auth($pdo);

ensure_invitation_schema($pdo);

$method = $_SERVER['REQUEST_METHOD'];

if ($method === 'GET') {
  handle_get($pdo, $user);
  exit;
}

if ($method === 'POST') {
  handle_post($pdo, $user);
  exit;
}

if ($method === 'PATCH') {
  handle_patch($pdo, $user);
  exit;
}

if ($method === 'DELETE') {
  handle_delete($pdo, $user);
  exit;
}

http_response_code(405);
echo json_encode(['error' => 'method_not_allowed']);
exit;

function ensure_invitation_schema(PDO $pdo): void {
  $pdo->exec(
    "CREATE TABLE IF NOT EXISTS team_invitations (
      id INT AUTO_INCREMENT PRIMARY KEY,
      team_id INT NOT NULL,
      email VARCHAR(255) NOT NULL,
      invited_user_id INT DEFAULT NULL,
      invited_by INT NOT NULL,
      status VARCHAR(20) NOT NULL DEFAULT 'pending',
      created_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP,
      responded_at DATETIME DEFAULT NULL,
      UNIQUE KEY uniq_team_email_status (team_id, email, status),
      INDEX idx_invite_email_status (email, status),
      INDEX idx_invite_team_status (team_id, status),
      CONSTRAINT fk_invite_team FOREIGN KEY (team_id) REFERENCES teams(id) ON DELETE CASCADE,
      CONSTRAINT fk_invite_user FOREIGN KEY (invited_user_id) REFERENCES users(id) ON DELETE SET NULL,
      CONSTRAINT fk_invite_author FOREIGN KEY (invited_by) REFERENCES users(id) ON DELETE CASCADE
    ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4"
  );
}

function handle_get(PDO $pdo, array $user): void {
  $ownedInvitations = [];
  $receivedInvitations = [];

  $ownedTeams = fetch_owned_team_ids($pdo, (int)$user['id']);
  if ($ownedTeams) {
    $placeholders = implode(',', array_fill(0, count($ownedTeams), '?'));
    $sql = "SELECT i.id, i.team_id, t.name AS team_name, i.email, i.created_at,
                   inviter.username AS invited_by, invited.username AS invited_username,
                   owners.owner_username AS team_owner
            FROM team_invitations i
            JOIN teams t ON t.id = i.team_id
            JOIN users inviter ON inviter.id = i.invited_by
            LEFT JOIN users invited ON invited.id = i.invited_user_id
            LEFT JOIN (
              SELECT m.team_id, MIN(u.username) AS owner_username
              FROM memberships m
              JOIN users u ON u.id = m.user_id
              WHERE LOWER(m.role) = 'owner'
              GROUP BY m.team_id
            ) owners ON owners.team_id = i.team_id
            WHERE i.status = 'pending' AND i.team_id IN ($placeholders)
            ORDER BY i.created_at DESC";
    $st = $pdo->prepare($sql);
    $st->execute($ownedTeams);
    $ownedInvitations = array_map('format_invitation_row', $st->fetchAll());
  }

  $email = strtolower(trim((string)($user['email'] ?? '')));
  if ($email !== '') {
    $sql = "SELECT i.id, i.team_id, t.name AS team_name, i.email, i.created_at,
                   inviter.username AS invited_by, invited.username AS invited_username,
                   owners.owner_username AS team_owner
            FROM team_invitations i
            JOIN teams t ON t.id = i.team_id
            JOIN users inviter ON inviter.id = i.invited_by
            LEFT JOIN users invited ON invited.id = i.invited_user_id
            LEFT JOIN (
              SELECT m.team_id, MIN(u.username) AS owner_username
              FROM memberships m
              JOIN users u ON u.id = m.user_id
              WHERE LOWER(m.role) = 'owner'
              GROUP BY m.team_id
            ) owners ON owners.team_id = i.team_id
            WHERE i.status = 'pending' AND LOWER(i.email) = ?
            ORDER BY i.created_at DESC";
    $st = $pdo->prepare($sql);
    $st->execute([$email]);
    $receivedInvitations = array_map('format_invitation_row', $st->fetchAll());
  }

  echo json_encode([
    'owned' => $ownedInvitations,
    'received' => $receivedInvitations,
  ]);
}

function handle_post(PDO $pdo, array $user): void {
  $raw = file_get_contents('php://input');
  $payload = json_decode($raw, true);

  $email = strtolower(trim((string)($payload['email'] ?? '')));
  if (!filter_var($email, FILTER_VALIDATE_EMAIL)) {
    http_response_code(422);
    echo json_encode(['error' => 'invalid_email']);
    return;
  }

  $teamId = current_team_id($pdo, (int)$user['id']);
  if (!$teamId) {
    http_response_code(409);
    echo json_encode(['error' => 'no_active_team']);
    return;
  }

  if (!user_is_team_owner($pdo, (int)$user['id'], $teamId)) {
    http_response_code(403);
    echo json_encode(['error' => 'forbidden']);
    return;
  }

  $st = $pdo->prepare('SELECT id, username FROM users WHERE LOWER(email)=?');
  $st->execute([$email]);
  $targetUser = $st->fetch();
  if (!$targetUser) {
    http_response_code(404);
    echo json_encode(['error' => 'user_not_found']);
    return;
  }

  $st = $pdo->prepare('SELECT 1 FROM memberships WHERE user_id=? AND team_id=?');
  $st->execute([(int)$targetUser['id'], $teamId]);
  if ($st->fetch()) {
    http_response_code(409);
    echo json_encode(['error' => 'already_member']);
    return;
  }

  $st = $pdo->prepare('SELECT id FROM team_invitations WHERE team_id=? AND LOWER(email)=? AND status=\'pending\'');
  $st->execute([$teamId, $email]);
  if ($st->fetch()) {
    http_response_code(409);
    echo json_encode(['error' => 'already_pending']);
    return;
  }

  $st = $pdo->prepare('INSERT INTO team_invitations (team_id, email, invited_user_id, invited_by) VALUES (?,?,?,?)');
  $st->execute([$teamId, $email, (int)$targetUser['id'], (int)$user['id']]);
  $invitationId = (int)$pdo->lastInsertId();

  $st = $pdo->prepare('SELECT name FROM teams WHERE id=?');
  $st->execute([$teamId]);
  $team = $st->fetch();

  $response = [
    'id' => $invitationId,
    'team_id' => $teamId,
    'team_name' => $team['name'] ?? null,
    'email' => $email,
    'invited_username' => $targetUser['username'] ?? null,
    'invited_by' => $user['username'] ?? null,
    'created_at' => gmdate('c'),
  ];

  echo json_encode($response);
}

function handle_patch(PDO $pdo, array $user): void {
  $raw = file_get_contents('php://input');
  $payload = json_decode($raw, true);

  $invitationId = (int)($payload['invitation_id'] ?? 0);
  $action = strtolower(trim((string)($payload['action'] ?? '')));

  if ($invitationId <= 0 || !in_array($action, ['accept', 'decline'], true)) {
    http_response_code(400);
    echo json_encode(['error' => 'invalid_request']);
    return;
  }

  $st = $pdo->prepare(
    "SELECT i.*, t.name AS team_name, owners.owner_username AS team_owner
     FROM team_invitations i
     JOIN teams t ON t.id = i.team_id
     LEFT JOIN (
       SELECT m.team_id, MIN(u.username) AS owner_username
       FROM memberships m
       JOIN users u ON u.id = m.user_id
       WHERE LOWER(m.role) = 'owner'
       GROUP BY m.team_id
     ) owners ON owners.team_id = i.team_id
     WHERE i.id=?"
  );
  $st->execute([$invitationId]);
  $invitation = $st->fetch();
  if (!$invitation) {
    http_response_code(404);
    echo json_encode(['error' => 'invitation_not_found']);
    return;
  }

  if ($invitation['status'] !== 'pending') {
    http_response_code(409);
    echo json_encode(['error' => 'invitation_closed']);
    return;
  }

  $email = strtolower(trim((string)($user['email'] ?? '')));
  if ($action === 'accept') {
    if ($email === '' || strtolower($invitation['email']) !== $email) {
      http_response_code(403);
      echo json_encode(['error' => 'forbidden']);
      return;
    }

    $pdo->beginTransaction();
    try {
      $upd = $pdo->prepare('UPDATE team_invitations SET status=\'accepted\', responded_at=NOW() WHERE id=?');
      $upd->execute([$invitationId]);

      $check = $pdo->prepare('SELECT 1 FROM memberships WHERE user_id=? AND team_id=?');
      $check->execute([(int)$user['id'], (int)$invitation['team_id']]);
      if (!$check->fetch()) {
        $insert = $pdo->prepare('INSERT INTO memberships (user_id, team_id, role) VALUES (?,?,?)');
        $insert->execute([(int)$user['id'], (int)$invitation['team_id'], 'Member']);
      }

      if (empty($user['active_team_id'])) {
        $pdo->prepare('UPDATE users SET active_team_id=? WHERE id=?')->execute([
          (int)$invitation['team_id'],
          (int)$user['id'],
        ]);
      }

      $pdo->commit();
    } catch (Throwable $e) {
      $pdo->rollBack();
      http_response_code(500);
      echo json_encode(['error' => 'accept_failed']);
      return;
    }

    echo json_encode([
      'ok' => true,
      'action' => 'accept',
      'team' => [
        'id' => (int)$invitation['team_id'],
        'name' => $invitation['team_name'] ?? null,
        'owner' => $invitation['team_owner'] ?? null,
        'role' => 'Membre',
      ],
    ]);
    return;
  }

  if (!user_is_team_owner($pdo, (int)$user['id'], (int)$invitation['team_id']) && $email !== strtolower($invitation['email'])) {
    http_response_code(403);
    echo json_encode(['error' => 'forbidden']);
    return;
  }

  $upd = $pdo->prepare('UPDATE team_invitations SET status=\'declined\', responded_at=NOW() WHERE id=?');
  $upd->execute([$invitationId]);

  echo json_encode(['ok' => true, 'action' => 'decline']);
}

function handle_delete(PDO $pdo, array $user): void {
  $invitationId = (int)($_GET['id'] ?? 0);
  if ($invitationId <= 0) {
    http_response_code(400);
    echo json_encode(['error' => 'invalid_request']);
    return;
  }

  $st = $pdo->prepare('SELECT team_id, status FROM team_invitations WHERE id=?');
  $st->execute([$invitationId]);
  $invitation = $st->fetch();
  if (!$invitation) {
    http_response_code(404);
    echo json_encode(['error' => 'invitation_not_found']);
    return;
  }

  if (!user_is_team_owner($pdo, (int)$user['id'], (int)$invitation['team_id'])) {
    http_response_code(403);
    echo json_encode(['error' => 'forbidden']);
    return;
  }

  if ($invitation['status'] !== 'pending') {
    http_response_code(409);
    echo json_encode(['error' => 'invitation_closed']);
    return;
  }

  $pdo->prepare('DELETE FROM team_invitations WHERE id=?')->execute([$invitationId]);
  echo json_encode(['ok' => true]);
}

function fetch_owned_team_ids(PDO $pdo, int $userId): array {
  $st = $pdo->prepare("SELECT team_id FROM memberships WHERE user_id=? AND LOWER(role)='owner'");
  $st->execute([$userId]);
  return array_map('intval', array_column($st->fetchAll(), 'team_id'));
}

function user_is_team_owner(PDO $pdo, int $userId, int $teamId): bool {
  $st = $pdo->prepare("SELECT 1 FROM memberships WHERE user_id=? AND team_id=? AND LOWER(role)='owner'");
  $st->execute([$userId, $teamId]);
  return (bool)$st->fetch();
}

function format_invitation_row(array $row): array {
  return [
    'id' => isset($row['id']) ? (int)$row['id'] : null,
    'team_id' => isset($row['team_id']) ? (int)$row['team_id'] : null,
    'team_name' => $row['team_name'] ?? null,
    'email' => isset($row['email']) ? strtolower($row['email']) : null,
    'created_at' => isset($row['created_at']) ? gmdate('c', strtotime($row['created_at'])) : null,
    'invited_by' => $row['invited_by'] ?? null,
    'invited_username' => $row['invited_username'] ?? null,
    'team_owner' => $row['team_owner'] ?? null,
  ];
}
