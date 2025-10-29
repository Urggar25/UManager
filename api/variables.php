<?php
require __DIR__.'/config.php';
$user = require_auth($pdo);

ensure_team_variables_schema($pdo);

$teamId = current_team_id($pdo, (int)$user['id']);
if (!$teamId) {
  http_response_code(409);
  echo json_encode(['error' => 'no_active_team']);
  exit;
}

$method = $_SERVER['REQUEST_METHOD'];

if ($method === 'GET') {
  $st = $pdo->prepare('SELECT k, v FROM team_variables WHERE team_id=?');
  $st->execute([$teamId]);
  $rows = $st->fetchAll();
  $out = [];
  foreach ($rows as $r) {
    $out[$r['k']] = json_decode($r['v'], true);
  }
  echo json_encode($out);
  exit;
}

if ($method === 'POST') {
  $raw = file_get_contents('php://input');
  $payload = json_decode($raw, true);
  if (!is_array($payload)) {
    http_response_code(400);
    echo json_encode(['error' => 'invalid_json']);
    exit;
  }

  $pdo->beginTransaction();
  try {
    $ins = $pdo->prepare('INSERT INTO team_variables (team_id, k, v) VALUES (?,?,?)');
    $upd = $pdo->prepare('UPDATE team_variables SET v=? WHERE team_id=? AND k=?');
    foreach ($payload as $k => $val) {
      $json = json_encode($val, JSON_UNESCAPED_UNICODE);
      try {
        $ins->execute([$teamId, $k, $json]);
      } catch (Throwable $e) {
        $upd->execute([$json, $teamId, $k]);
      }
    }
    $pdo->commit();
    echo json_encode(['ok' => true]);
  } catch (Throwable $e) {
    $pdo->rollBack();
    http_response_code(500);
    echo json_encode(['error' => 'write_failed']);
  }
  exit;
}

if ($method === 'DELETE') {
  // /api/variables.php?key=nom_variable
  $key = $_GET['key'] ?? '';
  if ($key === '') {
    http_response_code(400);
    echo json_encode(['error' => 'key_required']);
    exit;
  }
  $st = $pdo->prepare('DELETE FROM team_variables WHERE team_id=? AND k=?');
  $st->execute([$teamId, $key]);
  echo json_encode(['ok' => true]);
  exit;
}

http_response_code(405);
echo json_encode(['error' => 'method_not_allowed']);

function ensure_team_variables_schema(PDO $pdo): void {
  try {
    $pdo->exec(
      "CREATE TABLE IF NOT EXISTS team_variables (
        id INT AUTO_INCREMENT PRIMARY KEY,
        team_id INT NOT NULL,
        k VARCHAR(191) NOT NULL,
        v LONGTEXT NOT NULL,
        updated_at DATETIME NOT NULL DEFAULT CURRENT_TIMESTAMP ON UPDATE CURRENT_TIMESTAMP,
        UNIQUE KEY uniq_team_variable (team_id, k),
        INDEX idx_team_variable_team (team_id),
        CONSTRAINT fk_team_variables_team FOREIGN KEY (team_id) REFERENCES teams(id) ON DELETE CASCADE
      ) ENGINE=InnoDB DEFAULT CHARSET=utf8mb4"
    );
  } catch (Throwable $e) {
    http_response_code(500);
    echo json_encode(['error' => 'schema_setup_failed']);
    exit;
  }
}
