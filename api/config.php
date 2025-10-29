<?php
declare(strict_types=1);

header('Content-Type: application/json; charset=utf-8');

session_set_cookie_params([
  'lifetime' => 0,
  'path' => '/',
  'secure' => isset($_SERVER['HTTPS']),
  'httponly' => true,
  'samesite' => 'Lax',
]);
session_start();

/* ⚠️ REMPLACE ICI PAR TES VRAIS IDENTIFIANTS MYSQL */
const DB_HOST = 'localhost';
const DB_NAME = 'u706630068_up';      // ta BDD affichée à gauche dans phpMyAdmin
const DB_USER = 'u706630068_app';    // ton utilisateur MySQL
const DB_PASS = 'Gf2ds7825';            // son mot de passe

try {
  $pdo = new PDO(
    'mysql:host='.DB_HOST.';dbname='.DB_NAME.';charset=utf8mb4',
    DB_USER, DB_PASS,
    [ PDO::ATTR_ERRMODE => PDO::ERRMODE_EXCEPTION,
      PDO::ATTR_DEFAULT_FETCH_MODE => PDO::FETCH_ASSOC ]
  );
} catch (Throwable $e) {
  http_response_code(500);
  echo json_encode(['error'=>'db_connect_failed']); exit;
}

function require_auth(PDO $pdo): array {
  if (empty($_SESSION['uid'])) {
    http_response_code(401);
    echo json_encode(['error'=>'unauthorized']); exit;
  }
  $st = $pdo->prepare('SELECT * FROM users WHERE id=?');
  $st->execute([$_SESSION['uid']]);
  $u = $st->fetch();
  if (!$u) { session_destroy(); http_response_code(401); echo json_encode(['error'=>'unauthorized']); exit; }
  return $u;
}

function current_team_id(PDO $pdo, int $uid): ?int {
  $st = $pdo->prepare('SELECT active_team_id FROM users WHERE id=?');
  $st->execute([$uid]);
  $row = $st->fetch();
  return $row && $row['active_team_id'] ? (int)$row['active_team_id'] : null;
}
