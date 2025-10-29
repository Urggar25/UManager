<?php
require __DIR__.'/config.php';

$out = [];

// Quelle base est utilisée ?
$out['database'] = $pdo->query('SELECT DATABASE() AS db')->fetch()['db'] ?? null;

// Combien d'utilisateurs ?
$out['users_count'] = (int)$pdo->query('SELECT COUNT(*) FROM users')->fetchColumn();

// L'utilisateur admin existe ?
$st = $pdo->prepare('SELECT id, email, username, active_team_id FROM users WHERE email=?');
$st->execute(['admin@example.com']);
$out['admin_user'] = $st->fetch() ?: null;

header('Content-Type: application/json; charset=utf-8');
echo json_encode($out, JSON_PRETTY_PRINT|JSON_UNESCAPED_UNICODE);
