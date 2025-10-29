<?php
require __DIR__.'/config.php';

session_destroy();
echo json_encode(['ok' => true]);
