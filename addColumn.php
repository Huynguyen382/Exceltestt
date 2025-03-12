<?php
$databaseDir = __DIR__ . '/Database';
$databasePath = $databaseDir . '/ONESHIP.db';

    if (!file_exists($databaseDir)) {
        mkdir($databaseDir, 0777, true);
    }

    $db = new PDO("sqlite:$databasePath");
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $sql = "ALTER TABLE ONESHIP ADD COLUMN Ten_File TEXT;

    );";

    $db->exec($sql);
    echo "Them cot Ten_File";

$db = null;
?>