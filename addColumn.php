<?php
$databaseDir = __DIR__ . '/Database';
$databasePath = $databaseDir . '/ONESHIP.db';

// Tạo thư mục nếu chưa tồn tại
if (!file_exists($databaseDir)) {
    mkdir($databaseDir, 0777, true);
}

try {
    $db = new PDO("sqlite:$databasePath");
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $checkColumn = $db->query("PRAGMA table_info(ONESHIP)");
    $columnExists = false;
    foreach ($checkColumn as $column) {
        if ($column['name'] === 'Ten_File') {
            $columnExists = true;
            break;
        }
    }

    if (!$columnExists) {
        $sql = "ALTER TABLE ONESHIP ADD COLUMN Ten_File TEXT;";
        $db->exec($sql);
        echo "Đã thêm cột 'Ten_File' vào bảng ONESHIP.";
    } else {
        echo "Cột 'Ten_File' đã tồn tại, không cần thêm.";
    }

    $db = null; 
} catch (PDOException $e) {
    echo "Lỗi: " . $e->getMessage();
}
?>