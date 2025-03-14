<?php
$databaseDir = __DIR__ . '/Database';
$databasePath = $databaseDir . '/ONESHIP.db';

try {
    if (!file_exists($databaseDir)) {
        mkdir($databaseDir, 0777, true);
    }

    $db = new PDO("sqlite:$databasePath");
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    $sql = "CREATE TABLE IF NOT EXISTS tbl_VNPOST (
        Ma_E1          TEXT    NOT NULL PRIMARY KEY,
        Ngay_Phat_Hanh TEXT,
        KL_Tinh_Cuoc   INTEGER,
        Cuoc_Chinh     INTEGER,
        Nguoi_Nhan     TEXT,
        DCNhan         TEXT,
        Dien_Thoai     TEXT,
        Dich_Vu        TEXT,
        So_Tham_Chieu  TEXT,
        Ten_File       TEXT
    );";

    $db->exec($sql);
    echo "Bảng tbl_VNPOST đã được tạo thành công!";
} catch (PDOException $e) {
    echo "Lỗi: " . $e->getMessage();
}

$db = null;
?>