<?php
$databaseDir = __DIR__ . '/Database';
$databasePath = $databaseDir . '/ONESHIP.db';

try {
    // Tạo thư mục Database nếu chưa tồn tại
    if (!file_exists($databaseDir)) {
        mkdir($databaseDir, 0777, true);
    }

    // Kết nối đến SQLite database (tạo mới nếu chưa có)
    $db = new PDO("sqlite:$databasePath");
    $db->setAttribute(PDO::ATTR_ERRMODE, PDO::ERRMODE_EXCEPTION);

    // Câu lệnh SQL tạo bảng ONESHIP
    $sql = "CREATE TABLE IF NOT EXISTS ONESHIP (
        Ma_E1          TEXT    NOT NULL PRIMARY KEY,
        Ngay_Phat_Hanh TEXT,
        KL_Tinh_Cuoc   INTEGER,
        Cuoc_Chinh     INTEGER,
        Nguoi_Nhan     TEXT,
        DCNhan         TEXT,
        Dien_Thoai     TEXT,
        Dich_Vu        TEXT,
        So_Tham_Chieu  TEXT
    );";

    // Thực thi truy vấn tạo bảng
    $db->exec($sql);
    echo "Database ONESHIP.db và bảng ONESHIP đã được tạo thành công!";
} catch (PDOException $e) {
    echo "Lỗi: " . $e->getMessage();
}

// Đóng kết nối
$db = null;
?>