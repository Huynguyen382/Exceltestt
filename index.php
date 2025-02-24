<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;

ini_set('max_execution_time', 600); 
ini_set('memory_limit', '512M'); 
ini_set('max_file_uploads', 100);
ini_set('upload_max_filesize', '200M');
ini_set('post_max_size', '500M');


$db = new SQLite3(__DIR__ . '/Database/ONESHIP.db');
$db->exec("PRAGMA synchronous = OFF;");
$db->exec("PRAGMA journal_mode = WAL;");
$db->exec("PRAGMA temp_store = MEMORY;");
$db->exec("PRAGMA cache_size = 1000000;");

if (isset($_POST['submit']) || isset($_POST['uploadEX'])) {
    if (!file_exists('uploads')) {
        mkdir('uploads', 0777, true);
    }

    if (!empty($_FILES['excelFiles']['tmp_name']) && is_array($_FILES['excelFiles']['tmp_name'])) {
        $db->exec("BEGIN TRANSACTION;");

        foreach ($_FILES['excelFiles']['tmp_name'] as $index => $file) {
            $reader = IOFactory::createReaderForFile($file);
            $reader->setReadDataOnly(true);
            $spreadsheet = $reader->load($file);
            $sheet = $spreadsheet->getActiveSheet();
            $title = trim($sheet->getCell('A1')->getValue());
            $data = [];

            switch ($title) {
                case "BẢNG TỔNG HỢP NỢ CHI TIẾT":
                    for ($row = 3; $row <= $sheet->getHighestRow(); $row++) {
                        $Ma_E1 = trim($sheet->getCell("A$row")->getValue());
                        if (!preg_match('/^E.*VN$/', $Ma_E1)) continue;

                        $data[] = [
                            $Ma_E1,
                            date('Y-m-d', strtotime($sheet->getCell("B$row")->getValue())),
                            (int)$sheet->getCell("F$row")->getValue(),
                            (int)str_replace(',', '', $sheet->getCell("I$row")->getValue()),
                            trim($sheet->getCell("L$row")->getValue()),
                            trim($sheet->getCell("M$row")->getValue()),
                            trim($sheet->getCell("N$row")->getValue()),
                            null
                        ];
                    }
                    break;

                case "TỔNG HỢP SẢN LƯỢNG KHÁCH HÀNG TẠI ĐƠN VỊ":
                    for ($row = 3; $row <= $sheet->getHighestRow(); $row++) {
                        $Ma_E1 = trim($sheet->getCell("A$row")->getValue());
                        if (!preg_match('/^E.*VN$/', $Ma_E1)) continue;

                        $data[] = [
                            $Ma_E1,
                            date('Y-m-d', strtotime($sheet->getCell("K$row")->getValue())),
                            (int)$sheet->getCell("B$row")->getValue(),
                            (int)str_replace(',', '', $sheet->getCell("F$row")->getValue()),
                            trim($sheet->getCell("M$row")->getValue()),
                            null,
                            null,
                            null
                        ];
                    }
                    break;

                default:
                    die("Tiêu đề bảng không hợp lệ trong file: " . $_FILES['excelFiles']['name'][$index]);
            }

            // Nếu có dữ liệu, thực hiện ghi hàng loạt
            if (!empty($data)) {
                $stmt = $db->prepare("INSERT INTO ONESHIP 
                    (Ma_E1, Ngay_Phat_Hanh, KL_Tinh_Cuoc, Cuoc_Chinh, Nguoi_Nhan, DCNhan, Dien_Thoai, Dich_Vu) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(Ma_E1) DO UPDATE SET 
                    Ngay_Phat_Hanh=excluded.Ngay_Phat_Hanh, 
                    KL_Tinh_Cuoc=excluded.KL_Tinh_Cuoc, 
                    Cuoc_Chinh=excluded.Cuoc_Chinh, 
                    Nguoi_Nhan=excluded.Nguoi_Nhan, 
                    DCNhan=excluded.DCNhan, 
                    Dien_Thoai=excluded.Dien_Thoai, 
                    Dich_Vu=excluded.Dich_Vu");

                // Lặp qua từng dòng, bind giá trị và thực thi
                foreach ($data as $row) {
                    $stmt->reset();
                    for ($i = 0; $i < count($row); $i++) {
                        $stmt->bindValue($i + 1, $row[$i]);
                    }
                    $stmt->execute();
                }
            }
        }

        $db->exec("COMMIT;");
        echo "Nhập dữ liệu thành công!";
    }
}



$file_download = "";

if (isset($_POST['submit']) && isset($_FILES['file']) && $_FILES['file']['error'] == 0) {
    $file_tmp_path = $_FILES['file']['tmp_name'];
    $file_name = pathinfo($_FILES['file']['name'], PATHINFO_FILENAME);
    $timestamp = date('Ymd_His');
    $output_file = "uploads/{$file_name}_{$timestamp}.xlsx";

    $spreadsheet = IOFactory::load($file_tmp_path);
    $worksheet = $spreadsheet->getActiveSheet();
    $highestRow = $worksheet->getHighestRow();
    $highestColumnIndex = Coordinate::columnIndexFromString($worksheet->getHighestColumn());

    $colCuocChinh = Coordinate::stringFromColumnIndex($highestColumnIndex -2);
    $worksheet->setCellValue($colCuocChinh . '1', "Cuoc_Chinh");

    for ($row = 2; $row <= $highestRow; $row++) {
        $ma_e1 = trim($worksheet->getCell('C' . $row)->getValue() ?? '');
        if (!empty($ma_e1)) {
            $query = "SELECT Cuoc_Chinh FROM ONESHIP WHERE Ma_E1 = :ma_e1";
            $stmt = $db->prepare($query);
            $stmt->bindValue(':ma_e1', $ma_e1, SQLITE3_TEXT);
            $result = $stmt->execute();
            if ($row_data = $result->fetchArray(SQLITE3_ASSOC)) {
                $worksheet->setCellValue($colCuocChinh . $row, $row_data['Cuoc_Chinh']);
            }
        }
    }

    $writer = new Xlsx($spreadsheet);
    $writer->save($output_file);
    $file_download = $output_file;
}
// Xác định số dòng mỗi trang
$limit = 1000;
$page = isset($_GET['page']) ? max(1, intval($_GET['page'])) : 1;
$offset = ($page - 1) * $limit;

// Lấy từ khóa tìm kiếm (nếu có)
$search = isset($_GET['search']) ? trim($_GET['search']) : '';

// Truy vấn dữ liệu có tìm kiếm
$whereClause = "";
$params = [];
if ($search) {
    $whereClause = "WHERE Ma_E1 LIKE :search OR Ngay_Phat_Hanh LIKE :search";
    $params[':search'] = "%$search%";
}

// Đếm tổng số dòng để phân trang
$queryCount = "SELECT COUNT(*) FROM ONESHIP $whereClause";
$stmt = $db->prepare($queryCount);
foreach ($params as $key => $value) {
    $stmt->bindValue($key, $value, SQLITE3_TEXT);
}
$totalRows = $stmt->execute()->fetchArray()[0];
$totalPages = ceil($totalRows / $limit);

// Truy vấn dữ liệu có giới hạn phân trang
$query = "SELECT * FROM ONESHIP $whereClause ORDER BY Ngay_Phat_Hanh DESC LIMIT $limit OFFSET $offset";
$stmt = $db->prepare($query);
foreach ($params as $key => $value) {
    $stmt->bindValue($key, $value, SQLITE3_TEXT);
}
$result = $stmt->execute();


?>
<!DOCTYPE html>
<html lang="vi">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Upload Excel</title>
    <link rel="stylesheet" type="text/css" href="style.css">
</head>

<body>

    <div class="container">
        <!-- Khu vực upload -->
        <div class="upload-section">
            <form action="index.php" method="post" enctype="multipart/form-data">
                <input type="file" name="excelFiles[]" multiple class="file-input">
                <button type="submit" name="submit" class="btn">Import Data</button>
            </form>

            <form action="index.php" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".xls,.xlsx" required class="file-input">
                <button type="submit" name="submit" class="btn">Upload và Xử lý</button>
            </form>
        </div>

        <!-- Hiển thị file tải về -->
        <?php if (!empty($file_download)) : ?>
            <p class="download-link"><a href="<?php echo $file_download; ?>" download>Tải xuống file kết quả</a></p>
        <?php endif; ?>

        <h2 class="table-title">Danh Sách ONESHIP</h2>

        <!-- Thanh tìm kiếm -->
        <div class="search-container">
            <form method="GET">
                <input type="text" name="search" value="<?= htmlspecialchars($search) ?>" placeholder="Nhập Ma_E1 hoặc Ngày Phát Hành..." class="search-input">
                <button type="submit" class="btn">Tìm kiếm</button>
                <?php if ($search): ?>
                    <a href="Test2.php"><button type="button" class="btn cancel">Xóa tìm kiếm</button></a>
                <?php endif; ?>
            </form>
        </div>

        <!-- Khu vực bảng với thanh cuộn -->
        <div class="table-container">
            <?php if ($totalRows > 0): ?>
                <table class="styled-table">
                    <thead>
                        <tr>
                            <th>STT</th>
                            <th>Ma_E1</th>
                            <th>Ngày Phát Hành</th>
                            <th>KL Tính Cước</th>
                            <th>Cước Chính</th>
                            <th>Người Nhận</th>
                            <th>Địa Chỉ Nhận</th>
                            <th>Điện Thoại</th>
                            <th>Dịch Vụ</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php
                        $i = ($page - 1) * $limit + 1; // Tính số thứ tự theo trang
                        while ($row = $result->fetchArray(SQLITE3_ASSOC)):
                        ?>
                            <tr>
                                <td><?= $i++ ?></td>
                                <td><?= htmlspecialchars($row['Ma_E1']) ?></td>
                                <td><?= htmlspecialchars($row['Ngay_Phat_Hanh']) ?></td>
                                <td><?= htmlspecialchars($row['KL_Tinh_Cuoc']) ?></td>
                                <td><?= htmlspecialchars($row['Cuoc_Chinh']) ?></td>
                                <td><?= htmlspecialchars($row['Nguoi_Nhan']) ?></td>
                                <td><?= htmlspecialchars($row['DCNhan']) ?></td>
                                <td><?= htmlspecialchars($row['Dien_Thoai']) ?></td>
                                <td><?= htmlspecialchars($row['Dich_Vu']) ?></td>
                            </tr>
                        <?php endwhile; ?>
                    </tbody>
                </table>
            <?php else: ?>
                <p class="no-data">Không tìm thấy kết quả!</p>
            <?php endif; ?>
        </div>

        <!-- Phân trang -->
        <div class="pagination">
            <?php if ($page > 1): ?>
                <a href="?search=<?= urlencode($search) ?>&page=<?= $page - 1 ?>" class="btn">&lt; Prev</a>
            <?php endif; ?>

            <span>Trang <?= $page ?> / <?= $totalPages ?></span>

            <?php if ($page < $totalPages): ?>
                <a href="?search=<?= urlencode($search) ?>&page=<?= $page + 1 ?>" class="btn">Next &gt;</a>
            <?php endif; ?>
        </div>

    </div>

</body>

</html>