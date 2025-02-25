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
        $total_imported = 0;
        foreach ($_FILES['excelFiles']['tmp_name'] as $index => $file) {
            $reader = IOFactory::createReaderForFile($file);
            $reader->setReadDataOnly(true);
            $spreadsheet = $reader->load($file);
            $sheet = $spreadsheet->getActiveSheet();
            $title = trim($sheet->getCell('A1')->getValue());
            $data = [];

            switch ($title) {
                case "Mã E1":
                    for ($row = 2; $row <= $sheet->getHighestRow(); $row++) {
                        $Ma_E1 = trim($sheet->getCell("A$row")->getValue());
                        if (!preg_match('/^E.*VN$/', $Ma_E1)) continue;

                        $data[] = [
                            $Ma_E1,
                            date('Y-m-d', strtotime($sheet->getCell("B$row")->getValue())),
                            (int)$sheet->getCell("B$row")->getValue(),
                            (int)str_replace(',', '', $sheet->getCell("F$row")->getValue()),
                            trim($sheet->getCell("M$row")->getValue()),
                            NULL,
                            NULL,
                            null,
                            trim($sheet->getCell("N$row")->getValue())
                        ];
                    }
                    break;
                case "BẢNG TỔNG HỢP NỢ CHI TIẾT":
                    for ($row = 2; $row <= $sheet->getHighestRow(); $row++) {
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
                            trim($sheet->getCell("U$row")->getValue()),
                            trim($sheet->getCell("Q$row")->getValue())
                        ];
                    }
                    break;

                case "TỔNG HỢP SẢN LƯỢNG KHÁCH HÀNG TẠI ĐƠN VỊ":
                    for ($row = 2; $row <= $sheet->getHighestRow(); $row++) {
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
                            null,
                            trim($sheet->getCell("N$row")->getValue())
                        ];
                    }
                    break;
                case "TỔNG HỢP CHUYỂN HOÀN THEO NGÀY":
                    for ($row = 2; $row <= $sheet->getHighestRow(); $row++) {
                        $Ma_E1 = trim($sheet->getCell("A$row")->getValue());
                        if (!preg_match('/^E.*VN$/', $Ma_E1)) continue;

                        $data[] = [
                            $Ma_E1,
                            date('Y-m-d', strtotime($sheet->getCell("L$row")->getValue())),
                            null,
                            (int)str_replace(',', '', $sheet->getCell("E$row")->getValue()),
                            trim($sheet->getCell("N$row")->getValue()),
                            null,
                            null,
                            null,
                            trim($sheet->getCell("O$row")->getValue())
                        ];
                    }
                    break;
                case "TỔNG HỢP CÁC KHÁCH HÀNG SỬ DỤNG DỊCH VỤ ENN VÀ TMD":
                    for ($row = 2; $row <= $sheet->getHighestRow(); $row++) {
                        $Ma_E1 = trim($sheet->getCell("A$row")->getValue());
                        $Ma_E1 = preg_replace('/[^a-zA-Z0-9]/', '', $Ma_E1); // Loại bỏ ký tự ngoài số và chữ
                        if (!preg_match('/^E.*VN$/', $Ma_E1)) continue;
                        $data[] = [
                            $Ma_E1,
                            date('Y-m-d', strtotime($sheet->getCell("E$row")->getValue())),
                            (int)$sheet->getCell("B$row")->getValue(),
                            (int)str_replace(',', '', $sheet->getCell("C$row")->getValue()),
                            null,
                            null,
                            null,
                            null,
                            null
                        ];
                    }
                    break;
                default:
                    echo "<script>
                setTimeout(function() {
                    alert('Tiêu đề bảng không hợp lệ trong file: " . $_FILES['excelFiles']['name'][$index] . "');
                }, 500);
            </script>";
            }

            if (!empty($data)) {
                $stmt = $db->prepare("INSERT INTO ONESHIP 
                    (Ma_E1, Ngay_Phat_Hanh, KL_Tinh_Cuoc, Cuoc_Chinh, Nguoi_Nhan, DCNhan, Dien_Thoai, Dich_Vu, So_Tham_Chieu)
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?)
                    ON CONFLICT(Ma_E1) DO UPDATE SET 
                    Ngay_Phat_Hanh=excluded.Ngay_Phat_Hanh, 
                    KL_Tinh_Cuoc=excluded.KL_Tinh_Cuoc, 
                    Cuoc_Chinh=excluded.Cuoc_Chinh, 
                    Nguoi_Nhan=excluded.Nguoi_Nhan, 
                    DCNhan=excluded.DCNhan, 
                    Dien_Thoai=excluded.Dien_Thoai, 
                    Dich_Vu=excluded.Dich_Vu, 
                    So_Tham_Chieu=excluded.So_Tham_Chieu");

                
                foreach ($data as $row) {
                    $stmt->reset();
                    for ($i = 0; $i < count($row); $i++) {
                        $stmt->bindValue($i + 1, $row[$i]);
                    }
                    $stmt->execute();
                }
                $total_imported += count($data); 
            }
        }
        $db->exec("COMMIT;");
        echo "<script>
            setTimeout(function() {
                alert('Nhập dữ liệu thành công! Tổng số bản ghi đã nh: " . $total_imported . "');
            }, 500);
        </script>";
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
    $colCuocChinh = 'J';
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


$limit = 1000;
$page = isset($_GET['page']) ? max(1, intval($_GET['page'])) : 1;
$offset = ($page - 1) * $limit;


$search = isset($_GET['search']) ? trim($_GET['search']) : '';

$whereClause = "";
$params = [];
if ($search) {
    $whereClause = "WHERE Ma_E1 LIKE :search OR Ngay_Phat_Hanh LIKE :search";
    $params[':search'] = "%$search%";
}


$queryCount = "SELECT COUNT(*) FROM ONESHIP $whereClause";
$stmt = $db->prepare($queryCount);
foreach ($params as $key => $value) {
    $stmt->bindValue($key, $value, SQLITE3_TEXT);
}
$totalRows = $stmt->execute()->fetchArray()[0];
$totalPages = ceil($totalRows / $limit);


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
        <div class="upload-section">
            <form action="index.php" method="post" enctype="multipart/form-data">
                <input type="file" name="excelFiles[]" multiple class="file-input">
                <button type="submit" name="submit" class="btn">Nhập phí</button>
            </form>

            <form action="index.php" method="post" enctype="multipart/form-data">
                <input type="file" name="file" accept=".xls,.xlsx" required class="file-input">
                <button type="submit" name="submit" class="btn">Upload và Xử lý COD</button>
            </form>
        </div>
        <?php if (!empty($file_download)) : ?>
            <p class="download-link"><a href="<?php echo $file_download; ?>" download>Tải xuống file kết quả</a></p>
        <?php endif; ?>

        <h2 class="table-title">Danh Sách GOSHIP</h2>

        <div class="search-total-container">
            <div class="total-count">
                <strong>Tổng số lượng:</strong> <?= number_format($totalRows) ?>
            </div>
            <div class="search-container">
                <form method="GET">
                    <input type="text" name="search" value="<?= htmlspecialchars($search) ?>" placeholder="Nhập Ma_E1 hoặc Ngày Phát Hành..." class="search-input">
                    <button type="submit" class="btn">Tìm kiếm</button>
                    <?php if ($search): ?>
                        <a href="index.php"><button type="button" class="btn cancel">Xóa tìm kiếm</button></a>
                    <?php endif; ?>
                </form>
            </div>
        </div>
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
                            <th>Dịch vụ</th>
                            <th>Số tham chiếu</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php
                        $i = ($page - 1) * $limit + 1; 
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
                                <td><?= htmlspecialchars($row['So_Tham_Chieu']) ?></td>

                            </tr>
                        <?php endwhile; ?>
                    </tbody>
                </table>
            <?php else: ?>
                <p class="no-data">Không tìm thấy kết quả!</p>
            <?php endif; ?>
        </div>

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