<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Cell\Coordinate;
use PhpOffice\PhpSpreadsheet\Shared\Date;

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

function excelDateToPHP($excelDate)
{
    if (is_numeric($excelDate)) {
        if ($excelDate > 50000) {
            return $excelDate;
        }
        return date('Y-m-d', strtotime("1899-12-30 +$excelDate days"));
    }
    return $excelDate;
}
function cleanExcelValue($value)
{
    return preg_replace('/^="(.*)"$/', '$1', trim($value));
}

if (isset($_POST['submit']) || isset($_POST['uploadEX'])) {
    if (!file_exists('uploads')) {
        mkdir('uploads', 0777, true);
    }

    if (!empty($_FILES['excelFiles']['tmp_name']) && is_array($_FILES['excelFiles']['tmp_name'])) {
        $db->exec("BEGIN TRANSACTION;");
        $totalImported = 0;

        foreach ($_FILES['excelFiles']['tmp_name'] as $index => $file) {
            $reader = IOFactory::createReaderForFile($file);
            $reader->setReadDataOnly(true);
            $spreadSheet = $reader->load($file);
            $sheet = $spreadSheet->getActiveSheet();

            $headerRow = null;
            for ($row = 1; $row <= 5; $row++) {
                $cellValue = cleanExcelValue($sheet->getCell("A$row")->getValue());
                if (preg_match('/^(Mã E1|Ma_E1)$/i', $cellValue)) {
                    $headerRow = $row;
                    break;
                }
            }

            if (!$headerRow) {
                echo "<script>alert('Không tìm thấy tiêu đề Mã E1 trong file: " . $_FILES['excelFiles']['name'][$index] . "');</script>";
                continue;
            }

            $columns = [];
            $highestColumnIndex = Coordinate::columnIndexFromString($sheet->getHighestColumn());

            for ($col = 1; $col <= $highestColumnIndex; $col++) {
                $colLetter = Coordinate::stringFromColumnIndex($col);
                $colValue = cleanExcelValue($sheet->getCell("$colLetter$headerRow")->getValue());

                if (preg_match('/^(Mã E1|Ma_E1)$/i', $colValue)) {
                    $columns['Ma_E1'] = $col;
                } elseif (preg_match('/^(Ngày Đóng|Ngay_Phat_Hanh|Ngay_Dong)$/i', $colValue)) {
                    $columns['Ngay_Phat_Hanh'] = $col;
                } elseif (preg_match('/^(Khối lượng|Khoi_Luong|KL_Tinh_Cuoc|Khối Lượng \(gam\))$/i', $colValue)) {
                    $columns['KL_Tinh_Cuoc'] = $col;
                } elseif (preg_match('/^(Cuoc_E1|Cước E1|Cước E1 \(VNĐ\))$/i', $colValue)) {
                    $columns['Cuoc_Chinh'] = $col;
                } elseif (preg_match('/^(Cước Chính|Cuoc_Chinh)$/i', $colValue) && !isset($columns['Cuoc_Chinh'])) {
                    $columns['Cuoc_Chinh'] = $col;
                } elseif (preg_match('/^(Nguoi_Nhan|Người Nhận)$/i', $colValue)) {
                    $columns['Nguoi_Nhan'] = $col;
                } elseif (preg_match('/^(DC_Nhan|DCNhan)$/i', $colValue)) {
                    $columns['DCNhan'] = $col;
                } elseif (preg_match('/^(Dien_Thoai|Dien_Thoai_Nhan)$/i', $colValue)) {
                    $columns['Dien_Thoai'] = $col;
                } elseif (preg_match('/^(So_Tham_Chieu)$/i', $colValue)) {
                    $columns['So_Tham_Chieu'] = $col;
                }
            }

            if (!isset($columns['Ma_E1'])) {
                echo "<script>alert('Không tìm thấy cột Mã E1 trong file: " . $_FILES['excelFiles']['name'][$index] . "');</script>";
                continue;
            }

            $data = [];
            for ($row = $headerRow + 1; $row <= $sheet->getHighestRow(); $row++) {
                $Ma_E1 = cleanExcelValue($sheet->getCell(Coordinate::stringFromColumnIndex($columns['Ma_E1']) . $row)->getValue());
                if (!preg_match('/^E.*VN$/', $Ma_E1)) {
                    continue;
                }

                $rowData = [
                    'Ma_E1'          => $Ma_E1,
                    'Ngay_Phat_Hanh' => isset($columns['Ngay_Phat_Hanh']) ? cleanExcelValue(excelDateToPHP($sheet->getCell(Coordinate::stringFromColumnIndex($columns['Ngay_Phat_Hanh']) . $row)->getValue())) : null,
                    'KL_Tinh_Cuoc'   => isset($columns['KL_Tinh_Cuoc']) ? (int)$sheet->getCell(Coordinate::stringFromColumnIndex($columns['KL_Tinh_Cuoc']) . $row)->getValue() : null,
                    'Cuoc_Chinh'     => isset($columns['Cuoc_Chinh']) ? (int)str_replace(',', '', $sheet->getCell(Coordinate::stringFromColumnIndex($columns['Cuoc_Chinh']) . $row)->getCalculatedValue()) : null,
                    'Nguoi_Nhan'     => isset($columns['Nguoi_Nhan']) ? cleanExcelValue($sheet->getCell(Coordinate::stringFromColumnIndex($columns['Nguoi_Nhan']) . $row)->getValue()) : null,
                    'DCNhan'         => isset($columns['DCNhan']) ? cleanExcelValue($sheet->getCell(Coordinate::stringFromColumnIndex($columns['DCNhan']) . $row)->getValue()) : null,
                    'Dien_Thoai'     => isset($columns['Dien_Thoai']) ? cleanExcelValue($sheet->getCell(Coordinate::stringFromColumnIndex($columns['Dien_Thoai']) . $row)->getValue()) : null,
                    'So_Tham_Chieu'  => isset($columns['So_Tham_Chieu']) ? cleanExcelValue($sheet->getCell(Coordinate::stringFromColumnIndex($columns['So_Tham_Chieu']) . $row)->getValue()) : null,
                    'Ten_File'       => $_FILES['excelFiles']['name'][$index],
                ];
                $data[] = $rowData;
            }

            if (!empty($data)) {
                $stmt = $db->prepare("INSERT INTO ONESHIP (Ma_E1, Ngay_Phat_Hanh, KL_Tinh_Cuoc, Cuoc_Chinh, Nguoi_Nhan, DCNhan, Dien_Thoai, So_Tham_Chieu, Ten_File) 
                    VALUES (?, ?, ?, ?, ?, ?, ?, ?, ?) 
                    ON CONFLICT(Ma_E1) DO UPDATE SET 
                    Ngay_Phat_Hanh=excluded.Ngay_Phat_Hanh, KL_Tinh_Cuoc=excluded.KL_Tinh_Cuoc, 
                    Cuoc_Chinh=excluded.Cuoc_Chinh, Nguoi_Nhan=excluded.Nguoi_Nhan, 
                    DCNhan=excluded.DCNhan, Dien_Thoai=excluded.Dien_Thoai, So_Tham_Chieu=excluded.So_Tham_Chieu, Ten_File=excluded.Ten_File");

                foreach ($data as $row) {
                    $stmt->bindValue(1, $row['Ma_E1'], SQLITE3_TEXT);
                    $stmt->bindValue(2, $row['Ngay_Phat_Hanh'], SQLITE3_TEXT);
                    $stmt->bindValue(3, $row['KL_Tinh_Cuoc'], SQLITE3_INTEGER);
                    $stmt->bindValue(4, $row['Cuoc_Chinh'], SQLITE3_INTEGER);
                    $stmt->bindValue(5, $row['Nguoi_Nhan'], SQLITE3_TEXT);
                    $stmt->bindValue(6, $row['DCNhan'], SQLITE3_TEXT);
                    $stmt->bindValue(7, $row['Dien_Thoai'], SQLITE3_TEXT);
                    $stmt->bindValue(8, $row['So_Tham_Chieu'], SQLITE3_TEXT);
                    $stmt->bindValue(9, $row['Ten_File'], SQLITE3_TEXT);

                    $stmt->execute();
                }
                $totalImported += count($data);
            }
        }
        $db->exec("COMMIT;");
        echo "<script>alert('Nhập dữ liệu thành công! Tổng số bản ghi: $totalImported');</script>";
    }
}
$fileDownload = "";
if (isset($_POST['submit']) && isset($_FILES['file']) && $_FILES['file']['error'] == 0) {
    $fileTmpPath = $_FILES['file']['tmp_name'];
    $fileName = pathinfo($_FILES['file']['name'], PATHINFO_FILENAME);
    $timeStamp = date('Ymd_His');
    $outputFile = "uploads/{$fileName}_{$timeStamp}.xlsx";
    $spreadSheet = IOFactory::load($fileTmpPath);
    $workSheet = $spreadSheet->getActiveSheet();
    $highestRow = $workSheet->getHighestRow();
    $highestColumnIndex = Coordinate::columnIndexFromString($workSheet->getHighestColumn());
    $headerRow = null;
    $headerCol = null;
    for ($row = 1; $row <= 5; $row++) {
        for ($col = 1; $col <= $highestColumnIndex; $col++) {
            $colLetter = Coordinate::stringFromColumnIndex($col);
            $cellValue = cleanExcelValue($workSheet->getCell("$colLetter$row")->getValue());

            if (preg_match('/^E.*VN$/', $cellValue)) {
                $headerRow = $row;
                $headerCol = $col;
                break 2;
            }
        }
    }


    if ($headerCol === null) {
        echo "<script>alert('Không tìm thấy cột chứa Mã E1 trong file Excel.');</script>";
    } else {
        $colMaE1Letter = Coordinate::stringFromColumnIndex($headerCol);
        $colCuocChinhIndex = $highestColumnIndex + 1;
        $colCuocChinhLetter = Coordinate::stringFromColumnIndex($colCuocChinhIndex);
        $workSheet->setCellValue($colCuocChinhLetter . '1', "Cuoc_Chinh");

        for ($row = 2; $row <= $highestRow; $row++) {
            $ma_e1 = trim($workSheet->getCell($colMaE1Letter . $row)->getValue() ?? '');
            if (!empty($ma_e1)) {
                $query = "SELECT Cuoc_Chinh FROM ONESHIP WHERE Ma_E1 = :ma_e1";
                $stmt = $db->prepare($query);
                $stmt->bindValue(':ma_e1', $ma_e1, SQLITE3_TEXT);
                $result = $stmt->execute();
                if ($rowData = $result->fetchArray(SQLITE3_ASSOC)) {
                    $workSheet->setCellValue($colCuocChinhLetter . $row, $rowData['Cuoc_Chinh']);
                }
            }
        }
    }

    $writer = new Xlsx($spreadSheet);
    $writer->save($outputFile);
    $fileDownload = $outputFile;
    if ($fileDownload) {
        echo "<script>
        alert('Xử lý thành công! File sẽ được tải xuống tự động.');
        window.location.href = '$fileDownload';
         setTimeout(function() {
            window.location.href = 'index.php';
        },1000);
    </script>";
    }
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
    <title>Danh sách Goship</title>
    <link rel="stylesheet" type="text/css" href="style.css">
    <script>
        function validateFileInput(event, inputId, tooltipId) {
            let fileInput = document.getElementById(inputId);
            let tooltip = document.getElementById(tooltipId);

            if (fileInput.files.length === 0) {
                tooltip.classList.add("show-tooltip");
                event.preventDefault();
                return false;
            } else {
                tooltip.classList.remove("show-tooltip");
                return true;
            }
        }
    </script>
</head>

<body>

    <div class="container">
        <div class="upload-section">
            <form action="vnpost.php" method="post" enctype="multipart/form-data" class="upload-form">
                <div class="upload-group">
                    <input type="file" name="excelFiles[]" multiple class="file-input" id="fileInput1">
                    <span class="tooltip" id="tooltip1">Please select a file</span>
                </div>
                <button type="submit" name="submit" class="btn" onclick="return validateFileInput(event, 'fileInput1', 'tooltip1')">Nhập phí</button>
            </form>

            <form action="vnpost.php" method="post" enctype="multipart/form-data" class="upload-form">
                <div class="upload-group">
                    <input type="file" name="file" accept=".xls,.xlsx" required class="file-input" id="fileInput2">
                    <span class="tooltip" id="tooltip2">Please select a file</span>
                </div>
                <button type="submit" name="submit" class="btn" onclick="return validateFileInput(event, 'fileInput2', 'tooltip2')">Upload và Xử lý COD</button>
            </form>

            <a href="index.php">Danh sách EMS</a>
            <a href="vnpost.php">Danh sách VNPOST</a>

        </div>

        <?php if (!empty($fileDownload)) : ?>
            <p class="download-link"><a href="<?php echo $fileDownload; ?>" download>Tải xuống file kết quả</a></p>
        <?php endif; ?>
        <h2 class="table-title">Danh Sách EMS</h2>

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
                            <th>Nhập từ bảng</th>
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
                                <td><?= htmlspecialchars($row['Ten_File']) ?></td>
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