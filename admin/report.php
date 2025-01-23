<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Writer\Xlsx;
use PhpOffice\PhpSpreadsheet\Style\Alignment;
use PhpOffice\PhpSpreadsheet\Style\Border;

// Database connection
$host = '127.0.0.1';
$db = 'isadfc';
$user = 'root';
$pass = '';

try {
    $mysqli = new mysqli($host, $user, $pass, $db);
    if ($mysqli->connect_error) {
        throw new Exception("Connection failed: " . $mysqli->connect_error);
    }

    // Output directory setup
    $outputDir = 'C:/xampp/htdocs/roomasm-master/output/';
    if (!is_dir($outputDir)) {
        mkdir($outputDir, 0777, true);
    }
    if (!is_writable($outputDir)) {
        throw new Exception("Error: The output directory is not writable.");
    }

    $outputPath = $outputDir . 'master_schedule.xlsx';
    $spreadsheet = new Spreadsheet();

    // Get all distinct courses
    $courseQuery = "SELECT DISTINCT course FROM courses ORDER BY course";
    $courseResult = $mysqli->query($courseQuery);

    $sheetIndex = 0;
    while ($courseRow = $courseResult->fetch_assoc()) {
        $course = $courseRow['course'];

        // Create a new sheet for each course
        if ($sheetIndex > 0) {
            $sheet = $spreadsheet->createSheet($sheetIndex);
        } else {
            $sheet = $spreadsheet->getActiveSheet();
        }
        $sheet->setTitle($course);
        $spreadsheet->setActiveSheetIndex($sheetIndex);
        $sheetIndex++;

        // Add header information
        $sheet->setCellValue('A1', 'ASIAN DEVELOPMENT FOUNDATION COLLEGE');
        $sheet->setCellValue('A2', 'Tacloban City');
        $sheet->setCellValue('A4', 'MASTER SCHEDULE');
        $sheet->setCellValue('A5', 'Computer Studies & Engineering Department');
        $sheet->setCellValue('A6', '2nd Semester, S.Y. 2024-2025');

        // Format headers
        foreach (['A1:E1', 'A2:E2', 'A4:E4', 'A5:E5', 'A6:E6'] as $range) {
            $sheet->mergeCells($range);
            $sheet->getStyle($range)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
            $sheet->getStyle($range)->getFont()->setBold(true);
        }

        $currentRow = 8; // Start after headers

        // Get blocks for this course
        $blockQuery = "SELECT DISTINCT block FROM courses WHERE course = ? ORDER BY block";
        $blockStmt = $mysqli->prepare($blockQuery);
        $blockStmt->bind_param('s', $course);
        $blockStmt->execute();
        $blockResult = $blockStmt->get_result();

        while ($blockRow = $blockResult->fetch_assoc()) {
            $block = $blockRow['block'];

            // Add block header
            $sheet->setCellValue("A$currentRow", strtoupper("$course - $block"));
            $sheet->mergeCells("A$currentRow:E$currentRow");
            $sheet->getStyle("A$currentRow")->getFont()->setBold(true);
            $sheet->getStyle("A$currentRow")->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);

            // Add table headers
            $currentRow++;
            $headers = ['CODE', 'SUBJECTS', 'TIME', 'DAY/S', 'ROOM'];
            $sheet->fromArray($headers, NULL, "A$currentRow");

            // Format column headers
            $headerRange = "A$currentRow:E$currentRow";
            $sheet->getStyle($headerRange)->applyFromArray([
                'font' => ['bold' => true],
                'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
                'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]
            ]);

            // Get subjects for this block without sorting by subject
            $subjectQuery = "SELECT 
                course_code,
                subject_code,
                CONCAT(TIME_FORMAT(start_time, '%H:%i'), '-', TIME_FORMAT(end_time, '%H:%i')) as time,
                days,
                room
            FROM courses 
            WHERE course = ? AND block = ?";

            $stmt = $mysqli->prepare($subjectQuery);
            $stmt->bind_param('ss', $course, $block);
            $stmt->execute();
            $result = $stmt->get_result();

            // Add subject data
            while ($row = $result->fetch_assoc()) {
                $currentRow++;
                $data = [
                    $row['course_code'],
                    $row['subject_code'],
                    $row['time'],
                    $row['days'],
                    $row['room']
                ];
                $sheet->fromArray([$data], NULL, "A$currentRow");

                // Format data cells
                $sheet->getStyle("A$currentRow:E$currentRow")->applyFromArray([
                    'borders' => ['allBorders' => ['borderStyle' => Border::BORDER_THIN]],
                    'alignment' => ['horizontal' => Alignment::HORIZONTAL_CENTER]
                ]);
            }

            $currentRow += 2; // Add space between blocks
            $stmt->close();
        }
        $blockStmt->close();

        // Auto-size columns
        foreach (range('A', 'E') as $col) {
            $sheet->getColumnDimension($col)->setAutoSize(true);
        }

        // Set row height for header rows
        $sheet->getRowDimension(1)->setRowHeight(30);
        $sheet->getRowDimension(4)->setRowHeight(30);
    }

    // Set the first sheet as active
    $spreadsheet->setActiveSheetIndex(0);

    // Save the file
    $writer = new Xlsx($spreadsheet);
    $writer->save($outputPath);

    echo "Master schedule generated successfully at $outputPath";

} catch (Exception $e) {
    error_log("Error generating Excel report: " . $e->getMessage());
    echo "An error occurred while generating the report: " . $e->getMessage();
} finally {
    if (isset($mysqli)) {
        $mysqli->close();
    }
}
?>