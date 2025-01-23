<?php
require 'vendor/autoload.php';

use PhpOffice\PhpSpreadsheet\IOFactory;
use PhpOffice\PhpSpreadsheet\Spreadsheet;
use PhpOffice\PhpSpreadsheet\Style\Alignment;

// Database connection details
$servername = "localhost";
$username = "root"; // Default username for XAMPP
$password = "";     // Default password for XAMPP
$dbname = "isadfc";

// Connect to the database
$conn = new mysqli($servername, $username, $password, $dbname);

// Check connection
if ($conn->connect_error) {
    die("Connection failed: " . $conn->connect_error);
}

// SQL query to fetch schedule data grouped by floor
$sql = "
  SELECT
    schedules.schedule_id, 
    floors.floor_number, 
    rooms.room_name, 
    schedules.day_of_week, 
    schedules.start_time, 
    schedules.end_time, 
    schedules.subject_code, 
    teaching_faculty_information.first_name, 
    teaching_faculty_information.last_name
FROM schedules
INNER JOIN floors ON schedules.floor_id = floors.floor_id
INNER JOIN rooms ON schedules.room_id = rooms.room_id
INNER JOIN teaching_faculty_information ON schedules.id = teaching_faculty_information.id
ORDER BY schedules.day_of_week, schedules.start_time
";

$result = $conn->query($sql);
if (!$result) {
    die("Query failed: " . $conn->error);
}

// Create a new spreadsheet
$spreadsheet = new Spreadsheet();

// Initialize variables to track the current floor and room
$previousFloor = '';
$previousRoom = '';
$currentRow = 1;

// Function to set and format floor header
function setFloorHeader($sheet, $row, $floorNumber)
{
    $sheet->setCellValue('A' . $row, 'Floor: ' . $floorNumber);
    $sheet->mergeCells('A' . $row . ':L' . $row);
    $sheet->getStyle('A' . $row)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle('A' . $row)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
}

// Function to set and format room header
function setRoomHeader($sheet, $row, $roomName)
{
    $sheet->setCellValue('A' . $row, 'Room: ' . $roomName);
    $sheet->mergeCells('A' . $row . ':L' . $row);
    $sheet->getStyle('A' . $row)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle('A' . $row)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
}

// Insert headers for each day of the week
function insertDayHeaders($sheet, $row)
{
    $sheet->setCellValue('A' . $row, 'Monday / Wednesday');
    $sheet->mergeCells('A' . $row . ':D' . $row);
    $sheet->setCellValue('E' . $row, 'Tuesday / Thursday');
    $sheet->mergeCells('E' . $row . ':H' . $row);
    $sheet->setCellValue('I' . $row, 'Friday / Saturday');
    $sheet->mergeCells('I' . $row . ':L' . $row);
    $sheet->getStyle('A' . $row . ':L' . $row)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle('A' . $row . ':L' . $row)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
}

// Column mapping for each day of the week
$dayColumns = [
    'Monday' => 'A',
    'Tuesday' => 'E',
    'Wednesday' => 'A',
    'Thursday' => 'E',
    'Friday' => 'I',
    'Saturday' => 'I',
];

// Process each row of the result
while ($data = $result->fetch_assoc()) {
    $floorNumber = $data['floor_number'];
    $roomName = $data['room_name'];
    $dayOfWeek = $data['day_of_week'];

    // Handle new floor
    if ($floorNumber != $previousFloor) {
        // Create a new sheet for the new floor
        if ($previousFloor != '') {
            $sheet = $spreadsheet->createSheet();
        } else {
            $sheet = $spreadsheet->getActiveSheet();
        }
        $sheet->setTitle('Floor ' . $floorNumber);
        $currentRow = 1;
        setFloorHeader($sheet, $currentRow, $floorNumber);
        $currentRow += 2;
        $previousRoom = '';
    }

    // Handle new room
    if ($roomName != $previousRoom) {
        setRoomHeader($sheet, $currentRow, $roomName);
        $currentRow++;
        insertDayHeaders($sheet, $currentRow);
        $currentRow++;
        $sheet->fromArray(['Code', 'Teacher', 'Start Time', 'End Time', 'Code', 'Teacher', 'Start Time', 'End Time', 'Code', 'Teacher', 'Start Time', 'End Time'], null, 'A' . $currentRow);
        $sheet->getStyle('A' . $currentRow . ':L' . $currentRow)->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
        $sheet->getStyle('A' . $currentRow . ':L' . $currentRow)->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);
        $currentRow++;
        $previousRoom = $roomName;
    }

    // Insert schedule details in the appropriate columns
    $startColumn = $dayColumns[$dayOfWeek];
    $sheet->fromArray([$data['subject_code'], $data['first_name'] . ' ' . $data['last_name'], $data['start_time'], $data['end_time']], null, $startColumn . $currentRow);
    $sheet->getStyle($startColumn . $currentRow . ':' . chr(ord($startColumn) + 3) . $currentRow)
        ->getAlignment()->setHorizontal(Alignment::HORIZONTAL_CENTER);
    $sheet->getStyle($startColumn . $currentRow . ':' . chr(ord($startColumn) + 3) . $currentRow)
        ->getAlignment()->setVertical(Alignment::VERTICAL_CENTER);

    $previousFloor = $floorNumber;
}

// Ensure the output directory exists and is writable
$outputDir = 'C:/xampp/htdocs/roomassignment/admin/output/';
if (!is_dir($outputDir)) {
    mkdir($outputDir, 0777, true);
}
if (!is_writable($outputDir)) {
    die("Error: The output directory is not writable.");
}

$outputPath = $outputDir . 'roomsched_report.xlsx';
try {
    $writer = IOFactory::createWriter($spreadsheet, 'Xlsx');
    $writer->save($outputPath);

    // Success message with JavaScript alert
    echo "<script>
            alert('Excel report created successfully!');
            window.location.href = 'schedule.php';  // Redirect to your desired page
          </script>";
} catch (Exception $e) {
    die("Error saving the Excel file: " . $e->getMessage());
}
// Close the database connection
$conn->close();
?>