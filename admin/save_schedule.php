<?php
include('./connection/session.php');
include('./connection/dbcon.php');

if ($_SERVER['REQUEST_METHOD'] == 'POST') {
    // Retrieve and sanitize inputs
    $floor_id = intval($_POST['floor']);
    $room_id = intval($_POST['room']);
    $days = $_POST['day_of_week'] ?? '';
    $start_time = $conn->real_escape_string($_POST['start_time']);
    $end_time = $conn->real_escape_string($_POST['end_time']);
    $subject_id = intval($_POST['subject']);
    $teacher_id = intval($_POST['teacher']);

    // Validate inputs
    if (empty($floor_id) || empty($room_id) || empty($days) || empty($start_time) || empty($end_time) || empty($subject_id) || empty($teacher_id)) {
        http_response_code(400);
        echo json_encode(["status" => "error", "message" => "All fields are required."]);
        exit;
    }

    if ($start_time >= $end_time) {
        http_response_code(400);
        echo json_encode(["status" => "error", "message" => "Start time must be earlier than end time."]);
        exit;
    }

    // Calculate duration of the new schedule
    $new_schedule_duration = (strtotime($end_time) - strtotime($start_time)) / 3600;

    // Check teacher's total hours for the specific subject on the selected day
    $sql_check_subject_hours = "
        SELECT SUM(TIMESTAMPDIFF(HOUR, start_time, end_time)) AS subject_hours
        FROM schedules
        WHERE id = ? 
        AND subject_code = ?
        AND day_of_week = ?
    ";
    $stmt = $conn->prepare($sql_check_subject_hours);
    $stmt->bind_param("iis", $teacher_id, $subject_id, $days);
    $stmt->execute();
    $result = $stmt->get_result();
    $row = $result->fetch_assoc();
    $current_subject_hours = $row['subject_hours'] ?: 0;
    $stmt->close();

    // Calculate total hours including new schedule
    $total_subject_hours = $current_subject_hours + $new_schedule_duration;

    if ($total_subject_hours > 3) {
        http_response_code(409);
        echo json_encode([
            "status" => "error",
            "message" => "Teacher's teaching time for this subject cannot exceed 3 hours per day.",
            "details" => [
                "current_hours" => round($current_subject_hours, 2),
                "new_hours" => round($new_schedule_duration, 2),
                "total_hours" => round($total_subject_hours, 2),
                "max_allowed" => 3
            ]
        ]);
        exit;
    }

    // Check room availability
    $sql_check = "
        SELECT * 
        FROM schedules 
        WHERE room_id = ? AND day_of_week = ? 
        AND ((start_time < ? AND end_time > ?) OR (start_time < ? AND end_time > ?))
    ";
    $stmt = $conn->prepare($sql_check);
    $stmt->bind_param("isssss", $room_id, $days, $end_time, $start_time, $start_time, $end_time);
    $stmt->execute();
    $result = $stmt->get_result();

    if ($result->num_rows > 0) {
        http_response_code(409);
        echo json_encode(["status" => "error", "message" => "Room is already booked for the selected time slot."]);
        exit;
    }

    // Insert new schedule entry
    $sql_insert = "
        INSERT INTO schedules (floor_id, room_id, day_of_week, start_time, end_time, subject_code, id) 
        VALUES (?, ?, ?, ?, ?, ?, ?)
    ";
    $stmt = $conn->prepare($sql_insert);
    $stmt->bind_param("iisssii", $floor_id, $room_id, $days, $start_time, $end_time, $subject_id, $teacher_id);

    if ($stmt->execute()) {
        http_response_code(201);
        echo json_encode(["status" => "success", "message" => "Schedule added successfully."]);
    } else {
        http_response_code(500);
        echo json_encode(["status" => "error", "message" => "Failed to add the schedule. Please try again."]);
    }

    $stmt->close();
}
?>