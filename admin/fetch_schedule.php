<?php
include('./connection/dbcon.php');


$sql_fetch = "
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

$result = $conn->query($sql_fetch);

return $result;
?>