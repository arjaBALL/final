<?php
// Include necessary files
// include('./connection/session.php');
include('./components/header.php');
include('./connection/dbcon.php');
include('./components/nav-top1.php');
include('./components/main.php');
?>

<!DOCTYPE html>
<html lang="en">

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <meta http-equiv="X-UA-Compatible" content="ie=edge">
    <title>Class Schedule</title>
    <!-- Include your CSS and JavaScript files here -->
</head>

<body>

    <div class="wrapper">
        <div id="element" class="hero-body-subject-add">
            <h2>
                <font color="Black">Class Schedule List</font>
            </h2>
            <a class="btn btn-primary" href="add_schedule.php">
                <i class="icon-plus-sign icon-large"></i>&nbsp;Add Class Schedule
            </a>
            <hr>
            <div class="demo_jui">
                <table cellpadding="0" cellspacing="0" border="0" class="display jtable" id="log">
                    <thead>
                        <tr>
                            <th>Floor</th>
                            <th>Room</th>
                            <th>Day</th>
                            <th>Start Time</th>
                            <th>End Time</th>
                            <th>Class Code</th>
                            <th>Teacher</th>
                            <th>Actions</th>
                        </tr>
                    </thead>
                    <tbody>
                        <?php
                        include('fetch_schedule.php'); // Fetch schedule dat
                        if ($result->num_rows > 0) {
                            while ($row = $result->fetch_assoc()) {
                                $id = $row['schedule_id'];
                                ?>
                                <tr class="del<?php echo $id; ?>">
                                    <td><?php echo htmlspecialchars($row['floor_number']); ?></td>
                                    <td><?php echo htmlspecialchars($row['room_name']); ?></td>
                                    <td><?php echo htmlspecialchars($row['day_of_week']); ?></td>
                                    <td><?php echo htmlspecialchars($row['start_time']); ?></td>
                                    <td><?php echo htmlspecialchars($row['end_time']); ?></td>
                                    <td><?php echo htmlspecialchars($row['subject_code']); ?></td>
                                    <td><?php echo htmlspecialchars($row['first_name'] . " " . $row['last_name']); ?></td>
                                    <td align="center">
                                        <a class="btn btn-info"
                                            href="edit_schedule.php?id=<?php echo $row['schedule_id']; ?>">Edit</a>
                                        <a class="btn btn-danger" data-toggle="modal"
                                            href="#deleteModal<?php echo $row['schedule_id']; ?>">Delete</a>

                                        <!-- Delete Confirmation Modal -->
                                        <div class="modal fade" id="deleteModal<?php echo $row['schedule_id']; ?>" tabindex="-1"
                                            role="dialog">
                                            <div class="modal-dialog" role="document">
                                                <div class="modal-content">
                                                    <div class="modal-header">
                                                        <h5 class="modal-title">Confirm Deletion</h5>
                                                        <button type="button" class="close" data-dismiss="modal"
                                                            aria-label="Close">
                                                            <span aria-hidden="true">&times;</span>
                                                        </button>
                                                    </div>
                                                    <div class="modal-body">
                                                        <p>Are you sure you want to delete this schedule?</p>
                                                    </div>
                                                    <div class="modal-footer">
                                                        <a class="btn btn-danger"
                                                            href="delete_schedule.php?id=<?php echo $row['schedule_id']; ?>">
                                                            Yes, Delete
                                                        </a>
                                                        <button type="button" class="btn btn-secondary"
                                                            data-dismiss="modal">Cancel</button>
                                                    </div>
                                                </div>
                                            </div>
                                        </div>
                                    </td>
                                </tr>
                                <?php
                            }
                        } else {
                            echo "<tr><td colspan='8'>No schedules found</td></tr>";
                        }
                        ?>
                    </tbody>
                </table>
            </div>
        </div>
    </div>

    <!-- Log out Modal -->
    <div class="modal fade" id="myModal" tabindex="-1" role="dialog">
        <div class="modal-dialog" role="document">
            <div class="modal-content">
                <div class="modal-header">
                    <button type="button" class="close" data-dismiss="modal">&times;</button>
                    <h3>Confirm Logout</h3>
                </div>
                <div class="modal-body">
                    <p>
                        <font color="gray">Are you sure you want to log out?</font>
                    </p>
                </div>
                <div class="modal-footer">
                    <a href="#" class="btn" data-dismiss="modal">No</a>
                    <a href="logout.php" class="btn btn-primary">Yes</a>
                </div>
            </div>
        </div>
    </div>

    <!-- Include necessary JavaScript files here (e.g., jQuery, Bootstrap) -->
</body>

</html>