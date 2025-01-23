aa
<?php
include('./connection/session.php');
include('./components/header.php');
include('./connection/dbcon.php');
include('./components/nav-top1.php');
include('./components/main.php');
?>

<head>
    <meta charset="UTF-8">
    <meta name="viewport" content="width=device-width, initial-scale=1.0">
    <title>Select with Checkboxes</title>
    <link rel="stylesheet" href="addsched.css">
    <link rel="stylesheet" href="https://cdnjs.cloudflare.com/ajax/libs/font-awesome/6.0.0-beta3/css/all.min.css">
    <!-- Include Select2 CSS -->
    <link href="https://cdn.jsdelivr.net/npm/select2@4.0.13/dist/css/select2.min.css" rel="stylesheet" />
    <!-- jQuery (Make sure jQuery is included) -->
    <script src="https://ajax.googleapis.com/ajax/libs/jquery/3.5.1/jquery.min.js"></script>
    <!-- Include jQuery (required by Select2) -->
    <script src="https://code.jquery.com/jquery-3.6.0.min.js"></script>

    <script src="https://cdn.jsdelivr.net/npm/select2@4.0.13/dist/js/select2.min.js"></script>
</head>

<div class="wrapper">
    <div id="element" class="hero-body-schedule" style="background-color: white; ">
        <h2>
            <font color="Black">Add Schedule List</font>
        </h2>
        <a class="btn btn-primary" href="schedule.php"><i class="icon-arrow-left icon-large"></i>&nbsp;Back</a>
        <hr>


        <form id="save_voter" class="form-horizontal" method="POST" action="save_schedule.php">
            <fieldset>

                <div class="hai_naku">

                    <ul class="thumbnails_new_voter">
                        <li class="span3">
                            <div class="thumbnail_new_voter">
                                <!-- Success and Error Messages -->
                                <div class="confirmation" id="confirmation-message" style="display: none;">
                                    <i class="fas fa-check-circle"></i>
                                    <span>Schedule has been successfully added.</span>
                                </div>

                                <div class="error" id="error-message" style="display: none;">
                                    <i class="fas fa-times-circle"></i>
                                    <span>Error: The selected room is already used. Choose another
                                        time
                                        slot.</span>
                                </div>
                                <!-- Floor Selection -->
                                <div class="control-group">
                                    <label class="control-label" for="input01">Floor:</label>
                                    <div class="controls">
                                        <select id="floor" class="item" name="floor" onchange="fetchRooms()">
                                            <option value="">Select Floor</option>
                                            <?php include_once('fetch_floor.php'); ?>
                                        </select>
                                    </div>
                                </div>

                                <!-- Room Selection -->
                                <div class="control-group">
                                    <label class="control-label" for="input01">Rooms:</label>
                                    <div class="controls">
                                        <select id="room" class="item" name="room" required>
                                            <option value="">Select Room</option>
                                        </select>
                                    </div>
                                </div>

                                <!-- JavaScript for fetching rooms -->
                                <script>
                                    function fetchRooms() {
                                        var floorId = $('#floor').val();  // Get selected floor ID
                                        $('#room').html('<option>Loading...</option>'); // Show loading message in room dropdown

                                        $.ajax({
                                            url: 'fetch_rooms.php',
                                            type: 'GET',
                                            data: { floor_id: floorId },
                                            success: function (response) {
                                                $('#room').html(response);
                                            }
                                        });
                                    }
                                </script>

                                <!-- Day of the Lecture -->
                                <div class="control-group">
                                    <label class="control-label" for="input01">Day of Lecture:</label>
                                    <div class="controls">
                                        <select style="width: 75%;" id="schedule-day-input" name="day_of_week" required>
                                            <option value="Monday">Monday</option>
                                            <option value="Tuesday">Tuesday</option>
                                            <option value="Wednesday">Wednesday</option>
                                            <option value="Thursday">Thursday</option>
                                            <option value="Friday">Friday</option>
                                            <option value="Saturday">Saturday</option>
                                        </select>
                                    </div>
                                </div>

                                <!-- Time Start and End -->
                                <div class="control-group">
                                    <label class="control-label" for="input01">Time Start:</label>
                                    <div class="controls">
                                        <input class="item" type="time" id="start-time" name="start_time"
                                            style="width: 75%; height: 2.5em;" required>
                                    </div>
                                </div>

                                <div class="control-group">
                                    <label class="control-label" for="input01">Time End:</label>
                                    <div class="controls">
                                        <input class="item" type="time" id="end-time" name="end_time"
                                            style="width: 75%; height: 2.5em; " required>
                                    </div>
                                </div>

                                <!-- Subject Selection -->
                                <div class="control-group">
                                    <label class="control-label" for="input01">Class Code:</label>
                                    <div class="controls">
                                        <input type="text" class="item" id="subject" name="subject" required />
                                    </div>
                                </div>


                                <!-- Teacher Selection -->
                                <div class="control-group">
                                    <label class="control-label" for="input01">Teacher:</label>
                                    <div class="controls">
                                        <select id="teacher" class="item" name="teacher">
                                            <option value="">Select Teacher</option>
                                            <?php
                                            include 'fetch_teacher.php';
                                            while ($teacher = $teachers_result->fetch_assoc()) { ?>
                                                <option value="<?php echo $teacher['id']; ?>">
                                                    <?php echo $teacher['teacher_name']; ?>
                                                </option>
                                            <?php } ?>
                                        </select>
                                    </div>
                                </div>
                                <!-- Save Button -->
                                <div class="control-group">
                                    <div class="controls">
                                        <button type="submit" id="save_voter" class="btn btn-primary"><i
                                                class="icon-save icon-large"></i>Save</button>
                                    </div>
                                </div>

                            </div>
                        </li>
                    </ul>
                </div>
            </fieldset>
        </form>
    </div>
</div>

<script>
    $(document).ready(function () {
        $('#save_voter').on('submit', function (e) {
            e.preventDefault(); // Prevent page reload

            // Perform AJAX request
            $.ajax({
                url: 'save_schedule.php', // The PHP script to save the schedule
                type: 'POST',
                data: $(this).serialize(), // Serialize form data
                success: function (response) {
                    var result = JSON.parse(response); // Parse JSON response from server

                    if (result.status === 'success') {
                        showConfirmationMessage(); // Show success message
                    } else {
                        showErrorMessage(); // Show error message
                    }
                },
                error: function () {
                    showErrorMessage(); // Show error message on failure
                }
            });
        });
    });

    // Example function to show the confirmation message
    function showConfirmationMessage() {
        document.getElementById("confirmation-message").style.display = "block";
        document.getElementById("error-message").style.display = "none";
    }

    // Example function to show the error message
    function showErrorMessage() {
        document.getElementById("error-message").style.display = "block";
        document.getElementById("confirmation-message").style.display = "none";
    }
</script>

<?php include('footer.php'); ?>

<div class="modal hide fade" id="myModal">
    <div class="modal-header">
        <button type="button" class="close" data-dismiss="modal">×</button>
        <h3> </h3>
    </div>
    <div class="modal-body">
        <p>
            <font color="gray">Are You Sure you Want to LogOut?</font>
        </p>
    </div>
    <div class="modal-footer">
        <a href="#" class="btn" data-dismiss="modal">No</a>
        <a href="logout.php" class="btn btn-primary">Yes</a>
    </div>
</div>

</html>