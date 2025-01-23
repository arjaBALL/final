<div class="navbar navbar-fixed-top">
    <div class="navbar-inner">
        <div class="container">
            <a class="branded"></a>
            <a class="brand">
                <h1>Class Scheduling System</h1>
            </a>
            <div class="time_top">
                <font color="white">
                    <?php
                    // Display the current date
                    $Today = date('Y-m-d'); // Use full year format
                    $formattedDate = date('l, F d, Y', strtotime($Today));
                    echo $formattedDate;
                    ?>
                </font>
            </div>
        </div>
    </div>
</div>