<?php
// track.php

// Image URL
$imagePath = 'https://www.endocreative.com/wp-content/uploads/2014/11/banner-772x250.png';

// Database connection details
$servername = "NEL";
$username = "NAAH";
$password = "NO";
$dbname = "NO";

// Connect to the database
$conn = new mysqli($servername, $username, $password, $dbname);

// Check connection
if ($conn->connect_error) {
    error_log("Connection failed: " . $conn->connect_error);
    exit;
}

// Get the email ID from the query parameter
$email_id = isset($_GET['email_id']) ? $conn->real_escape_string($_GET['email_id']) : '';

// Insert into the tracking table if email ID is provided
if ($email_id) {
    $sql = "INSERT INTO email_tracking (email_id, status, timestamp) VALUES ('$email_id', 'Abierto', NOW(),)";
    if (!$conn->query($sql)) {
        error_log("Error inserting tracking data: " . $conn->error);
    }
}

// Serve the image
header('Content-Type: image/png');
header('Content-Disposition: inline; filename="track_image.png"');

// Read and output the image file
$imageData = file_get_contents($imagePath);
echo $imageData;

// Close the database connection
$conn->close();
?>
