<?php
session_start();
if ($_SESSION['OnLine']!=1) {
	header("Location: http://" . $_SERVER["HTTP_HOST"]."/");
}  
?>