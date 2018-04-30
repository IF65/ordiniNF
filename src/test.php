<?php

	// per prima cosa verifico che il file sia stato effettivamente caricato
	if (!isset($_FILES['userfile']) || !is_uploaded_file($_FILES['userfile']['tmp_name'])) {
	  	echo 'Non hai inviato nessun file...';
		exit;
	}
	//echo json_encode($_FILES['userfile'], true);
	//echo json_encode(getallheaders(), true);

	if (move_uploaded_file( $_FILES['userfile']['tmp_name'], "/phpUpload/".$_FILES['userfile']['name'])) {
		echo "file spostato\n";
	} else {
		echo json_encode($_FILES, true);
	}

	//$info = pathinfo($_FILES['userfile']['name']);
	//$ext = $info['extension']; // get the extension of the file
	//$newname = "newname.".$ext;

	//$target = '/'.$newname;


