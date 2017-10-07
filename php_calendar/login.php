
<?php 
echo htmlspecialchars($_POST['csv_file']);
RunMyApp(htmlspecialchars($_POST['csv_file']));
 ?>

<?php 
//------------------------------------------------------
//------------------------------------------------------
function RunMyApp($csvfile)
{

require_once __DIR__.'/../vendor/autoload.php';
    session_start();

	define('APPLICATION_NAME', 'Gesthand Web App');
    define('CLIENT_SECRET_PATH', 'client_secret_gesthand_web.json');
	//define('SCOPES', implode(' ', array(Google_Service_Calendar::CALENDAR)));
echo "AAAAAAAAAAAAAAAAAAAAA";

	$client = new Google_Client();
	$client->setAuthConfig(CLIENT_SECRET_PATH);
	$client->setAccessType("offline");        // offline access
	$client->setIncludeGrantedScopes(true);   // incremental auth
	$client->addScope(Google_Service_Calendar::CALENDAR);
//$client->revokeToken(); die();
//session_unset();
	$client->setRedirectUri('http://' . $_SERVER['HTTP_HOST'] . '/oauth2callback.php');

	if (!isset($_SESSION['access_token']) || !$_SESSION['access_token']) {
		$redirect_uri = 'http://' . $_SERVER['HTTP_HOST'] . '/oauth2callback.php';
  		header('Location: ' . filter_var($redirect_uri, FILTER_SANITIZE_URL));
	} else {
  	$client->setAccessToken($_SESSION['access_token']);

  	$service = new Google_Service_Calendar($client);
	$cnt =readCSV($service, $csvfile);
	echo "<br>Fermeture... ".$cnt." evenements ont été traités.";
   
}
}

//------------------------------------------------------
//07/10/2017Array ( [0] => 2017-40 [1] => M610035151 [2] => championnat régional honneur - 18 ans masculin 
//[3] => Poule 16 [4] => 3 [5] => 07/10/2017 [6] => 14:00:00 [7] => PIGNAN HB 2 [8] => VILLENEUVE HB 
//[9] => PIGNAN HANDBALL [10] => [11] => [12] => [13] => [14] => MABKSAE [15] => Salle du Bicentenaire 
//[16] => rue de l'europe [17] => 34570 [18] => PIGNAN [19] => [20] => [21] => [22] => Bleu [23] => Noir 
//[24] => [25] => [26] => [27] => [28] => ALVERNHES ROMAIN [29] => 0669258916 [30] => ALVERNHES ROMAIN 
//[31] => 0669258916 [32] => 6134060 [33] => 6134078 )
//------------------------------------------------------
function WriteAnEvent($serv, $line, & $array_color) {

    $event = new Google_Service_Calendar_Event();  
    ##print_r ($line);
    $event->setSummary(sprintf("%s/%s",$line[7], $line[8]));
    $event->setLocation( sprintf("%s,%s,%s,%s",$line[15],$line[16],$line[17],$line[18]));
    $event->setDescription( sprintf("J%s %s", $line[4],$line[2]));

    $start = new Google_Service_Calendar_EventDateTime();
  
    $aa = sprintf("%s %s",$line[5],$line[6]);
    $date = DateTime::createFromFormat('j/m/Y H:i:s', $aa);
	$start->setDateTime($date->format('Y-m-d\TH:i:s'));	
	$start->setTimeZone('Europe/Paris');
    $event->setStart($start);

    $end = new Google_Service_Calendar_EventDateTime();
    //$end->setDateTime('2017-10-02T10:25:00.000-05:00');
    // TODO Add 2 hours
    $date->add(new DateInterval('PT2H'));
    $end->setDateTime($date->format('Y-m-d\TH:i:s'));
  	$end->setTimeZone('Europe/Paris');
    $event->setEnd($end);

    // Set the colorId according to team
    // the first color is 4 (kind of green/blue)
    $col_id = count($array_color) + 4;
    if (isset($array_color[$line[1]])) {
    	$col_id = $array_color[$line[1]];	// Get existing color
    } else {
    	// Add the color
    	$array_color[$line[1]]= $col_id; // 4 is the first color !
    	}

    // limit...	
    if ($col_id > 23)
    	$col_id = 23;
     
	$event->setColorId($col_id);
    $event->id=strtolower(sprintf("%s%s", $line[1], $line[4]));

    // Create or update
    // the id is based on Poule and day_of_match (Jx)
    $event->id=strtolower(sprintf("%s%s", $line[1], $line[4]));

	try {
    	$createdEvent = $serv->events->insert('primary', $event);
    } catch (Google_Service_Exception $e ) {
    //} catch (Exception $e ) {
    	    //echo "Caught Google_ServiceException:\n";
    	    if ($e->getCode() == "409") {
    	    ##echo "Event $event->id will be updated...\n";
   	    	$createdEvent = $serv->events->update('primary',$event->id, $event);
 			echo $createdEvent->getId();
 		} else
 		{
 			echo "Event not created, fatal !\n";
    	    echo $e->getCode().":";
    	    echo $e->getMessage();
    	    echo "<br>";
 		}
 		
	}
    //echo "ID => <br>"; 
    #echo $createdEvent->getId() . "<br>";
    //echo "HTML => <br>";
    //echo $createdEvent->getHtmlLink() . "<br>";
}

//------------------------------------------------------
//------------------------------------------------------
function readCSV($serv, $csvFile)
{
    $header = NULL;
	//$data = array();
	$array_color=array();
	$cnt = 0;

    //echo "<br>Ouverture fichier CSV (".$csvFile.")";

  	$fp = fopen($csvFile,'r') or die("can not open file");
	while($row = fgetcsv($fp,0,";")) { 
	if(!$header)
		$header = $row;			//Save 1st line..the headers
		else {
			if (!empty($row[5])) {
				//print_r($row);
				echo "<br> Evenement: ".$row[5]."<br>";
                WriteAnEvent($serv, $row, $array_color);		
                $cnt = $cnt +1;
	//			$data[] = array_combine($header, $row);
		}
	}
	}
	fclose($fp) or die("can not close file");
	return $cnt;
	}


