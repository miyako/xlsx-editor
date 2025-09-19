Class extends _XLSX_Controller

Class constructor($CLI : cs:C1710._CLI)
	
	Super:C1705($CLI)
	
Function onData($worker : 4D:C1709.SystemWorker; $params : Object)
	
	If (Form:C1466#Null:C1517)
		Form:C1466.stdOut+=$params.data
	End if 
	
Function onDataError($worker : 4D:C1709.SystemWorker; $params : Object)
	
	If (Form:C1466#Null:C1517)
		Form:C1466.stdErr+=$params.data
	End if 
	
Function onResponse($worker : 4D:C1709.SystemWorker; $params : Object)
	
	Super:C1706.onResponse($worker; $params)
	
	If (Form:C1466#Null:C1517)
		$data:=This:C1470.instance.data
		ALERT:C41([JSON Stringify:C1217($data; *); "-"*36; $data.sum("count")].join("\r"))
	End if 
	
Function onError($worker : 4D:C1709.SystemWorker; $params : Object)
	
	If (Form:C1466#Null:C1517)
		
	End if 
	
Function onTerminate($worker : 4D:C1709.SystemWorker; $params : Object)
	
	If (Form:C1466#Null:C1517)
		Form:C1466.setEnabled("STOP"; False:C215).setEnabled("TEST"; True:C214)
	End if 