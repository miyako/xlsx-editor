property _stdOut : Text
property _stdErr : Text
property stdOut : Collection
property stdErr : Collection

Class extends _CLI_Controller

Class constructor($CLI : cs:C1710._CLI)
	
	Super:C1705($CLI)
	
	This:C1470._stdOut:=""
	This:C1470._stdErr:=""
	
Function onData($worker : 4D:C1709.SystemWorker; $params : Object)
	
	This:C1470._stdOut+=$params.data
	
Function onDataError($worker : 4D:C1709.SystemWorker; $params : Object)
	
	This:C1470._stdErr+=$params.data
	
Function onResponse($worker : 4D:C1709.SystemWorker; $params : Object)
	
	This:C1470.stdOut:=Split string:C1554(This:C1470._stdOut; This:C1470.instance.EOL)
	This:C1470.stdErr:=Split string:C1554(This:C1470._stdErr; This:C1470.instance.EOL)
	
Function onError($worker : 4D:C1709.SystemWorker; $params : Object)
	
Function onTerminate($worker : 4D:C1709.SystemWorker; $params : Object)
	