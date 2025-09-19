Class extends _Form

Class constructor
	
	Super:C1705()
	
	$window:=Open form window:C675("loc")
	DIALOG:C40("loc"; This:C1470; *)
	
Function onLoad()
	
	var $folder : 4D:C1709.Folder
	$folder:=Folder:C1567("/SOURCES/")
	
	Form:C1466.loc:=cs:C1710.loc.new($folder; cs:C1710._locUI_Controller)
	
	Form:C1466.setEnabled("STOP"; False:C215)
	
Function onUnload()
	
	Form:C1466.loc.terminate()
	
Function count() : Object
	
	Form:C1466.stdOut:=""
	Form:C1466.stdErr:=""
	
	Form:C1466.loc.count()
	
	return Form:C1466
	
Function setEnabled($objectNames : Variant; $enabled : Boolean) : Object
	
	Case of 
		: (Value type:C1509($objectNames)=Is text:K8:3)
			$objectNames:=[$objectNames]
		: (Value type:C1509($objectNames)=Is collection:K8:32)
			//
		Else 
			$objectNames:=[]
	End case 
	
	For each ($objectName; $objectNames)
		OBJECT SET ENABLED:C1123(*; $objectName; $enabled)
	End for each 
	
	return Form:C1466