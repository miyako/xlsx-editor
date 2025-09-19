property data : Collection

Class extends _CLI

Class constructor($controller : 4D:C1709.Class)
	
	If (Not:C34(OB Instance of:C1731($controller; cs:C1710._XLSX_Controller)))
		$controller:=cs:C1710._XLSX_Controller
	End if 
	
	Super:C1705("xlsx-editor"; $controller)
	
Function get worker() : 4D:C1709.SystemWorker
	
	return This:C1470.controller.worker
	
Function terminate()
	
	This:C1470.controller.terminate()
	
Function _fromBlob($blob : Blob) : Text
	
	var $text : Text
	
	BASE64 ENCODE:C895($blob; $text)
	
	return $text
	
Function _fromPicture($picture : Picture; $codec : Text) : Text
	
	var $blob : Blob
	
	CONVERT PICTURE:C1002($picture; $codec)
	PICTURE TO BLOB:C692($picture; $blob; $codec)
	
	return "data:"+$codec+";base64,"+This:C1470._fromBlob($blob)
	
Function _fromDate($date : Date) : Integer
	
	Case of 
		: ($date<!1900-01-01!)
			
			//out of range
			
		: (!1900-01-01!<$date) & ($date<!1900-03-01!)
			
			$value:=$date-!1899-12-31!
			
		: ($date>!1900-02-28!)
			
			$value:=$date-!1899-12-30!
			
		Else 
/*
			
Microsoft date 60 (29th February 1900) does not exist!
			
https://en.wikipedia.org/wiki/Year_1900_problem
			
*/
	End case 
	
	return $value
	
Function _fromTime($time : Time) : Real
	
	return $time/86400
	
Function update($option : Variant) : cs:C1710.XLSX
	
	var $options : Collection
	
	Case of 
		: (Value type:C1509($option)=Is object:K8:27)
			$options:=[$option]
		: (Value type:C1509($option)=Is collection:K8:32)
			$options:=$option
		Else 
			$options:=[]
	End case 
	
	var $commands : Collection
	$commands:=[]
	
	For each ($option; $options)
		
		If ($option=Null:C1517) || (Value type:C1509($option)#Is object:K8:27)
			continue
		End if 
		
		var $file : Text
		
		If (Not:C34(OB Instance of:C1731($option.file; 4D:C1709.File))) || (Not:C34($option.file.exists))
			$file:="-"
		Else 
			$file:=This:C1470.expand($option.file).path
		End if 
		
		If (Not:C34(OB Instance of:C1731($option.output; 4D:C1709.File)))
			continue
		End if 
		
		var $output : 4D:C1709.File
		$output:=$option.output
		
		If ($option.values=Null:C1517) || (Value type:C1509($option)#Is object:K8:27)
			continue
		End if 
		
		var $values : Object
		$values:=$option.values
		
		var $datum : Object
		var $data : Collection
		$data:=[]
		
		var $sheet; $cell : Text
		For each ($sheet; $values)
			
			If ($sheet="")
				continue
			End if 
			
			If ($values[$sheet]=Null:C1517) || (Value type:C1509($values[$sheet])#Is object:K8:27)
				continue
			End if 
			
			var $value : Object
			$value:=$values[$sheet]
			
			For each ($cell; $value)
				If ($cell="")
					continue
				End if 
				
				$datum:={sheet: $sheet; cell: $cell}
				
				var $c : Variant
				$c:=$value[$cell]
				
				If (Value type:C1509($c)=Is object:K8:27)
					//get format,formula
					If ($c.format#Null:C1517) && (Value type:C1509($c.format)=Is text:K8:3)
						$datum.format:=$c.format
					End if 
					If ($c.formula#Null:C1517) && (Value type:C1509($c.formula)=Is text:K8:3)
						$datum.formula:=$c.formula
					End if 
					$c:=$c.value
				End if 
				
				var $vt : Integer
				$vt:=Value type:C1509($c)
				
				Case of 
					: ($vt=Is null:K8:31) || ($vt=Is undefined:K8:13)
						$datum.value:=Null:C1517
					: ($vt=Is date:K8:7)
						$datum.value:=This:C1470._fromDate($c)
					: ($vt=Is time:K8:8)
						$datum.value:=This:C1470._fromTime($c)
					: ($vt=Is BLOB:K8:12)
						$datum.value:=This:C1470._fromBlob($c)
					: ($vt=Is picture:K8:10)
						$datum.value:=This:C1470._fromPicture($c; "image/png")
					: ($vt=Is object:K8:27) || ($vt=Is collection:K8:32)
						$datum.value:=JSON Stringify:C1217($c)
					Else 
						$datum.value:=$c
				End case 
				$data.push($datum)
			End for each 
		End for each 
		
		$command:=This:C1470.escape(This:C1470.executablePath)
		$command+=" "
		$command+=This:C1470.escape($file)
		$command+=" "
		$command+=This:C1470.escape(This:C1470.expand($output).path)
		
		var $worker : 4D:C1709.SystemWorker
		$worker:=This:C1470.controller.execute($command; $data).worker
		$worker.wait()
		
	End for each 
	
	return This:C1470