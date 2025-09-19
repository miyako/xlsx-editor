//%attributes = {"invisible":true}
/*

an existing spreadsheet template;
you may create one with Microsoft Excel
a blank spreadsheet is used if omitted
 
*/
var $templateFile : 4D:C1709.File
$templateFile:=File:C1566("/DATA/sample.xlsx")

/*

suppose we want to set values in the sheet #1
which is titled "TEST"
D10, D12 D14
setup the data object like so:

*/

var $values : Object

$values:={}

$values.TEST:={}
$values.TEST.D10:="miyako"
$values.TEST.D12:="kesuke.miyako@4d.com"
$values.TEST.D14:=!1974-09-22!  //simple value
$values.TEST.D14:={value: !1974-09-22!; format: "dd-mm-yyyy"}  //value with format specifier

$values.TEST.A1:=1.23456
$values.TEST.A2:=2.34567
$values.TEST.A3:=3.45678

/*

applications like Microsoft Excel automatically recalculate formula when a file is opening
but you should set the values beforehand if the spreadsheet is to be parsed directly

*/

$values.TEST.A4:={formula: "SUM(A1:A3)"; format: "0.00"; value: $values.TEST.A1+$values.TEST.A2+$values.TEST.A3}

$values.TEST.A5:={\
value: "I have style y'all"; \
bold: True:C214; \
italic: True:C214; \
size: 14; \
font: "Arial"; \
stroke: "FFFF0000"; \
fill: "FFFFFF00"; \
halign: "center"; \
valign: "center"; \
left: {style: "thin"; color: "FF0000FF"}; \
right: {style: "thin"; color: "FF0000FF"}; \
top: {style: "double"; color: "FF00FF00"}; \
bottom: {style: "double"; color: "FF00FF00"}}

/*

border styles

"none" (default)
"thin"
"medium"
"thick"
"dashed"
"dotted"
"double"
"hair"
"mediumDashed"
"dashDot"
"mediumDashDot"

*/

var $outputFile : 4D:C1709.File
$outputFile:=Folder:C1567(fk desktop folder:K87:19).file("test.xlsx")

var $XLSX : cs:C1710.XLSX
$XLSX:=cs:C1710.XLSX.new()

/*

1 pass: use object

onData() is called for stdOut stream
onDataError() is called for stdErr stream
onResponse() is called when the export is complete

*/

$XLSX.update({file: $templateFile; values: $values; output: $outputFile})

/*

3 passes: use collection of objects

you can safely run this in a non-worker process
you may stop receiving callbacks but workers will complete silently

*/

If (True:C214)
	$XLSX.update([\
		{file: $templateFile; values: $values; output: Folder:C1567(fk desktop folder:K87:19).file("1.xlsx")}; \
		{file: $templateFile; values: $values; output: Folder:C1567(fk desktop folder:K87:19).file("2.xlsx")}; \
		{file: $templateFile; values: $values; output: Folder:C1567(fk desktop folder:K87:19).file("3.xlsx")}])
End if 
