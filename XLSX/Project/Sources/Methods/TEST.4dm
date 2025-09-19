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
$values.TEST.D14:=!1974-09-22!

var $outputFile : 4D:C1709.File
$outputFile:=Folder:C1567(fk desktop folder:K87:19).file("test.xlsx")

var $XLSX : cs:C1710.XLSX
$XLSX:=cs:C1710.XLSX.new()

/*

sync syntax (1 pass: use object)

*/

//$XLSX.update({file: $templateFile; values: $values; output: $outputFile})


/*

sync syntax (3 passes: use collection of objects)

*/

$XLSX.update([\
{file: $templateFile; values: $values; output: $outputFile}; \
{file: $templateFile; values: $values; output: $outputFile}; \
{file: $templateFile; values: $values; output: $outputFile}])

/*

async syntax (must use in a worker or dialog)

*/