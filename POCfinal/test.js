var formula = require('formulajs');
//console.log(formula);

var c = formula.SUM(formula.LARGE([10,23,13], 1), 3);

console.log(c);
var kexcel = require('kexcel');
var fs = require('fs');
kexcel.open( 'Book1.xlsx', function(err, workbook) {

   // Get first sheet 
   var sheet1 = workbook.getSheet(0);
   console.log('workbook',JSON.stringify(workbook));
   // Duplicate a sheet 
   var duplicatedSheet = workbook.duplicateSheet(0,'My duplicated sheet');

   // Save the file 
   var output = fs.createWriteStream('tester.xlsx');
   workbook.pipe(output);
})
