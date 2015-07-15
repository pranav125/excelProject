var XLSX = require('xlsx');
var parseString = require('xml2js').parseString;
var FR = require('formulajs');
var _ = require('underscore');
var kexcel = require('kexcel');
var fs = require('fs');

var formulaUtil = require('./excelFormulaUtilities-0.9.43.min.js');
var formula = formulaUtil.excelFormulaUtilities.formula2JavaScript('=SUM(10, LARGE([10,20],2))');
console.log("formula:  ", formula);

function executeFunc(FR, fn) {
    return new Function('FR', 'return ' + fn)(FR);    
}

var f = executeFunc(FR, formula);
console.log(f);

/*Read the excel file*/
var workbook = XLSX.readFile('Book1.xlsx', {
    cellNF: true,    
    cellStyles: true,
    cellDates: true,
    sheetStubs: true,  
    bookDeps: true,
    bookFiles: true,
    bookVBA	: true,
    cellFormula: true
});



/* Model Excel File Read*/
var sheet_name_list = workbook.SheetNames;
var modelObject=workbook.Sheets[sheet_name_list[0]];
modelObject.definedNames = [];

var xml = workbook.Workbook.data;
parseString(xml, function (err, result) {
   result.workbook.definedNames[0].definedName.forEach(function(d) {       
       modelObject.definedNames.push({name: d.$.name, sheet: d._});
       
   });
});

//console.log(modelObject);

//var definedName1 =modelObject.definedNames[0].name;
//var definedName2 = modelObject.definedNames[1].name;
//var formula = modelObject.definedNames[2].sheet;
//
//var sheetSplitArray = formula.split('$');
//var cellNumber =  sheetSplitArray[1]+sheetSplitArray[2];
//var newFormula= modelObject[cellNumber]['f'];



//console.log(FR['SUM']);
