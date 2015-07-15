var XLSX = require('xlsx');
var parseString = require('xml2js').parseString;
var FR = require('formulajs');
var express = require('express');
var path = require('path');
var bodyParser = require('body-parser');
var formulaUtil = require('./excelFormulaUtilities-0.9.43.min.js');
var app = express();

var done = false;
var parsedSheetArray = [];
var definedNamesArray = [];
var formulaObject = {};
var definedNamedObject = {};

app.use(bodyParser.urlencoded({ extended: true }));
app.use(bodyParser.json());
app.use(express.static(path.join(__dirname, 'public')));

app.get('/',function(req,res){
  res.sendfile("./public/index.html");
});

app.post('/executeformula', function(req, res){
	console.log('\n\n------------- Client Input JSON ---------------\n');
	console.log(req.body);	
	res.end('Result: ' + parseUserInput(req.body));
});

/*Read the excel file*/
var workbook = XLSX.readFile('excelModel.xlsx', {
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
var replaceHtmlEntites = (function() {
    var translate_re = /&(nbsp|amp|quot|lt|gt);/g,
        translate = {
            'nbsp': String.fromCharCode(160), 
            'amp' : '&', 
            'quot': '"',
            'lt'  : '<', 
            'gt'  : '>'
        },
        translator = function($0, $1) { 
            return translate[$1]; 
        };

    return function(s) {
        return s.replace(translate_re, translator);
    };
})();


var sheetListLength = sheet_name_list.length;
var xml = workbook.Workbook.data;

for(var sheet = 0; sheet < sheetListLength; sheet++){
	var modelObject=workbook.Sheets[sheet_name_list[sheet]];
	parsedSheetArray.push(modelObject);
	sheet_name_list[sheet] = replaceHtmlEntites(sheet_name_list[sheet])
}

//console.log(sheet_name_list);

//console.log(JSON.stringify(parsedSheetArray));
formDefinedNamesArray();
//console.log('----------->'+parsedSheetArray.length);
function formDefinedNamesArray(){
	parseString(xml, function (err, result) {
		if(result.workbook.definedNames){
		   result.workbook.definedNames[0].definedName.forEach(function(d) {       
			   definedNamesArray.push({name: d.$.name, value: d._});       
		   });
		}
		//console.log(definedNamesArray);
		formulateDefinedJSONs();
	});
}


function formulateDefinedJSONs(){
	for(var index = 0; index < definedNamesArray.length; index++){
		var namedObject = definedNamesArray[index];
		var sheetValue = namedObject.value;
		var valueStringArray = sheetValue.split('!');
		var sheetName = valueStringArray[0];
		var regex = new RegExp("'", 'g');
		sheetName = sheetName.replace(regex, "");
		if(valueStringArray[1]){
			var sheetLocation = sheet_name_list.indexOf(sheetName);
			if(parseInt(sheetLocation) >= 0){
			//console.log(sheetName + '  ' + namedObject.value + '  ' + sheet_name_list.indexOf(sheetName) + '-----\n');	
				var cellNumberArray = valueStringArray[1].split('$');
				var cellNumber = cellNumberArray[1] + cellNumberArray[2];
				var cellObject = parsedSheetArray[sheetLocation][cellNumber];
				//console.log(cellObject);
				if(cellObject){
					if(cellObject.f){
						try{
							var formula = formulaUtil.excelFormulaUtilities.formula2JavaScript('='+cellObject.f);
							formulaObject[namedObject.name] = formula;
						}catch(e){
							console.log(e);
							//process.exit(1);
						}
					}else{
						if(definedNamedObject[sheetName] != undefined){
							definedNamedObject[sheetName][namedObject.name] = cellObject.v;
						} else{
							definedNamedObject[sheetName] = {};
							definedNamedObject[sheetName][namedObject.name] = cellObject.v;
						}
					}
				}
			}
		}
	}
}

console.log('\n\n----------------- Formula Object ------------------\n');
console.log(formulaObject);
console.log('\n\n-------------- Defined Names Object ---------------\n');
console.log(definedNamedObject);

function parseUserInput(clientObject){
	var formula = clientObject.formula;
	if(formula){
		var modelFormula = formulaObject[formula];
		if(modelFormula){
			delete clientObject.formula;
			return parseFormula(modelFormula, clientObject);
		}
	}
}

function parseFormula(formula, clientObject){
	for(var key in clientObject){
		var regex = new RegExp(key, 'g');
		formula = formula.replace(regex, clientObject[key]);
	}
	console.log('\n\n-------------- Parsed Formula ---------------\n');
	console.log(formula);
	return executeFormula(formula);
}

function executeFormula(formula){
	function executeFunc(FR, fn) {
	    return new Function('FR', 'return ' + fn)(FR);    
	}
	
	var f = executeFunc(FR, formula);
	console.log('\n\n------------------ Result -------------------\n');
	console.log(f);
	return f;
}
/*Modification*/




//
app.listen(4000);