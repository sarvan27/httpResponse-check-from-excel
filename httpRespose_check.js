//import exceljs and XMLHttpRequest
var Excel = require('exceljs');
var XMLHttpRequest = require("xmlhttprequest").XMLHttpRequest;

//Read Excel file
var workbook = new Excel.Workbook();
var file = <excel_file name>;
workbook.xlsx.readFile(file).then(function () {
            
//Get sheet by Name
var worksheet=workbook.getWorksheet('<sheet name>');
            
//get url value and check response and write response
var i=2;
do{
    //'A' column in sheet to get the list of URL
    var cell_url = 'A'+i;
    //'B' column in sheet to write the result
    var cell_res = 'B'+i;
    var url = worksheet.getCell(cell_url).value;
    console.log(url);
    worksheet.getCell(cell_res).value=validate(url);
i++;
}
while(i<=worksheet.rowCount)

// function to validate URL
function validate(uri) {
    var http = new XMLHttpRequest();
    http.open('GET', uri, false);
    http.send();
    
    var b = http.status;
    if(b==200){
        var response = http.responseText;
        if(response.includes("version=1.10.3")){
            return true;
        }
        else
        return false;
        
    }
    else
    return b;
    }

//Save the workbook
return workbook.xlsx.writeFile(file);

});




