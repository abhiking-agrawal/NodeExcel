// Require library
const xl = require('excel4node')
const constants = require("./config/constant")
const styleConst = require("./config/style")
 
// Create a new instance of a Workbook class
var wb = new xl.Workbook();
 
// Add Worksheets to the workbook
var ws = wb.addWorksheet(constants.sheetName);
 
//Set the Name cell
ws.cell(2,2).string("Name").style(styleConst.nameHeader);
 
//Set the Name value cell
ws.cell(2,3,2,4,true).string("Agrawal, Abhijeet").style(styleConst.nameHeaderValue);

//Set the table headers
ws.cell(4,1).string("SNO").style(styleConst.tableHeader);
ws.cell(4,2).string("Genre").style(styleConst.tableHeader);
ws.cell(4,3).string("Credit Score").style(styleConst.tableHeader);
ws.cell(4,4).string("Album Name").style(styleConst.tableHeader);
ws.cell(4,5).string("Artist").style(styleConst.tableHeader);
ws.cell(4,6).string("Release Date").style(styleConst.tableHeader);

wb.write('data/'+constants.outputFileName +'.xlsx', function (err, stats) {
    if (err) {
        console.error(err);
    }  else {
        console.log(stats); // Prints out an instance of a node.js fs.Stats object
    }
});