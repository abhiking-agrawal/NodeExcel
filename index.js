// Require library
const xl = require('excel4node')
const excelToJson = require("convert-excel-to-json")
const sortBy = require('sort-by')

const constants = require("./config/constant")
const styleConst = require("./config/style")

//read the input excel file
const result = excelToJson({
    sourceFile: constants.inputFileSrc,
    header: {
        rows: 1
    },
    columnToKey: {
        '*': '{{columnHeader}}'
    },
    sheets: [constants.inputSheetName]
});

//store the json converted data to a variable
var inputData = result[constants.inputSheetName]

//sort the data based on the specified keys, put minus(-) before key for descending
var sortedData = inputData.sort(sortBy("Genre", "-Critic Score"))
// Create a new instance of a Workbook class
var wb = new xl.Workbook();

// Add Worksheets to the workbook
var ws = wb.addWorksheet(constants.sheetName);

//Set the Name cell
ws.cell(2, 2).string("Name").style(styleConst.nameHeader);

//Set the Name value cell
ws.cell(2, 3, 2, 4, true).string("Agrawal, Abhijeet").style(styleConst.nameHeaderValue);

//Set the table headers
ws.cell(4, 1).string("SNO").style(styleConst.tableHeader);
ws.cell(4, 2).string("Genre").style(styleConst.tableHeader);
ws.cell(4, 3).string("Credit Score").style(styleConst.tableHeader);
ws.cell(4, 4).string("Album Name").style(styleConst.tableHeader);
ws.cell(4, 5).string("Artist").style(styleConst.tableHeader);
ws.cell(4, 6).string("Release Date").style(styleConst.tableHeader);

//f
for(i=1; i <= sortedData.length; i++){
    var rowType = "" +i % 2;
    var temp = sortedData[i]
    ws.cell(5 , 1).number(temp["SNO"]).style(styleConst.row1);
    ws.cell(5 , 2).string(temp["Genre"]).style(styleConst.row1);
    ws.cell(5 , 3).number(temp["Credit Score"]).style(styleConst.row1);
    ws.cell(5 , 4).string(temp["Album Name"]).style(styleConst.row1);
    ws.cell(5 , 5).string(temp["Artist"]).style(styleConst.row1);
    ws.cell(5 , 6).string(temp["Release Date"]).style(styleConst.row1);
    // if(i == sortedData.length){
    //     writeToTheFile()
    // }
}

function writeToTheFile() {
    wb.write('data/' + constants.outputFileName + '.xlsx', function (err, stats) {
        if (err) {
            console.error(err);
        } else {
            console.log(stats); // Prints out an instance of a node.js fs.Stats object
        }
    });
}
writeToTheFile()
