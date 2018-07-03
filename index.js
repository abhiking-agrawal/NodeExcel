// Require library
const xl = require('excel4node')
const excelToJson = require("convert-excel-to-json")
const sortBy = require('sort-by')

//Config files
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
ws.cell(4, 3).string("Critic Score").style(styleConst.tableHeader);
ws.cell(4, 4).string("Album Name").style(styleConst.tableHeader);
ws.cell(4, 5).string("Artist").style(styleConst.tableHeader);
ws.cell(4, 6).string("Release Date").style(styleConst.tableHeader);

//TO keep track of Genre group
var oldGenre = ""
var rowType = "1"

//Loop to iterate all the input data
for (i = 1; i <= sortedData.length; i++) {
    var temp = sortedData[i - 1]
    //Check for Genre change
    if(oldGenre != temp["Genre"]){
        var rowType = rowType == "1" ? "0" : "1";
        oldGenre = temp["Genre"]
    }
    
    
    ws.cell(4 + i, 1).number(temp["SNO"]).style(styleConst["row" + rowType]);
    ws.cell(4 + i, 2).string(temp["Genre"]).style(styleConst["row" + rowType]);
    ws.cell(4 + i, 3).number(temp["Critic Score"]).style(styleConst["row" + rowType]);
    ws.cell(4 + i, 4).string(temp["Album Name"]).style(styleConst["row" + rowType]);
    ws.cell(4 + i, 5).string(temp["Artist"]).style(styleConst["row" + rowType]);
    ws.cell(4 + i , 6).date(temp["Release Date"])
            .style(styleConst["row" + rowType]).style({ numberFormat: 'd-mmm-yy' });;

    //call the function once ready with the sheet
    if (i == sortedData.length) {
        writeToTheFile()
    }
}

//Generate the output excel file
function writeToTheFile() {
    wb.write('data/' + constants.outputFileName + '.xlsx', function (err, stats) {
        if (err) {
            console.error(err);
        } else {
            console.log("Output file generated successfuly...");
            console.log(stats); // Prints out an instance of a node.js fs.Stats object
        }
    });
}

