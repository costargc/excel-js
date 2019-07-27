var Excel = require("exceljs");

var RowHeader = [
    "Client ID",
    "Employee ID",
    "Married Status",
    "Home State",
    "State Work In",
    "Total Allowances",
    "Additional Allowances",
    "Dependent Allowances",
    "Personal Allowances"
];

var workbook = new Excel.Workbook();

workbook.xlsx.readFile("./File Provided.xlsx").then(() => {
    var worksheet = workbook.getWorksheet("Total Impact");
    //Loop over Rows
    let clientID = worksheet.getRow(2).values[1];
    let clientName = worksheet.getRow(2).values[2];
    let newWorkbook = new Excel.Workbook();
    newWorkbook.addWorksheet("Impact");
    let newWorksheet = newWorkbook.getWorksheet("Impact");
    newWorksheet.addRow(RowHeader);

    worksheet.eachRow(function (row, rowNumber) {
        if (rowNumber > 1) {
            if (
                row.values[1] === clientID &&
                rowNumber !== worksheet.actualRowCount
            ) {
                newWorksheet.addRow(row.values);
            } else if (
                row.values[1] === clientID &&
                rowNumber === worksheet.actualRowCount
            ) {
                newWorksheet.addRow(row.values);
                newWorkbook.xlsx
                    .writeFile(`./Folder/${clientName} (${clientID}).xlsx`)
                    .then(function () {
                        null;
                    });
            } else {
                newWorkbook.xlsx
                    .writeFile(`./Folder/${clientName} (${clientID}).xlsx`)
                    .then(function () {
                        null;
                    });
                clientName = worksheet.getRow(rowNumber).values[2];
                clientID = worksheet.getRow(rowNumber).values[1];
                newWorkbook = new Excel.Workbook();
                newWorksheet = newWorkbook.addWorksheet("Impact");
                newWorksheet.addRow(RowHeader);
                newWorksheet.addRow(worksheet.getRow(rowNumber).values);
            }
        }
    });
});