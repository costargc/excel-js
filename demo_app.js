var Excel = require("exceljs");

var RowHeader = [
    "Date",
    "Open",
    "High",
    "Low",
    "Close*",
    "Adj Close**",
    "Volume"
];

var workbook = new Excel.Workbook();

workbook.xlsx.readFile("./data_demo.xlsx")
    .then(() => {
        // console.log(workbook);

        var worksheet = workbook.getWorksheet("data");
        // console.log(worksheet.getRow(2).values[5]);
        // console.log(worksheet.getRow(2).values[5]);
        console.log(worksheet._rows.length);
        max_row = worksheet._rows.length;

        console.log(worksheet._columns.length);
        max_col = worksheet._columns.length;


        for (col = 2; col <= max_col; col++) {
            sum = 0;
            for (row = 2; row <= max_row; row++) {
                sum = sum + worksheet.getRow(row).getCell(col).value;
            }

            console.log(sum);
            worksheet.getRow(1).getCell(col).value = sum;
            workbook.xlsx.writeFile('./data_demo.xlsx');
        }


    });