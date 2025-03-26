const ExcelJS = require('exceljs'); // import ExcelJS class from exceljs

async function writeExcelTest(searchText, replaceText, change, filePath) {
    
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile(filePath);
    const worksheet = workbook.getWorksheet('Sheet1');// hold the data of Sheet1

    const output = await readExcel(worksheet, searchText);

    const cell = worksheet.getCell(output.row, output.column+change.colChange);
    cell.value = replaceText; // replace the cell's value
    await workbook.xlsx.writeFile(filePath); // rewrite the file with new content

}

async function readExcel(worksheet, searchText)
{
    let output = {row: -1, column: -1};
    //read and print data in the worksheet
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            if(cell.value === searchText)
            {
                output.row = rowNumber;
                output.column = colNumber;
            }
        }
        )
    })
    return output;
}

//update Mango price to 250
writeExcelTest("Mango", 350, {rowChange:0, colChange: 2}, "D:\\Udemy\\ExcelJSUtil\\downloadTest.xlsx");

