const ExcelJS = require('exceljs'); // import ExcelJS class from exceljs

async function excelTest() {
    const workbook = new ExcelJS.Workbook();
    await workbook.xlsx.readFile("D:\\Udemy\\ExcelJSUtil\\excelDownloadTest.xlsx");
    const worksheet = workbook.getWorksheet('Sheet1');// hold the data of Sheet1

    //read and print data in the worksheet
    worksheet.eachRow((row, rowNumber) => {
        row.eachCell((cell, colNumber) => {
            console.log(cell.value);
        }
        )
    })
}

excelTest();

