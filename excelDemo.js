const ExcelJS = require('exceljs'); // import ExcelJS class from exceljs

const workbook = new ExcelJS.Workbook();
workbook.xlsx.readFile("D:\\Udemy\\ExcelJSUtil\\excelDownloadTest.xlsx").then(function(){
    const worksheet = workbook.getWorksheet('Sheet1');// hold the data of Sheet1

//read and print data in the worksheet
worksheet.eachRow( (row, rowNumber) =>
{
    row.eachCell((cell, colNumber) =>
    {
        console.log(cell.value);
    }
    )
})
});

