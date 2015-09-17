var Excel = require("exceljs");
//var win32ole = require('win32ole');

var workbook = null;
//var xl = null;
var wshshell = null;

function Open_Excel(filename){

    var shtTmp = null;

    workbook = new Excel.Workbook();

    workbook.xlsx.readFile(filename).then(function() {

        // 先頭のワークシートを取得
        shtTmp =  workbook.getWorksheet(1);

        // セルに値を設定
        shtTmp.getCell("C3").value = 123;

        // ファイルを書込
        workbook.xlsx.writeFile(filename).then(function() {
            // done
        });

        workbook.eachSheet(function (worksheet, sheetId) {
        console.log(worksheet.name);
        }); 

    });



    // Excelファイルを開く
//    window.open (filename);

    wshshell = new ActiveXObject("WScript.Shell"); 
    wshshell.run("excel " + filename); 


    // Excelの起動
//    xl = new win32ole.client.Dispatch('Excel.Application');
//    xl.Visible = true;


};

// pipe from stream
//var workbook = new Excel.Workbook();
//stream.pipe(workbook.xlsx.createInputStream());


