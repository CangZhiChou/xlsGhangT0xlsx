// See https://aka.ms/new-console-template for more information


using Magicodes.ExporterAndImporter.Core;
using Magicodes.ExporterAndImporter.Excel;
using NPOI.HSSF.UserModel;
using NPOI.XSSF.UserModel;
using 将.xls转换为.xlsx;
using 将.xls转换为.xlsx.FixtureDataImportFromExcel.Common;
using NPOI.HSSF.UserModel;
using NPOI.SS.UserModel;
using System.IO;


var xlsPath = @"C:\Users\Lenovo\Desktop\新建文件夹 (8)\产品资讯.xls";
var   newExcelPath= @"C:\Users\Lenovo\Desktop\新建文件夹 (8)\产品资讯.xlsx";
var oldWorkbook = new HSSFWorkbook(new FileStream(xlsPath, FileMode.Open));
var oldWorkSheet = oldWorkbook.GetSheetAt(0);

using (var fileStream = new FileStream(newExcelPath, FileMode.Create))
{
    var newWorkBook1 = new XSSFWorkbook();
    var sheet = oldWorkSheet.CrossCloneSheet(newWorkBook1, "Sheet1");
    newWorkBook1.Add(sheet);
    newWorkBook1.Write(fileStream);
    newWorkBook1.Close();

}



IImporter Importer = new ExcelImporter ();
var import = await Importer.Import<Class1>(newExcelPath);


//这是将IFormFile转换成HSSFWorkbook
//public void ReadExcel(IFormFile file)
//{
//    using (var stream = file.OpenReadStream())
//    {
//        var workbook = new HSSFWorkbook(stream);
//        var sheet = workbook.GetSheetAt(0);
//        for (int i = sheet.FirstRowNum; i <= sheet.LastRowNum; i++)
//        {
//            var row = sheet.GetRow(i);
//            if (row == null) continue;
//            for (int j = row.FirstCellNum; j < row.LastCellNum; j++)
//            {
//                var cell = row.GetCell(j);
//                if (cell == null) continue;
//                var value = cell.ToString();
//                // 处理单元格数据
//            }
//        }
//    }
//}
Console.WriteLine("Hello, World!");
