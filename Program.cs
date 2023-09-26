using ExcelSpace;
using Excel = Microsoft.Office.Interop.Excel;
using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static System.Reflection.Metadata.BlobBuilder;

Excel.Application xlApp = new Excel.Application();

string fileName = "C:\\VSC_PRO_B\\VejlederMatrix\\VejlederMatrix.xlsx";

ExcelFileReading exFileReader  = new ExcelFileReading(fileName, xlApp);
int[] coords = exFileReader.Find(1, "Vejleder1");

Console.WriteLine($"Række: {coords[0]}\nKolonne: {coords[1]}");

// Prøv at hoppe af appen:
Excel.Workbooks books = null;
Excel.Workbook book = null;
Excel.Sheets sheets = null;
Excel.Worksheet sheet = null;
Excel.Range range = null;

try
{
    books = xlApp.Workbooks;
    book = books.Add();
    sheets = book.Sheets;
    sheet = sheets.Add();
    range = sheet.Range["A1"];
    range.Value = "Lorem Ipsum";
    book.SaveAs(@"C:\Temp\ExcelBook" + DateTime.Now.Millisecond + ".xlsx");
    book.Close();
    xlApp.Quit();
}
finally
{
    if (range != null) Marshal.ReleaseComObject(range);
    if (sheet != null) Marshal.ReleaseComObject(sheet);
    if (sheets != null) Marshal.ReleaseComObject(sheets);
    if (book != null) Marshal.ReleaseComObject(book);
    if (books != null) Marshal.ReleaseComObject(books);
    if (xlApp != null) Marshal.ReleaseComObject(xlApp);
}