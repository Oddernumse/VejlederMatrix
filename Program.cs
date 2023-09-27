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

ExcelFileHandling exFileReader  = new ExcelFileHandling(fileName, xlApp);
List<string> rowFive = exFileReader.GetExcelRow(5, 5);

// Prøv at hoppe af appen:
xlApp.Quit();
Marshal.ReleaseComObject(xlApp);