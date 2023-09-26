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
exFileReader.Release();

Console.WriteLine($"Række: {coords[0]}\nKolonne: {coords[1]}");

// Prøv at hoppe af appen:
xlApp.Quit();
Marshal.ReleaseComObject(xlApp);