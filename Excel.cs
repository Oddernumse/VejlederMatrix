﻿using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace ExcelSpace {
    public class ExcelFileReading {
        // Properties:
        public string fileName;

        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorksheet;
        Excel.Range xlRange;

        // Constructor:
        public ExcelFileReading(string fileName, Excel.Application application) {
            //Create COM Objects. Create a COM object for everything that is referenced
            xlApp = application;
            xlWorkbook = xlApp.Workbooks.Open(fileName);
        }

        // Methods:
        public List<string> getExcelColumn(int sheet, int col) {
            xlWorksheet = xlWorkbook.Sheets[sheet];
            xlRange = xlWorksheet.UsedRange;
            int rowCount = xlRange.Rows.Count;
            int colCount = xlRange.Columns.Count;

            List<string> column = new List<string>();
            for(int i = 1; i < rowCount; i++) {
                column.Add(xlRange.Cells[i, col].Value2.ToString());
            }
            return column;
        }
        public List<string> getExcelRow(int sheet, int row)
        {
            xlWorksheet = xlWorkbook.Sheets[sheet];
            xlRange = xlWorksheet.UsedRange;
            int colCount = xlRange.Columns.Count;

            List<string> column = new List<string>();
            for (int i = 1; i < colCount; i++)
            {
                column.Add(xlRange.Cells[i, row].Value2.ToString());
            }
            return column;
        }
        public int[] Find(int sheet, string lookFor) {
            int[] coords = new int[2];
            coords[0] = xlWorkbook.Worksheets[sheet].UsedRange.Find(lookFor).Row;
            coords[1] = xlWorkbook.Worksheets[sheet].UsedRange.Find(lookFor).Column;
            return coords;
        }
    }
}