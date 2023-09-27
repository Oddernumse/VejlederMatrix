using Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static System.Net.Mime.MediaTypeNames;
using Excel = Microsoft.Office.Interop.Excel;       //microsoft Excel 14 object in references-> COM tab

namespace ExcelSpace {
    public class ExcelFileHandling {
        // Properties:
        Excel.Application xlApp;
        Excel.Workbook xlWorkbook;
        Excel.Worksheet xlWorksheet;
        Excel.Range xlRange;

        new ArgumentOutOfRangeException tooFewSheets = new ArgumentOutOfRangeException("", "The desired sheet does not exist!");

        // Constructor:
        public ExcelFileHandling(string fileName, Excel.Application application) {
            //Create COM Objects. Create a COM object for everything that is referenced
            xlApp = application;
            xlWorkbook = xlApp.Workbooks.Open(fileName);
        }

        // Destructor:
        ~ExcelFileHandling() {
            xlWorkbook.Save();
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
        }

        // Read methods:
        public List<string> GetExcelColumn(int sheet, int col)
        {
            if (xlWorkbook.Sheets.Count >= sheet)
            {
                xlWorksheet = xlWorkbook.Sheets[sheet];
                xlRange = xlWorksheet.UsedRange;
                int rowCount = xlRange.Rows.Count;

                List<string> column = new List<string>();
                for (int i = 1; i < rowCount; i++)
                {
                    column.Add(xlRange.Cells[i, col].Value2.ToString());
                }
                return column;
            }
            else
            {
                throw tooFewSheets;
            }
        }
        public List<string> GetExcelRow(int sheet, int row)
        {
            if (xlWorkbook.Sheets.Count >= sheet)
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
            else
            {
                throw tooFewSheets;
            }
            
        }
        public int[] Find(int sheet, string lookFor) {
            if (xlWorkbook.Sheets.Count >= sheet)
            {
                int[] coords = new int[2];
                coords[0] = xlWorkbook.Worksheets[sheet].UsedRange.Find(lookFor).Row;
                coords[1] = xlWorkbook.Worksheets[sheet].UsedRange.Find(lookFor).Column;
                return coords;
            }
            else
            {
                throw tooFewSheets;
            }
            
        }

        // Write methods:
        public void WriteCell(int sheet, int[] coords, object value)
        {
            if (xlWorkbook.Sheets.Count >= sheet)
            {
                
                return;
            }
            else
            {
                xlWorkbook.Sheets.Add();
                WriteCell(sheet, coords, value);
            }
        }
        public void WriteColumn(int sheet, int[] coords, List<object> values)
        {
            for(int i = coords[0]; i < coords[0] + values.Count; i++)
            {

            }
        }
        public void WriteColumn(int sheet, int[] coords, List<object> values, bool reversed)
        {
            if(!reversed) {WriteColumn(sheet, coords, values);}
            else
            {

            }
        }

        // Other:
        public void Release() {
            // Will probably make the instance calling this useless
            xlWorkbook.Save();
            xlWorkbook.Close();
            Marshal.ReleaseComObject(xlWorkbook);
        }
    }
}