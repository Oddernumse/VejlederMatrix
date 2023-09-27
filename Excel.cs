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

        ArgumentOutOfRangeException tooFewSheets = new ArgumentOutOfRangeException("", "The desired sheet does not exist!");

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
                    if (xlRange.Cells[i, col].Value2 != null) column.Add(xlRange.Cells[i, col].Value2.ToString());
                    else continue;
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
                Excel.Range location = xlWorkbook.Worksheets[sheet].UsedRange.Find(lookFor);
                coords[0] = location.Row - 1;
                coords[1] = location.Column - 1;
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
                xlWorkbook.Worksheets[sheet].Cells[coords[0], coords[1]].Value = value;
                xlWorkbook.Save();
                return;
            }
            else
            {
                xlWorkbook.Sheets.Add();
                WriteCell(sheet, coords, value);
            }
        }
        public void WriteColumn(int sheet, int[] startCell, List<object> values)
        {
            if (xlWorkbook.Sheets.Count >= sheet)
            {
                for(int i = 0; i < values.Count; i++)
                {
                    WriteCell(sheet, new int[] { i + startCell[0], startCell[1] }, values[i]);
                }
                return;
            }
            else
            {
                xlWorkbook.Sheets.Add();
                WriteColumn(sheet, startCell, values);
            }
        }
        public void WriteColumn(int sheet, int[] startCell, List<object> values, bool reversed)
        {
            if (xlWorkbook.Sheets.Count >= sheet && reversed)
            {
                WriteColumn(sheet, startCell, values);
                return;
            }
            else if(xlWorkbook.Sheets.Count >= sheet)
            {
                for (int i = values.Count - 1; i >= 0; i--)
                {
                    WriteCell(sheet, new int[] { i + startCell[0], startCell[1] }, values[i]);
                }
                return;
            }
            else
            {
                xlWorkbook.Sheets.Add();
                WriteColumn(sheet, startCell, values, reversed);
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
    public class LærerPar
    {
        public int møder;
        public string lærer1;
        public string lærer2;
        public LærerPar(int møder, string lærer1, string lærer2)
        {
            this.møder = møder;
            this.lærer1 = lærer1;
            this.lærer2 = lærer2;
        }
    }
}