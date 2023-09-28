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
        public Excel.Workbook xlWorkbook;
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