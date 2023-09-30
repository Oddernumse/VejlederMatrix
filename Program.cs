using ExcelSpace;
using Select;
using Excel = Microsoft.Office.Interop.Excel;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Runtime.InteropServices;
using System.Text;
using System.Threading.Tasks;
using static System.Reflection.Metadata.BlobBuilder;
using Microsoft.Office.Interop.Excel;

// -------------- Åbn excel-filen (der læses og skrives fra samme fil): --------------
Excel.Application xlApp = new Excel.Application();
string fileName = ExcelSelect.SelectFile("C:\\", "Vælg venligst Excel-filen");
Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);

// -------------- Selve programmet: --------------
Console.Clear();
Console.ResetColor();
int i;
Dictionary<string, bool> tilgængeligeLærere = new Dictionary<string, bool>();
List<LærerPar> lærerPar = new List<LærerPar> ();
List<Dictionary<LærerPar, bool>> plan = new List<Dictionary<LærerPar, bool>>();

// Find arket med dataene
int arkN = 1;
ExcelSelect.SelectSheet(xlWorkbook.Sheets, ref arkN, "På hvilket ark ligger dataene?");
Excel.Worksheet datasheet = xlWorkbook.Worksheets[arkN];
Excel.Range dataRange = datasheet.UsedRange;
Console.WriteLine("\nOg i hvilken kolonne (nr.) findes den første række vejledere?");
int dataColumn1 = Int32.Parse(Console.ReadLine()) - 2;

for(i = 1; dataRange.Cells[i, dataColumn1].Value2 != null || tilgængeligeLærere.Count == 0 ; i++)
{
    object lærer1 = dataRange.Cells[i, dataColumn1].Value2;
    if (lærer1.ToString().Length <= 4)
    {
        object lærer2 = dataRange.Cells[i, dataColumn1 + 1].Value2;
        string mødeString = dataRange.Cells[i, dataColumn1 + 2].Value2.ToString();
        int møder = Int32.Parse(mødeString);
        if (!tilgængeligeLærere.ContainsKey(lærer1.ToString())) tilgængeligeLærere.Add(lærer1.ToString(), true);
        lærerPar.Add(new LærerPar(møder, lærer1.ToString(), lærer2.ToString()));
    }
}

// Make a plan:


// Print the plan:
Excel.Worksheet printSheet = datasheet;
bool printExists = false;
foreach(Excel.Worksheet sheet in xlWorkbook.Sheets)
{
    if(sheet.Name == "Resultat")
    {
        printExists = true;
        printSheet = sheet;
        break;
    }
}
if (!printExists) { xlWorkbook.Sheets.Add(); int printSheetN = xlWorkbook.Sheets.Count - 1; printSheet = xlWorkbook.Sheets[printSheetN]; printSheet.Name = "Resultat"; }
Excel.Range printRange = printSheet.UsedRange;

printRange[1, 1].Value = "Vejleder 1:";
printRange[1, 2].Value = "Vejleder 2:";

i = 2;
foreach (LærerPar par in lærerPar)
{
    printRange[i, 1] = par.lærer1;
    printRange[i, 2] = par.lærer2;
    i++;
}

Console.WriteLine("Møder: " + plan.Count);

for(i  = 0; i < plan.Count; i++)
{
    int row = 2;
    foreach(LærerPar planpar in plan[i].Keys)
    {
        if (plan[i][planpar] == true)
        {
            printRange[row, i + 3].Value = $"Optaget ({planpar.lærer1} og {planpar.lærer2})";
        }
        row++;
    }
}

// -------------- Prøv at hoppe af appen (KILL THE PLAN): --------------
foreach (Excel.Workbook book in xlApp.Workbooks)
{
    foreach(Excel.Worksheet sheet in book.Worksheets)
    {
        Marshal.ReleaseComObject(sheet);
    }
    book.Save();
    book.Close();
    Marshal.ReleaseComObject(book);
}
xlApp.Quit();
Marshal.ReleaseComObject(xlApp);