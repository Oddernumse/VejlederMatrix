using ExcelSpace;
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

// Åbn excel-filen (der læses og skrives fra samme fil):
Excel.Application xlApp = new Excel.Application();

Console.WriteLine("Hvad hedder filen? (Husk store og små bogstaver!");
string fileSearch = Console.ReadLine();
IEnumerable<string> fileNames = Directory.EnumerateFiles("C:\\VSC_PRO_B", fileSearch + ".xlsx", SearchOption.AllDirectories);
string fileName = fileNames.First();
Excel.Workbook xlWorkbook = xlApp.Workbooks.Open(fileName);

// Selve programmet:
Console.Clear();
Dictionary<string, bool> tilgængeligeLærere = new Dictionary<string, bool>();
List<LærerPar> lærerPar = new List<LærerPar> ();
List<Dictionary<string, bool>> plan = new List<Dictionary<string, bool>>();

Console.WriteLine("På hvilket ark (nr.) ligger dataene?");
int dataRangeNr = Int32.Parse(Console.ReadLine());
Excel.Worksheet datasheet = xlWorkbook.Sheets[dataRangeNr];
Excel.Range dataRange = datasheet.UsedRange;
Console.WriteLine("\nOg i hvilken kolonne (nr.) findes den første række vejledere?");
int dataColumn1 = Int32.Parse(Console.ReadLine()) - 2;

for(int i = 1; dataRange.Cells[i, dataColumn1].Value2 != null || tilgængeligeLærere.Count == 0 ; i++)
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

bool planlægning = false;
for (int blok = 0; !planlægning; blok++)
{
    LærerPar emptyPar = new LærerPar(0, "", "");
    planlægning = true;
    bool blokDone = false;
    while (!blokDone)
    {
        int maksMøder = 0;
        int n = 0;
        foreach (LærerPar par in lærerPar)
        {
            if (par.møder == 0) lærerPar.Remove(par);
            else if (par.møder > maksMøder)
            {
                maksMøder = par.møder;
                n = lærerPar.IndexOf(par);
                planlægning = false;
            }
        }
        lærerPar[n]
    }
}
// Prøv at hoppe af appen:
foreach(Excel.Workbook book in xlApp.Workbooks)
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