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
List<Dictionary<LærerPar, bool>> plan = new List<Dictionary<LærerPar, bool>>();

Console.WriteLine("På hvilket ark (nr.) ligger dataene?");
int dataSheetN = Int32.Parse(Console.ReadLine());
Excel.Worksheet datasheet = xlWorkbook.Sheets[dataSheetN];
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
    plan.Add(new Dictionary<LærerPar, bool> ());
    planlægning = true;
    bool blokDone = false;
    foreach (LærerPar par in lærerPar)
    {
        plan[blok].Add(par, false);
    }
    while (!blokDone)
    {
        blokDone = true;
        int maksMøder = 0;
        int lærerN = 0;
        foreach (LærerPar par in lærerPar)
        {
            if (par.møder > maksMøder && !tilgængeligeLærere[par.lærer1] && !tilgængeligeLærere[par.lærer2])
            {
                maksMøder = par.møder;
                lærerN= lærerPar.IndexOf(par);
                planlægning = false;
                blokDone = false;
            }
        }

        plan[blok][lærerPar[lærerN]] = true;

        foreach (string fork in tilgængeligeLærere.Keys)
        {
            if (fork == lærerPar[lærerN].lærer1 || fork == lærerPar[lærerN].lærer2)
            {
                tilgængeligeLærere[fork] = false;
            }
        }
    }
}

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

int i = 2;
foreach (LærerPar par in lærerPar)
{

    i++;
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