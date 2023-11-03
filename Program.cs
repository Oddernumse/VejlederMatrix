using ExcelSpace;
using Select;
using Excel = Microsoft.Office.Interop.Excel;
using System.Runtime.InteropServices;

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
int dataCol = dataRange.Rows.Count + 1;
while (dataCol > dataRange.Rows.Count)
{
    Console.WriteLine("\nOg i hvilken kolonne (nr.) findes den første række vejledere?");
    dataCol = Int32.Parse(Console.ReadLine());
}

for(i = 1; dataRange[i, dataCol].Value2 != null || tilgængeligeLærere.Count == 0; i++)
{
    object lærer1 = dataRange.Cells[i, dataCol].Value2;
    if (lærer1.ToString().Length <= 4)
    {
        object lærer2 = dataRange.Cells[i, dataCol + 1].Value2;
        string mødeString = dataRange.Cells[i, dataCol + 2].Value2.ToString();
        int møder = Int32.Parse(mødeString);
        if (!tilgængeligeLærere.ContainsKey(lærer1.ToString())) tilgængeligeLærere.Add(lærer1.ToString(), true);
        if (!tilgængeligeLærere.ContainsKey(lærer2.ToString())) tilgængeligeLærere.Add(lærer2.ToString(), true);
        lærerPar.Add(new LærerPar(møder, lærer1.ToString(), lærer2.ToString()));
    }
}

// Lav en plan:
bool planlægning = false;
for(int blok = 0; !planlægning; blok++)
{
    planlægning = true;
    plan.Add(new Dictionary<LærerPar, bool> ());

    // Gør alle lærere tilgængelige igen:
    foreach(KeyValuePair<string, bool> pair in tilgængeligeLærere)
    {
        tilgængeligeLærere[pair.Key] = true;
    }

    // Lav en blok
    plan[blok].Clear();
    while(true)
    {
        LærerPar prioPar = new LærerPar(-1, "", "");
        foreach (LærerPar par in lærerPar)
        {
            try
            {
                plan[blok].Add(par, false);
            } catch
            {

            }
        }
        foreach (LærerPar par in lærerPar)
        {
            if (par.møder > prioPar.møder && par.møder > 0 && tilgængeligeLærere[par.lærer1] && tilgængeligeLærere[par.lærer2])
            {
                prioPar = par;
            }
        }
        if(prioPar.møder == -1) break;
        else
        {
            tilgængeligeLærere[prioPar.lærer1] = false;
            tilgængeligeLærere[prioPar.lærer2] = false;
            plan[blok][prioPar] = true;
            lærerPar[lærerPar.IndexOf(prioPar)].møder--;
        }
    }

    // Tjek om alle lærere er blevet tildelt møder:
    foreach(LærerPar par in lærerPar)
    {
        if(par.møder > 0) planlægning = false;
    }
}

// Print planen:
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
    printRange[i, 1].Clear();
    printRange[i, 1] = par.lærer1;
    printRange[i, 2].Clear();
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
            printRange[row, i + 3].Clear();
            printRange[row, i + 3].Value = "Optaget";
        }
        row++;
    }
}

// -------------- Prøv at hoppe af appen (DRÆB PLANEN): --------------
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