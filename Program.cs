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

// Åbn excel-filen (der læses og skrives fra samme fil):
Excel.Application xlApp = new Excel.Application();
string fileName = "C:\\VSC_PRO_B\\VejlederMatrix\\VejlederMatrix.xlsx"; // Vil gerne at nikolaj selv kan vælge filen
ExcelFileHandling exFileHandler  = new ExcelFileHandling(fileName, xlApp);

// Selve programmet:
Console.Clear();
Dictionary<string, bool> tilgængeligeLærere = new Dictionary<string, bool>();
List<LærerPar> lærerPar = new List<LærerPar> ();
List<Dictionary<string, bool>> plan = new List<Dictionary<string, bool>>();

Console.WriteLine("På hvilket ark (nr.) ligger dataene?");
int dataSheet = Int32.Parse(Console.ReadLine());
Console.WriteLine("\nOg under hvilken overskrift findes den første række vejledere?");
string startCelleIndhold = Console.ReadLine();

int[] startCelle = exFileHandler.Find(1, startCelleIndhold);
for (int i = 0; i < 2; i++)
{
    List<string> forkortelser = exFileHandler.GetExcelColumn(dataSheet, startCelle[1] + i);
    foreach (string f in forkortelser)
    {
        if (tilgængeligeLærere.ContainsKey(f) || f == null || f.Length >= 4) continue;
        else
        {
            tilgængeligeLærere.Add(f, true);
            Console.WriteLine(f);
        }
    }
}

// Prøv at hoppe af appen:
xlApp.Quit();
Marshal.ReleaseComObject(xlApp);