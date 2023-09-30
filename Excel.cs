using Microsoft.Office.Interop.Excel;

namespace ExcelSpace {
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
namespace Select
{
    public static class ExcelSelect
    {
        public static void SelectSheet(Sheets options, ref int n, string message)
        {
            if (n < 1 || n > options.Count) n = 1;
            while (true)
            {
                Console.Clear();
                Console.WriteLine(message + "\n");
                foreach (Worksheet option in options)
                {
                    if(option == options[n]) { Console.BackgroundColor = ConsoleColor.White; Console.ForegroundColor = ConsoleColor.Black; }
                    Console.WriteLine(option.Name);
                    Console.ResetColor();

                }
                ConsoleKey input = Console.ReadKey().Key;
                if(input == ConsoleKey.Enter)
                {
                    break;
                }
                switch(input)
                {
                    case ConsoleKey.UpArrow: if(n > 1) n--; break;
                    case ConsoleKey.DownArrow: if(n < options.Count) n++; break;
                    case ConsoleKey.LeftArrow: n = 1; break;
                    case ConsoleKey.RightArrow: n = options.Count; break;
                }
            }
        }
    }
}