namespace PseudoExcelReader
{
    public static class Funcs
    {
        public static string SelectFile(string currentDir, string message, string fileType = ".txt")
        {
            List<string> previousDir = new List<string>();
            List<string> options = new List<string>();
            int n = 0;
            string path;
            Console.CursorVisible = false;
            bool newDir = true;
            while (true)
            {
                List<string> dirs = Directory.GetDirectories(currentDir).ToList();
                List<string> files = Directory.GetFiles(currentDir).ToList();
                List<string> remove = new List<string>();
                foreach(string file in files)
                {
                    if (!file.Contains(fileType)) remove.Add(file);
                }
                foreach(string item in remove)
                {
                    files.Remove(item);
                }
                options.Clear();
                options.AddRange(dirs.ToArray());
                options.AddRange(files.ToArray());
                Console.Clear();
                Console.WriteLine(message + "\n");
                foreach (string option in options)
                {
                    if (option == options[n]) { Console.BackgroundColor = ConsoleColor.White; Console.ForegroundColor = ConsoleColor.Black; }
                    Console.WriteLine(option.Replace(currentDir, ""));
                    Console.ResetColor();
                    if (newDir == true) Thread.Sleep(2);
                }
                newDir = false;
                ConsoleKey input = Console.ReadKey().Key;
                switch (input)
                {
                    case ConsoleKey.UpArrow: if (n > 0) n--; break;
                    case ConsoleKey.DownArrow: if (n < options.Count - 1) n++; break;
                    case ConsoleKey.LeftArrow: n = 0; break;
                    case ConsoleKey.RightArrow: n = options.Count - 1; break;

                    case ConsoleKey.Enter: 
                        string end = options[n].Substring(options[n].Length - fileType.Length);
                        if(end == fileType)
                        {
                            path = options[n];
                            Console.CursorVisible = true;
                            return path;
                        }
                        else
                        {
                            if (!options[n].Contains('.'))
                            {
                                previousDir.Add(currentDir);
                                currentDir = options[n];
                                n = 0;
                                newDir = true;
                            }
                        }
                        break;

                    case ConsoleKey.Escape:
                        if (previousDir.Count > 0)
                        {
                            currentDir = previousDir.Last();
                            previousDir.RemoveAt(previousDir.Count - 1);
                            n = 0;
                        }
                        break;
                }
            }
        }

        public static List<LærerPar> GetPairs(string path)
        {
            List<LærerPar> lærerPar = new List<LærerPar> ();
            lærerPar.Clear();
            string content = File.ReadAllText(path);
            List<string> rows = content.Split('\n').ToList<string>();
            List<List<string>> cols = new List<List<string>>();
            foreach (string row in rows)
            {
                string[] col = row.Split('\t');
                lærerPar.Add(new LærerPar(Int32.Parse(col[2].Trim()), col[0], col[1]));
            }
            return lærerPar;
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
        public LærerPar() { }
        
    }
}