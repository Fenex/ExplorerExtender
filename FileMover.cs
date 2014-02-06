using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using Microsoft.VisualBasic.FileIO;

namespace ExplorerExtender
{
    public static class FileMover
    {
        private static string GetListFileName()
        {
            return Path.Combine(Path.GetDirectoryName(typeof(SettingHelper).Assembly.Location), "file.txt");
        }

        public static int GetListCount()
        {
            string filename = GetListFileName();
            if (File.Exists(filename))
                return File.ReadAllLines(filename).Where(i => !string.IsNullOrWhiteSpace(i) && (File.Exists(i) || Directory.Exists(i))).Count();
            else
                return 0;
        }

        public static List<string> GetList()
        {
            string filename = GetListFileName();

            if (File.Exists(filename))
                return File.ReadAllLines(filename).Where(i => !string.IsNullOrWhiteSpace(i) && (File.Exists(i) || Directory.Exists(i))).Distinct().OrderBy(i => i).ToList();
            else
                return new List<string>();
        }

        public static void ClearList()
        {
            string filename = GetListFileName();
            if (File.Exists(filename))
                try { File.Delete(filename); }
                catch { }
        }

        public static void AddPathToList(string file)
        {
            string filename = GetListFileName();

            if(File.Exists(filename))                
                using (StreamWriter streamWriter = File.AppendText(filename))
                    streamWriter.WriteLine(file);
            else
                using (StreamWriter streamWriter = new StreamWriter(filename))
                    streamWriter.WriteLine(file);
        }

        public static void RemovePathFromList(string file)
        {
            string filename = GetListFileName();

            if (!File.Exists(filename))
                return;

            var listOfFile = File.ReadAllLines(filename).Where(i => !string.IsNullOrWhiteSpace(i) && (File.Exists(i) || Directory.Exists(i)) && !i.Equals(file, StringComparison.InvariantCultureIgnoreCase)).OrderBy(i => i).ToList();

            File.Delete(filename);
            using (StreamWriter streamWriter = new StreamWriter(filename))
                listOfFile.ForEach(i => streamWriter.WriteLine(i));
        }

        public static void ExecuteListAction(string folder)
        {
            string errorMessage = string.Empty;

            string filename = GetListFileName();
            var listOfFile = File.ReadAllLines(filename).Where(i => !string.IsNullOrWhiteSpace(i) && File.Exists(i)).OrderBy(i => i).ToList();

            listOfFile.ForEach(i => {
                if (File.Exists(i) && !File.Exists(Path.Combine(folder, Path.GetFileName(i))))
                    try
                    {
                        FileSystem.MoveFile(i, Path.Combine(folder, Path.GetFileName(i)), UIOption.AllDialogs);
                        //File.Move(i, Path.Combine(folder, Path.GetFileName(i)));
                    }
                    catch
                    {
                        errorMessage += "\n" + i;
                    }
                else if (Directory.Exists(i))
                    try
                    {
                        FileSystem.MoveDirectory(i, Path.Combine(folder, Path.GetDirectoryName(i)), UIOption.AllDialogs);
                    }
                    catch 
                    {
                        errorMessage += "\n" + i;
                    }
                else
                    errorMessage += "\n" + i;
            });

            if (!string.IsNullOrWhiteSpace(errorMessage))
            {
                errorMessage = "The following file(s) cannot be moved" + errorMessage;
                System.Windows.Forms.MessageBox.Show(errorMessage);
            }

            ClearList();
        }
    }
}
