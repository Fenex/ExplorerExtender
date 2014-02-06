using System.Collections.Generic;
using System.IO;
using System.Text;
using System.Threading.Tasks;

namespace ExplorerExtender
{
    public static class FolderHelper
    {
        private static DirectoryInfo GetParentFolder(string path)
        {
            if (Directory.Exists(path))
                return (new DirectoryInfo(path)).Parent;
            else if (File.Exists(path))
                return (new FileInfo(path)).Directory;
            else
                return null;
        }

        private static string GetNewFolderName(string path)
        {
            if (Directory.Exists(path))
                return (new DirectoryInfo(path)).Name;
            else if (File.Exists(path))
                return Path.GetFileNameWithoutExtension(path);
            else
                return null;
        }

        private static void MoveFolder(string source, string dest)
        {
            if (!Directory.Exists(source))
                return;

            if (!Directory.Exists(dest))
                Directory.Move(source, dest);
            else
            {
                DirectoryInfo diSource = new DirectoryInfo(source);

                var task1 = Task.Factory.StartNew(() => {
                    Parallel.ForEach(diSource.GetDirectories(), diSub => MoveFolder(diSub.FullName, Path.Combine(dest, diSub.Name)));
                });

                var task2 = Task.Factory.StartNew(() =>{
                    Parallel.ForEach(diSource.GetFiles(), fi =>
                    {
                        if (!File.Exists(Path.Combine(dest, fi.Name)))
                            fi.MoveTo(Path.Combine(dest, fi.Name));
                    });                  
                });

                Task.WaitAll(task1, task2);

                if (diSource.GetFiles().Length == 0 && diSource.GetDirectories().Length == 0)
                    diSource.Delete();
            }
        }

        public static void BreakFolder(List<string> SelectedItems)
        {
            var error = new List<string>();
             
            foreach (string s in SelectedItems)
                if (Directory.Exists(s))
                {
                    DirectoryInfo di = new DirectoryInfo(s);

                    var task1 = Task.Factory.StartNew(() => {
                        Parallel.ForEach(di.GetFiles(), fi =>
                        {
                            string newPath = Path.Combine(di.Parent.FullName, fi.Name);
                            if (!File.Exists(newPath))
                                fi.MoveTo(newPath);
                        });
                    });

                    var task2 = Task.Factory.StartNew(() => {
                        Parallel.ForEach(di.GetDirectories(), diSub =>
                        {
                            string newPath = Path.Combine(di.Parent.FullName, diSub.Name);
                            MoveFolder(diSub.FullName, newPath);
                        });
                    });

                    Task.WaitAll(task1, task2);                    
                    
                    if (di.GetFiles().Length == 0 && di.GetDirectories().Length == 0)
                        try { di.Delete(); }
                        catch { error.Add(di.FullName); }
                }

            if (error.Count > 0)
            {
                StringBuilder sbError = new StringBuilder();
                sbError.AppendLine("Error deleting the following folder");
                error.ForEach(e => sbError.AppendLine(e));
                System.Windows.Forms.MessageBox.Show(sbError.ToString());
            }
        }
        
        public static void BuildFolder(List<string> SelectedItems)
        {
            DirectoryInfo parent = GetParentFolder(SelectedItems[0]);

            if (parent == null)
                return;

            string newFolder = Path.GetRandomFileName();
            while (Directory.Exists(Path.Combine(parent.FullName, newFolder)))
                newFolder = Path.GetRandomFileName();

            DirectoryInfo diNewFolder = Directory.CreateDirectory(Path.Combine(parent.FullName, newFolder));

            string newFolderName = GetNewFolderName(SelectedItems[0]);

            Parallel.ForEach(SelectedItems, s => {
                if (Directory.Exists(s))
                    Directory.Move(s, Path.Combine(diNewFolder.FullName, new DirectoryInfo(s).Name));
                else if (File.Exists(s))
                    File.Move(s, Path.Combine(diNewFolder.FullName, Path.GetFileName(s)));
            });

            if (!string.IsNullOrEmpty(newFolderName))
                diNewFolder.MoveTo(Path.Combine(diNewFolder.Parent.FullName, newFolderName));
        }
    
    
    }
}
