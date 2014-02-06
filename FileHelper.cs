using System.IO;
using System.Web;
using Microsoft.VisualBasic.FileIO;

namespace ExplorerExtender
{
    public static class FileHelper
    {
        public static void ReplaceFileName(this FileInfo fileInfo)
        {
            try
            {
                if (fileInfo == null || !fileInfo.Exists)
                    return;

                string pair = SettingHelper.ReadSetting("ReplacePair");

                if (string.IsNullOrWhiteSpace(pair))
                    return;

                string[] pairs = pair.Split(new char[] { '|' });

                if (pairs.Length % 2 != 0)
                    return;

                string filename = fileInfo.Name;

                for (int i = 0; i <= pairs.Length; i += 2)
                    filename = filename.Replace(pairs[i], pairs[i + 1]);

                filename = Path.Combine(fileInfo.DirectoryName, filename);

                if (!File.Exists(filename))
                    FileSystem.MoveFile(fileInfo.FullName, filename, UIOption.AllDialogs);
                    //fileInfo.MoveTo(filename);
            }
            catch { }
        }

        public static void DecodeFileNameUrl(this FileInfo fileInfo)
        {
            try
            {
                if (fileInfo == null || !fileInfo.Exists)
                    return;

                string filename = fileInfo.Name;

                filename = HttpUtility.UrlDecode(filename);

                filename = Path.Combine(fileInfo.DirectoryName, filename);

                if (!File.Exists(filename))
                    FileSystem.MoveFile(fileInfo.FullName, filename, UIOption.AllDialogs);
                    //fileInfo.MoveTo(filename);
            }
            catch { }
        }

        public static void DecodeFileNameHtml(this FileInfo fileInfo)
        {
            try
            {
                if (fileInfo == null || !fileInfo.Exists)
                    return;

                string filename = fileInfo.Name;

                filename = HttpUtility.HtmlDecode(filename);

                filename = Path.Combine(fileInfo.DirectoryName, filename);

                if (!File.Exists(filename))
                    FileSystem.MoveFile(fileInfo.FullName, filename, UIOption.AllDialogs);
                    //fileInfo.MoveTo(filename);
            }
            catch { }
        }
    }
}
