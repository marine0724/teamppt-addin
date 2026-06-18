using System;
using System.IO;

namespace TeampptAddin
{
    public static class Logger
    {
        public static void Log(string msg)
        {
            try
            {
                var dir = Path.Combine(
                    Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                    "TeampptAddin");
                Directory.CreateDirectory(dir);
                File.AppendAllText(
                    Path.Combine(dir, "debug.log"),
                    $"[{DateTime.Now:HH:mm:ss.fff}] {msg}\r\n");
            }
            catch { }
        }
    }
}
