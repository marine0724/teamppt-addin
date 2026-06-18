using System.IO;
using System.Reflection;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public static class Globals
    {
        public static PowerPoint.Application Application { get; set; }

        public static string AssetsDir
        {
            get
            {
                var dir = Path.GetDirectoryName(Assembly.GetExecutingAssembly().Location);
                return Path.Combine(dir, "Assets");
            }
        }

        public static string ThumbnailDir
        {
            get
            {
                var dir = Path.Combine(
                    System.Environment.GetFolderPath(System.Environment.SpecialFolder.LocalApplicationData),
                    "TeampptAddin", "thumbnails");
                if (!Directory.Exists(dir))
                    Directory.CreateDirectory(dir);
                return dir;
            }
        }
    }
}
