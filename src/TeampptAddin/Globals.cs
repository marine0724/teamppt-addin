using System.IO;
using System.Reflection;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 전역 공유 상태.
    /// - Application: PowerPoint COM 인스턴스 (Connect.OnConnection에서 설정)
    /// - AssetsDir: DLL과 같은 폴더의 Assets/ (header_N.pptx 파일 위치)
    /// - ThumbnailDir: %LocalAppData%\TeampptAddin\thumbnails\ (PNG 캐시)
    /// </summary>
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
