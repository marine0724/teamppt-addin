using System;
using System.IO;

namespace TeampptAddin
{
    /// <summary>
    /// 디버그 로깅 유틸리티.
    /// 로그 파일: %LocalAppData%\TeampptAddin\debug.log
    /// 형식: [HH:mm:ss.fff] message
    /// 모든 예외를 삼키므로 어디서든 안전하게 호출 가능.
    /// </summary>
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
