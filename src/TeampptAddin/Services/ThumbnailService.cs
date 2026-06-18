using System;
using System.Drawing;
using System.IO;
using System.IO.Compression;

namespace TeampptAddin
{
    internal static class ThumbnailService
    {
        public static Image LoadThumbnail(string pptxPath, string cachePath)
        {
            if (File.Exists(cachePath))
            {
                try
                {
                    if (File.GetLastWriteTime(cachePath) >= File.GetLastWriteTime(pptxPath))
                        return LoadImageNoLock(cachePath);
                    File.Delete(cachePath);
                }
                catch { }
            }

            try
            {
                ThumbnailGenerator.Generate(pptxPath, cachePath);
                if (File.Exists(cachePath))
                    return LoadImageNoLock(cachePath);
            }
            catch (Exception ex)
            {
                Logger.Log($"COM thumb fail [{Path.GetFileName(pptxPath)}]: {ex.Message}");
            }

            try
            {
                using (var zip = ZipFile.OpenRead(pptxPath))
                {
                    foreach (var entry in zip.Entries)
                    {
                        if (!entry.FullName.StartsWith("docProps/thumbnail", StringComparison.OrdinalIgnoreCase))
                            continue;
                        using (var stream = entry.Open())
                        using (var ms = new MemoryStream())
                        {
                            stream.CopyTo(ms);
                            return Image.FromStream(new MemoryStream(ms.ToArray()));
                        }
                    }
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"ZIP thumb fail [{Path.GetFileName(pptxPath)}]: {ex.Message}");
            }

            return null;
        }

        public static Image LoadImageNoLock(string path)
        {
            return Image.FromStream(new MemoryStream(File.ReadAllBytes(path)));
        }
    }
}
