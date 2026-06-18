using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public static class ThumbnailGenerator
    {
        public static void Generate(string headerPptxPath, string outputPngPath)
        {
            var app = Globals.Application;
            if (app == null)
            {
                Logger.Log("GenerateThumbnail: app is null");
                return;
            }
            if (File.Exists(outputPngPath)) return;

            var dir = Path.GetDirectoryName(outputPngPath);
            if (!Directory.Exists(dir))
                Directory.CreateDirectory(dir);

            PowerPoint.Presentation pres = null;
            try
            {
                pres = app.Presentations.Open(
                    headerPptxPath,
                    MsoTriState.msoTrue,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);

                var slide = pres.Slides[1];
                int count = slide.Shapes.Count;
                if (count == 0) return;

                try
                {
                    ExportShapesOnly(slide, count, outputPngPath);
                    Logger.Log($"Shape-only export OK: {Path.GetFileName(outputPngPath)}");
                }
                catch (Exception ex)
                {
                    Logger.Log($"Shape-only export failed, fallback to slide: {ex.Message}");
                    slide.Export(outputPngPath, "PNG", 480, 270);
                }
            }
            catch (Exception ex)
            {
                Logger.Log($"GenerateThumbnail FAILED: {ex.Message}");
                throw;
            }
            finally
            {
                if (pres != null)
                {
                    pres.Close();
                    Marshal.ReleaseComObject(pres);
                }
            }
        }

        private static void ExportShapesOnly(PowerPoint.Slide slide, int count, string outputPath)
        {
            var indices = new int[count];
            for (int i = 0; i < count; i++)
                indices[i] = i + 1;

            var range = slide.Shapes.Range(indices);

            if (count == 1)
            {
                range[1].Export(outputPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
            }
            else
            {
                var group = range.Group();
                group.Export(outputPath, PowerPoint.PpShapeFormat.ppShapeFormatPNG);
                group.Ungroup();
            }
        }
    }
}
