using System;
using System.IO;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public static class ShapeInserter
    {
        public static void CopyShapesToClipboard(string headerPptxPath)
        {
            var app = Globals.Application;
            if (app == null) return;

            PowerPoint.Presentation srcPres = null;
            try
            {
                srcPres = app.Presentations.Open(
                    headerPptxPath,
                    MsoTriState.msoTrue,
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);

                var slide = srcPres.Slides[1];
                int count = slide.Shapes.Count;
                if (count == 0) return;

                var indices = new int[count];
                for (int i = 0; i < count; i++)
                    indices[i] = i + 1;

                slide.Shapes.Range(indices).Copy();
            }
            finally
            {
                if (srcPres != null)
                {
                    srcPres.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(srcPres);
                }
            }
        }

        public static void InsertToActiveSlide(string headerPptxPath)
        {
            var app = Globals.Application;
            if (app == null) return;

            try
            {
                var window = app.ActiveWindow;
                if (window == null) return;

                CopyShapesToClipboard(headerPptxPath);

                var slide = (PowerPoint.Slide)window.View.Slide;
                slide.Shapes.Paste();
            }
            catch (Exception ex)
            {
                System.Windows.Forms.MessageBox.Show(
                    $"Shape 삽입 실패:\n{ex.Message}",
                    "TEAMPPT",
                    System.Windows.Forms.MessageBoxButtons.OK,
                    System.Windows.Forms.MessageBoxIcon.Warning);
            }
        }

        public static void GenerateThumbnail(string headerPptxPath, string outputPngPath)
        {
            var app = Globals.Application;
            if (app == null)
            {
                TaskPaneHost.LogDebug("GenerateThumbnail: app is null");
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
                    TaskPaneHost.LogDebug($"Shape-only export OK: {Path.GetFileName(outputPngPath)}");
                }
                catch (Exception ex)
                {
                    TaskPaneHost.LogDebug($"Shape-only export failed, fallback to slide: {ex.Message}");
                    slide.Export(outputPngPath, "PNG", 480, 270);
                }
            }
            catch (Exception ex)
            {
                TaskPaneHost.LogDebug($"GenerateThumbnail FAILED: {ex.Message}");
                throw;
            }
            finally
            {
                if (pres != null)
                {
                    pres.Close();
                    System.Runtime.InteropServices.Marshal.ReleaseComObject(pres);
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
