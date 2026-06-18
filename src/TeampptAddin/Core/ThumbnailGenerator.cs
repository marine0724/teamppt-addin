using System;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// PowerPoint COM API를 사용한 에셋 썸네일 PNG 생성기.
    ///
    /// 생성 전략 (우선순위):
    /// 1. Shape-only export: 슬라이드의 모든 Shape을 Group → Export PNG → Ungroup
    ///    → 투명 배경으로 Shape만 추출됨 (단일 Shape이면 직접 Export)
    /// 2. Slide export 폴백: Shape-only가 실패하면 슬라이드 전체를 PNG로 (배경 포함)
    ///
    /// 캐시: %LocalAppData%\TeampptAddin\thumbnails\header_N.png
    /// 캐시 유효성은 TaskPaneHost.LoadThumbnail에서 pptx 수정일과 비교하여 판단.
    /// </summary>
    public static class ThumbnailGenerator
    {
        /// <summary>
        /// pptx 파일에서 Shape-only 썸네일 PNG를 생성.
        /// 이미 outputPngPath가 존재하면 스킵. 실패 시 slide-level export로 폴백.
        /// </summary>
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

        /// <summary>
        /// Shape-only export: 복수 Shape → Group → Export PNG → Ungroup.
        /// 단일 Shape이면 Group 없이 직접 Export.
        /// </summary>
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
