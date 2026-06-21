using System.Collections.Generic;
using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 로컬 인제스트 코어: 묶음 pptx → 섹션 읽기 → 슬라이드 1장=에셋 1개 split + 768px 썸네일.
    /// LLM/Supabase 없음. 분할된 pptx/png는 outputDir에 저장.
    /// </summary>
    public static class IngestRunner
    {
        public static List<AssetSplitItem> Run(string bundlePptxPath, string outputDir)
        {
            var app = Globals.Application;
            if (app == null)
            {
                Logger.Log("IngestRunner: app is null");
                return new List<AssetSplitItem>();
            }
            if (!Directory.Exists(outputDir)) Directory.CreateDirectory(outputDir);

            PowerPoint.Presentation source = null;
            try
            {
                source = app.Presentations.Open(
                    bundlePptxPath,
                    MsoTriState.msoTrue,    // ReadOnly
                    MsoTriState.msoFalse,
                    MsoTriState.msoFalse);  // WithWindow=False

                var sections = SectionReader.Read(source);
                Logger.Log($"Ingest: {sections.Count} sections in {Path.GetFileName(bundlePptxPath)}");

                var plan = IngestPlanner.Plan(sections);

                foreach (var item in plan)
                {
                    var pptxPath = Path.Combine(outputDir, item.PptxFileName);
                    var pngPath = Path.Combine(outputDir, item.ThumbFileName);
                    SlideSplitter.SplitSlide(app, source, item.SourceSlideIndex, pptxPath);
                    SlideImageRenderer.Render(source, item.SourceSlideIndex, pngPath);
                    Logger.Log($"Ingest: {item.AssetId} (slide {item.SourceSlideIndex})");
                }

                source.Close();
                return plan;
            }
            finally
            {
                if (source != null) Marshal.ReleaseComObject(source);
            }
        }
    }
}
