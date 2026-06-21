using System.IO;
using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>소스 묶음 pptx에서 슬라이드 1장을 복사해 단일 슬라이드 pptx로 저장.</summary>
    public static class SlideSplitter
    {
        public static void SplitSlide(
            PowerPoint.Application app,
            PowerPoint.Presentation source,
            int slideIndex,
            string outputPptxPath)
        {
            var dir = Path.GetDirectoryName(outputPptxPath);
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            PowerPoint.Presentation dest = null;
            try
            {
                // 화면에 띄우지 않고 새 프레젠테이션 생성
                dest = app.Presentations.Add(MsoTriState.msoFalse);

                source.Slides[slideIndex].Copy();
                dest.Slides.Paste();  // 인덱스 1에 붙음

                // 빈 기본 슬라이드가 있으면 제거 (Add 직후 기본 슬라이드 없음 — Paste만 존재)
                dest.SaveAs(
                    outputPptxPath,
                    PowerPoint.PpSaveAsFileType.ppSaveAsOpenXMLPresentation,
                    MsoTriState.msoFalse);
                dest.Close();
            }
            finally
            {
                if (dest != null) Marshal.ReleaseComObject(dest);
            }
        }
    }
}
