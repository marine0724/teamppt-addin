using System.Runtime.InteropServices;
using Microsoft.Office.Core;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// PowerPoint COM API를 사용한 Shape 삽입 엔진.
    /// 헤더 에셋 pptx 파일의 첫 번째 슬라이드에서 모든 Shape을 복사하여
    /// 현재 활성 슬라이드에 붙여넣는다.
    ///
    /// 동작 흐름:
    /// 1. CopyShapesToClipboard: pptx를 ReadOnly로 열고, 슬라이드1의 전체 Shape을 클립보드에 복사
    /// 2. InsertToActiveSlide: Copy → 활성 슬라이드에 Paste (클릭 삽입용)
    ///
    /// 드래그 삽입 시에는 DragHandler가 CopyShapesToClipboard만 호출하고,
    /// EndDrag에서 직접 slide.Shapes.Paste() + CoordinateConverter로 위치 지정.
    /// </summary>
    public static class ShapeInserter
    {
        /// <summary>
        /// pptx 파일을 열어 첫 슬라이드의 모든 Shape을 클립보드에 복사.
        /// pptx는 ReadOnly + 창 없이(msoFalse) 열리며, 복사 후 즉시 닫힘.
        /// </summary>
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
                    Marshal.ReleaseComObject(srcPres);
                }
            }
        }

        /// <summary>
        /// 클릭 삽입용: pptx의 Shape을 복사 → 활성 슬라이드 중앙에 Paste.
        /// 위치 지정 없이 기본 위치에 붙여넣기됨.
        /// </summary>
        public static void InsertToActiveSlide(string headerPptxPath)
        {
            var app = Globals.Application;
            if (app == null) return;

            var window = app.ActiveWindow;
            if (window == null) return;

            CopyShapesToClipboard(headerPptxPath);

            var slide = (PowerPoint.Slide)window.View.Slide;
            slide.Shapes.Paste();
        }
    }
}
