using System.Runtime.InteropServices;
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
                    Marshal.ReleaseComObject(srcPres);
                }
            }
        }

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
