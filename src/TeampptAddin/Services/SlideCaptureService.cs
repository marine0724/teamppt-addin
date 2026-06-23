using System;
using System.IO;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    public class SlideCapture
    {
        public string PngPath { get; set; }
        public int SlideNumber { get; set; }
    }

    public static class SlideCaptureService
    {
        public static SlideCapture CaptureCurrentSlide()
        {
            var app = Globals.Application;
            var win = app?.ActiveWindow;
            if (win == null) return null;

            PowerPoint.Slide slide;
            try { slide = (PowerPoint.Slide)win.View.Slide; }
            catch { return null; }
            if (slide == null) return null;

            var pres = win.Presentation;
            int index = slide.SlideIndex;

            var dir = Path.Combine(
                Environment.GetFolderPath(Environment.SpecialFolder.LocalApplicationData),
                "TeampptAddin", "cache", "screen-share");
            if (!Directory.Exists(dir)) Directory.CreateDirectory(dir);

            var png = Path.Combine(dir, $"slide-{index}.png");
            SlideImageRenderer.Render(pres, index, png);

            return new SlideCapture { PngPath = png, SlideNumber = index };
        }
    }
}
