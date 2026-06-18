using System.Drawing;

namespace TeampptAddin
{
    public class HeaderAsset
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string PptxPath { get; set; }
        public string ThumbnailPath { get; set; }
        public Image Thumbnail { get; set; }
    }
}
