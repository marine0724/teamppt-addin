using System.Drawing;

namespace TeampptAddin
{
    /// <summary>
    /// 헤더 에셋 데이터 모델.
    /// Assets 폴더의 header_N.pptx 파일 하나를 나타냄.
    /// System.Drawing.Image 기반 (WPF BitmapImage 의존성 없음).
    /// </summary>
    public class HeaderAsset
    {
        public int Index { get; set; }
        public string Name { get; set; }
        public string PptxPath { get; set; }
        public string ThumbnailPath { get; set; }
        public Image Thumbnail { get; set; }
    }
}
