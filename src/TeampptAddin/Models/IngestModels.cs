namespace TeampptAddin
{
    public enum IngestStage { Understanding, Embedding, Uploading, AssetDone }

    public class IngestProgress
    {
        public int Index { get; set; }
        public int Total { get; set; }
        public string AssetId { get; set; }
        public IngestStage Stage { get; set; }
        public string PngPath { get; set; }
        public string Kind { get; set; }
        public string Name { get; set; }
    }

    /// <summary>PowerPoint 섹션 1개 = 카테고리 1개. Interop SectionProperties에서 읽음.</summary>
    public class SectionInfo
    {
        public string Name { get; set; }           // 섹션명 = category (예: "표지")
        public int FirstSlideIndex { get; set; }    // 1-based, 이 섹션 첫 슬라이드
        public int SlideCount { get; set; }         // 이 섹션 슬라이드 수
    }

    /// <summary>슬라이드 1장 = 에셋 1개. 분할 계획의 한 항목.</summary>
    public class AssetSplitItem
    {
        public int SourceSlideIndex { get; set; }   // 1-based, 원본 묶음 pptx 안 위치
        public string Category { get; set; }        // 섹션명
        public string AssetId { get; set; }         // "표지_01"
        public string PptxFileName { get; set; }    // "표지_01.pptx"
        public string ThumbFileName { get; set; }   // "표지_01.png"
    }
}
