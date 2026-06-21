using System.Collections.Generic;
using PowerPoint = Microsoft.Office.Interop.PowerPoint;

namespace TeampptAddin
{
    /// <summary>
    /// 열린 Presentation의 PowerPoint 섹션을 읽어 SectionInfo 목록으로 변환.
    /// SectionProperties는 PowerPoint 2013+ 제공. 1-based 인덱스.
    /// </summary>
    public static class SectionReader
    {
        public static List<SectionInfo> Read(PowerPoint.Presentation pres)
        {
            var result = new List<SectionInfo>();
            var props = pres.SectionProperties;
            int count = props.Count;
            for (int i = 1; i <= count; i++)
            {
                result.Add(new SectionInfo
                {
                    Name = props.Name(i),
                    FirstSlideIndex = props.FirstSlide(i),
                    SlideCount = props.SlidesCount(i),
                });
            }
            return result;
        }
    }
}
