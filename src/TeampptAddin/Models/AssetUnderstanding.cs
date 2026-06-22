using System.Collections.Generic;

namespace TeampptAddin
{
    public class AssetUnderstanding
    {
        public HeaderAsset Asset { get; set; }
        public List<string> ExampleIntents { get; set; } = new List<string>();
    }
}
