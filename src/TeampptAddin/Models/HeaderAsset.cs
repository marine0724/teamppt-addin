using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public class HeaderAsset
    {
        [JsonProperty("schemaVersion")] public int SchemaVersion { get; set; } = 1;
        [JsonProperty("file")] public string File { get; set; }
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("kind")] public string Kind { get; set; } = "component";
        [JsonProperty("category")] public string Category { get; set; }
        [JsonProperty("scope")] public string Scope { get; set; } = "slide";
        [JsonProperty("content_fit")] public List<string> ContentFit { get; set; }
        [JsonProperty("use_when")] public string UseWhen { get; set; }
        [JsonProperty("provenance")] public string Provenance { get; set; }
        [JsonProperty("grid_columns")] public int GridColumns { get; set; } = 1;
        [JsonProperty("tags")] public List<string> Tags { get; set; }
        [JsonProperty("colors")] public List<AssetColor> Colors { get; set; }
        [JsonProperty("fonts")] public List<AssetFont> Fonts { get; set; }
        [JsonProperty("slots")] public List<AssetSlot> Slots { get; set; }
        [JsonExtensionData] public Dictionary<string, JToken> Extra { get; set; }
    }
}
