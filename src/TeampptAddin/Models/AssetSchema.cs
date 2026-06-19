using System.Collections.Generic;
using Newtonsoft.Json;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public class AssetColor
    {
        [JsonProperty("role")] public string Role { get; set; }
        [JsonProperty("value")] public string Value { get; set; }
        [JsonProperty("locked")] public bool Locked { get; set; }
    }

    public class AssetFont
    {
        [JsonProperty("role")] public string Role { get; set; }
        [JsonProperty("family")] public string Family { get; set; }
        [JsonProperty("fallback")] public string Fallback { get; set; }
        [JsonProperty("weight")] public string Weight { get; set; }
        [JsonProperty("source")] public string Source { get; set; }
    }

    public class AssetSlot
    {
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("type")] public string Type { get; set; }
        [JsonProperty("perSlide")] public bool PerSlide { get; set; }
        [JsonExtensionData] public Dictionary<string, JToken> Extra { get; set; }
    }
}
