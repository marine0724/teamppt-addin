using System.Collections.Generic;
using Newtonsoft.Json;

namespace TeampptAddin
{
    public class DesignConcept
    {
        [JsonProperty("id")] public string Id { get; set; }
        [JsonProperty("name")] public string Name { get; set; }
        [JsonProperty("styleTags")] public List<string> StyleTags { get; set; }
        [JsonProperty("colors")] public Dictionary<string, string> Colors { get; set; }
        [JsonProperty("fonts")] public Dictionary<string, string> Fonts { get; set; }
    }
}
