using System.Collections.Generic;
using Newtonsoft.Json;

namespace TeampptAddin
{
    public class PaletteColors
    {
        [JsonProperty("background")]
        public string Background { get; set; }

        [JsonProperty("main")]
        public string Main { get; set; }

        [JsonProperty("sub1")]
        public string Sub1 { get; set; }

        [JsonProperty("sub2")]
        public string Sub2 { get; set; }

        [JsonProperty("text")]
        public string Text { get; set; }
    }

    public class StylePalette
    {
        [JsonProperty("id")]
        public string Id { get; set; }

        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("colors")]
        public PaletteColors Colors { get; set; }

        [JsonProperty("mood")]
        public List<string> Mood { get; set; }

        [JsonProperty("use_when")]
        public string UseWhen { get; set; }
    }

    public class StyleFont
    {
        [JsonProperty("name")]
        public string Name { get; set; }

        [JsonProperty("fallback")]
        public string Fallback { get; set; }

        [JsonProperty("mood")]
        public List<string> Mood { get; set; }

        [JsonProperty("use_when")]
        public string UseWhen { get; set; }
    }

    public class StyleConfig
    {
        [JsonProperty("palettes")]
        public List<StylePalette> Palettes { get; set; }

        [JsonProperty("fonts")]
        public List<StyleFont> Fonts { get; set; }
    }
}
