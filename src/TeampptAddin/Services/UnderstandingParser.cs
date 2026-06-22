using System.Collections.Generic;
using System.Linq;
using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public static class UnderstandingParser
    {
        public static AssetUnderstanding Parse(string llmJson, string category, string file)
        {
            var o = JObject.Parse(llmJson);

            var asset = new HeaderAsset
            {
                SchemaVersion = 2,
                File = file,
                Category = category,
                Scope = "slide",
                Name = o["name"]?.ToString() ?? "",
                Kind = o["kind"]?.ToString() ?? "component",
                UseWhen = o["use_when"]?.ToString() ?? "",
                ContentFit = StrList(o["content_fit"]),
                Tags = StrList(o["tags"]),
                Colors = (o["colors"] as JArray)?.Select(c => new AssetColor
                {
                    Role = c["role"]?.ToString(),
                    Value = c["value"]?.ToString(),
                    Locked = c["locked"]?.Value<bool>() ?? false
                }).ToList() ?? new List<AssetColor>(),
                Fonts = (o["fonts"] as JArray)?.Select(f => new AssetFont
                {
                    Role = f["role"]?.ToString(),
                    Family = f["family"]?.ToString(),
                    Weight = f["weight"]?.ToString()
                }).ToList() ?? new List<AssetFont>(),
                Slots = (o["slots"] as JArray)?.Select(s => new AssetSlot
                {
                    Name = s["name"]?.ToString(),
                    Type = s["type"]?.ToString(),
                    PerSlide = s["perSlide"]?.Value<bool>() ?? false
                }).ToList() ?? new List<AssetSlot>()
            };

            return new AssetUnderstanding
            {
                Asset = asset,
                ExampleIntents = StrList(o["example_intents"])
            };
        }

        private static List<string> StrList(JToken token)
        {
            return (token as JArray)?.Select(t => t.ToString()).ToList() ?? new List<string>();
        }
    }
}
