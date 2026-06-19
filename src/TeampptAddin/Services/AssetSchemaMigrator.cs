using Newtonsoft.Json.Linq;

namespace TeampptAddin
{
    public static class AssetSchemaMigrator
    {
        private static readonly string[] V1ColorRoles = { "main", "sub1", "sub2", "text" };

        public static JObject Migrate(JObject raw)
        {
            var obj = (JObject)raw.DeepClone();

            if (obj["kind"] == null)
                obj["kind"] = "component";

            if (obj["scope"] == null)
                obj["scope"] = "slide";

            var colors = obj["colors"];
            if (colors is JObject colorObj)
            {
                var arr = new JArray();
                foreach (var role in V1ColorRoles)
                {
                    var val = colorObj[role];
                    if (val == null) continue;
                    arr.Add(new JObject
                    {
                        ["role"] = role,
                        ["value"] = val,
                        ["locked"] = false
                    });
                }
                obj["colors"] = arr;
            }

            obj["schemaVersion"] = 2;
            return obj;
        }
    }
}
