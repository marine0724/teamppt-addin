using System.Collections.Generic;
using System.Linq;

namespace TeampptAddin
{
    public static class CatalogBuilder
    {
        public static List<CatalogEntry> Build(IEnumerable<HeaderAsset> assets)
        {
            return assets.Select(a => new CatalogEntry
            {
                File = a.File,
                Name = a.Name,
                Kind = a.Kind,
                Category = a.Category,
                Scope = a.Scope,
                Tags = a.Tags ?? new List<string>(),
                UseWhen = a.UseWhen,
                SlotNames = (a.Slots ?? new List<AssetSlot>()).Select(s => s.Name).ToList(),
                ColorRoles = (a.Colors ?? new List<AssetColor>()).Select(c => c.Role).ToList(),
                FontRoles = (a.Fonts ?? new List<AssetFont>()).Select(f => f.Role).ToList()
            }).ToList();
        }
    }
}
