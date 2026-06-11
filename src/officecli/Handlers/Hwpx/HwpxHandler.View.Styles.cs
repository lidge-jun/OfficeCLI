using System.Text;
using System.Text.Json.Nodes;
using System.Xml.Linq;
using OfficeCli.Core;

namespace OfficeCli.Handlers;

public partial class HwpxHandler
{
    // ==================== Styles ====================

    public string ViewAsStyles()
    {
        if (_doc.Header?.Root == null) return "(no header.xml)";
        var sb = new StringBuilder();
        var styles = _doc.Header.Root.Descendants(HwpxNs.Hh + "style").ToList();
        sb.AppendLine($"Styles: {styles.Count}");
        foreach (var style in styles)
        {
            var id = style.Attribute("id")?.Value ?? "?";
            var name = style.Attribute("name")?.Value ?? "(unnamed)";
            var engName = style.Attribute("engName")?.Value ?? "";
            var type = style.Attribute("type")?.Value ?? "PARA";
            var charPrId = style.Attribute("charPrIDRef")?.Value ?? "0";
            var paraPrId = style.Attribute("paraPrIDRef")?.Value ?? "0";
            var eng = !string.IsNullOrEmpty(engName) ? $" ({engName})" : "";
            sb.AppendLine($"  [{id}] {name}{eng} [{type}] charPr={charPrId} paraPr={paraPrId}");
        }
        return sb.ToString().TrimEnd();
    }
}
