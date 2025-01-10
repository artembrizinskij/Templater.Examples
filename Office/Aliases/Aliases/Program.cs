using NGS.Templater;
using System.Collections;

namespace Aliases;

public class Program
{
    private static readonly string BasePath = Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\..\\Templates\\");

    public static object TopExpression(object parent, object value, string member, string metadata)
    {
        var col = value as ICollection;

        if (!metadata.StartsWith("top(") || col == null || col.Count < 2) return value;

        var property = metadata.Substring(4, metadata.Length - 5);

        if (!int.TryParse(property, out var count))
        {
            return value;
        }

        return col.Cast<object>().Take(count);
    }

    public static void Main(string[] args)
    {
        var destPath = BasePath + "Results\\aliases.docx";
        File.Copy($"{BasePath}aliases.docx", destPath, true);

        var factory = Configuration.Builder
                        .NavigateSeparator('|', null)
                        .Include(TopExpression)
                        .Build();

        var data = new Dictionary<string, object>()
        {
            {
                "RiMgt1", new[]
                {
                    new { RiskTitle = "RiMgt1-1" },
                    new { RiskTitle = "RiMgt1-2" },
                    new { RiskTitle = "RiMgt1-3" },
                    new { RiskTitle = "RiMgt1-4" },
                    new { RiskTitle = "RiMgt1-5" },
                }
            },
            {
                "RiMgt1NO", new[]
                {
                    new { RiskTitle = "item1" }
                }
            }
        };

        using var doc = factory.Open(destPath);
        doc.Process(data);
    }
}