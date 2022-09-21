using DocumentFormat.OpenXml.Packaging;
using DocumentFormat.OpenXml.Spreadsheet;
using NGS.Templater;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Threading.Tasks;

namespace BrokenXlsx
{
    internal class Program
    {
        static async Task Main()
        {
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");

            var input = new Dictionary<string, string>()
            {
                { "prop1", "val1" },
                { "prop2", "val2" },
            };

            var factory = Configuration.Builder.Build();

            using var ms = new MemoryStream();
            await using (var fs = File.Open(path + "template/template.xlsx", FileMode.Open, FileAccess.ReadWrite))
            {
                await fs.CopyToAsync(ms);
            }

            PrepareDoc(ms);

            using (var output = new MemoryStream())
            {
                using (var doc = factory.Open(ms, "xlsx", output))
                {
                    doc.Process(input);
                }

                await using var fileStream = new FileStream(path + "result.xlsx", FileMode.OpenOrCreate, FileAccess.Write);
                output.Seek(0, SeekOrigin.Begin);
                await output.CopyToAsync(fileStream);
                fileStream.Close();
            }
        }

        private static void PrepareDoc(Stream input)
        {
            using var excelDoc = SpreadsheetDocument.Open(input, true);
            var wbPart = excelDoc.WorkbookPart;

            var stringTable = wbPart
                .GetPartsOfType<SharedStringTablePart>()
                .FirstOrDefault();

            if (stringTable != null)
            {
                var cell = wbPart.WorksheetParts.FirstOrDefault()
                    .Worksheet.Descendants<Row>().FirstOrDefault()
                    .Elements<Cell>()
                    .FirstOrDefault();

                var value = stringTable.SharedStringTable.ElementAt(3).InnerText;
                if (value == "remove me")
                {
                    cell.Remove();
                }
            }
        }
    }
}
