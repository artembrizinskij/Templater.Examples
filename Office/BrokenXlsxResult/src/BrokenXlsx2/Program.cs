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

            var factory = Configuration.Builder.Build();

            await using var fs = File.Open(path + "template/template.xlsx", FileMode.Open, FileAccess.ReadWrite);
            using var output = new MemoryStream();
            using (var doc = factory.Open(fs, "xlsx", output))
            {
                doc.Process(string.Empty);
            }

            await using var fileStream = new FileStream(path + "result.xlsx", FileMode.OpenOrCreate, FileAccess.Write);
            output.Seek(0, SeekOrigin.Begin);
            await output.CopyToAsync(fileStream);
            fileStream.Close();
        }
    }
}


