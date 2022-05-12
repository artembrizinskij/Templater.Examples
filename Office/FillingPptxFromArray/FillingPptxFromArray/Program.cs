using NGS.Templater;
using System.Collections.Generic;
using System.IO;

namespace FillingPptxFromArray
{
    internal class Program
    {
        static void Main()
        {
            var input = new Dictionary<string, object>
            {
                { "text", "common text" },
                { "items", new object[]
                {
                    new { title = "title" }, 
                    new { title = "title2" }
                } },

                { "title", "common title" },
                { "items2", new object[]
                {
                    new { text = "text 1" },
                    new { text = "text 2" }
                } }
            };

            ProcessingResavedPptxWithMicrosoftPowerPoint(input);
            ProcessingOriginalPptx(input);
        }

        private static void ProcessingOriginalPptx(Dictionary<string, object> input)
        {
            var fileName = "template.pptx";
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");
            File.Copy(path + $"template/{fileName}", path + fileName, true);

            using var doc = Configuration.Builder
                .Build()
                .Open(path + fileName);

            doc.Process(input);
        }

        private static void ProcessingResavedPptxWithMicrosoftPowerPoint(Dictionary<string, object> input)
        {
            var fileName = "Re-saved_WithMicrosoftPowerPoint.pptx";
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");
            File.Copy(path + $"template/{fileName}", path + fileName, true);

            using var doc = Configuration.Builder
                .Build()
                .Open(path + fileName);

            doc.Process(input);
        }
    }
}
