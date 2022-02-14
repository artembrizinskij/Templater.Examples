using NGS.Templater;
using System.Collections.Generic;
using System.IO;
using System.Xml.Linq;

namespace MergeXml
{
    internal class Program
    {
        static void Main()
        {
            var fileName = "example.docx";
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");
            File.Copy(path + $"template/{fileName}", path + fileName, true);

            var input = new Dictionary<string, object>
            {
                { "prop1", "value1_with_merge-xml" },
                { "prop2", "value2_with_merge-xml" },
                { "prop3", "value3" },
                { "prop4", "value4" },
            };

            using var doc = Configuration.Builder.Include(ToXml).Build().Open(path + fileName);
            doc.Process(input);
        }

        private static object ToXml(object value, string metadata)
        {
            if (metadata == "xml")
            {
                return XElement.Parse(@"
                                        <w:p xmlns:w=""http://schemas.openxmlformats.org/wordprocessingml/2006/main"">
		                                        <w:r>
			                                        <w:t>"+value+@"</w:t>
		                                        </w:r>
                                        </w:p>");
            }

            return value;
        }
    }
}
