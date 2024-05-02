using NGS.Templater;

var fileName = "template.docx";
var root = Directory.GetCurrentDirectory();
var path = Path.Combine(root + "\\..\\..\\..\\");
File.Copy(path + $"template/{fileName}", path + fileName, true);

var input = new Dictionary<string, object>
            {
                {"name", "Giles\nGregg" }
            };
var builder = Configuration.Builder;

using var doc = builder.Build().Open(path + fileName);
doc.Process(input);