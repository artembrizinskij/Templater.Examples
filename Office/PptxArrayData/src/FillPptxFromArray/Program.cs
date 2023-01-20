using System.Collections.Generic;
using NGS.Templater;
using System.IO;
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
            var ext = "pptx";
            var file = "template."+ext;

            await using var fs = File.Open(path + $"template/{file}", FileMode.Open, FileAccess.ReadWrite);
            using var output = new MemoryStream();
            using (var doc = factory.Open(fs, ext, output))
            {
                var input = new Dictionary<string, object>[]
                {
                    new()
                    {
                        {
                            "AssessmentDomains", new [] {
                                new
                                {
                                    Domain = new {Name = "test1"},
                                    Score = 1
                                },
                                new
                                {
                                    Domain = new {Name = "test2"},
                                    Score = 2
                                }
                            }
                        }
                    }
                };
                doc.Process(input);
            }

            await using var fileStream = new FileStream(path + "result."+ext, FileMode.OpenOrCreate, FileAccess.Write);
            output.Seek(0, SeekOrigin.Begin);
            await output.CopyToAsync(fileStream);
            fileStream.Close();
        }
    }
}


