using System.Collections.Generic;
using NGS.Templater;
using System.IO;
namespace FillInCellsByLink
{
    internal class Program
    {
        static void Main(string[] args)
        {
            var fileName = "template.xlsx";
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");
            File.Copy(path + $"template/{fileName}", path + fileName, true);

            var input = new Dictionary<string, object>
            {
                {"collection1", new [] {
                        new
                        {
                            title = "title1",
                            nest_collection = new []
                            {
                                new {title = "nest title1 from obj1"},
                                new {title = "nest title2 from obj1"},
                                new {title = "nest title3 from obj1"}
                            }
                        },
                        new
                        {
                            title = "title2",
                            nest_collection = new []
                            {
                                new {title = "nest title from obj2"}
                            }
                        },
                        new
                        {
                            title = "title3",
                            nest_collection = new []
                            {
                                new {title = "nest title from obj3"}
                            }
                        }
                    }
                },
                {"collection2", new [] {
                        new
                        {
                            title = "title1"
                        },
                        new
                        {
                            title = "title2"
                        }
                    }
                }
            };
			var builder = Configuration.Builder;

            using var doc = builder.Build().Open(path + fileName);
            doc.Process(input);
        }
    }
}
