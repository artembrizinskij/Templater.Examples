using NGS.Templater;

namespace PptxResizeError
{
    internal class Program
    {
        static async Task Main()
        {
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");

            var factory = Configuration.Builder.NavigateSeparator(':', null).Include(new Handler().Handle).Build();
            var ext = "pptx";
            var file = "template." + ext;

            await using var fs = File.Open(path + $"template/{file}", FileMode.Open, FileAccess.ReadWrite);
            using var output = new MemoryStream();
            using (var doc = factory.Open(fs, ext, output))
            {
                var input = new Dictionary<string, object>[]
                {
                    new()
                    {
                        {
                            "Projet", new
                            {
                                Gpt = false,
                                SEP = false
                            }
                        },
                        {
                            "Slides", new
                            {
                                Travaux = new []
                                {
                                    new { Title="Slide1" },
                                    new { Title="Slide2" }
                                }
                            }
                        }
                    }
                };
                doc.Process(input);
            }

            await using var fileStream = new FileStream(path + "result." + ext, FileMode.OpenOrCreate, FileAccess.Write);
            output.Seek(0, SeekOrigin.Begin);
            await output.CopyToAsync(fileStream);
            fileStream.Close();
        }

        private class Handler
        {
            public Handled Handle(object value, string metadata, string tag, int position, ITemplater templater)
            {
                if (metadata == "collapse-If-false" && value is false)
                {
                    var result = templater.Resize(tag, 0);

                    return result ? Handled.NestedTags : Handled.Nothing;
                }

                return Handled.Nothing;
            }
        }
    }
}