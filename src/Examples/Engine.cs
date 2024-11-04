using NGS.Templater;

namespace Examples;


public static class Engine
{
    private static readonly string BasePath = Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\..\\Templates\\");

    public static void Run(string fileName, Dictionary<string, object> data)
    {
        var destPath = BasePath + $"Results\\{fileName}";
        var builder = Configuration.Builder;

        File.Copy($"{BasePath}{fileName}", destPath, true);

        using var doc = builder
                        .Build()
                        .Open(destPath);

        doc.Process(data);
    }
}