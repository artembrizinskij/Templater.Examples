using System;
using System.Collections.Generic;
using System.Drawing;
using System.IO;
using System.Linq;
using System.Text.Json;
using System.Threading.Tasks;
using System.Xml.Linq;
using NGS.Templater;

namespace EmbedImageError
{
    internal class Program
    {
        static async Task Main()
        {
            var fileName = "example.docx";
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");
            File.Copy(path + $"template/{fileName}", path + fileName, true);

            var input = await ReadAsync<Dictionary<string, string>>($"{path}template/data.json");

            using var doc = Configuration.Builder
                .Include(ToImage)
                .Build()
                .Open(path + fileName);

            doc.Process(input);
        }

        public static async Task<T> ReadAsync<T>(string filePath)
        {
            await using FileStream stream = File.OpenRead(filePath);
            return await JsonSerializer.DeserializeAsync<T>(stream);
        }

        private static object ToImage(object value, string metadata)
        {
            //picture(200,300)
            if (metadata.StartsWith("picture("))
            {
                var str = value?.ToString();

                if (string.IsNullOrEmpty(str))
                {
                    return value;
                }

                //(200,300)
                var args = metadata.Split('(', ')')[1].Split(',').Select(int.Parse).ToArray();
                var width = args[0];
                var height = args[1];

                var imageBytes = Convert.FromBase64String(str);
                
                var imageStream = new MemoryStream(imageBytes);
                using var img = Image.FromStream(imageStream);

                var (hres, vres) = GetImageResoulutions(img);

                return new ImageInfo(imageStream, "jpg", width, hres, height, vres);
            }

            return value;
        }

        private static (int MaxWidth, int MaxHeight) GetImageWidthAndHeight(Image image) => (image.Width, image.Height);

        private static (float HRes, float VRes) GetImageResoulutions(Image image) => (image.HorizontalResolution, image.VerticalResolution);
    }   
}
