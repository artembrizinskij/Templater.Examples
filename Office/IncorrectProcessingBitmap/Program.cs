using NGS.Templater;
using System.Drawing;
using System.Drawing.Imaging;

var fileName = "template.docx";
var root = Directory.GetCurrentDirectory();
var path = Path.Combine(root + "\\..\\..\\..\\");
File.Copy(path + $"template/{fileName}", path + fileName, true);

var input = new Dictionary<string, object>
{
    {"value", "test" }
};
var builder = Configuration.Builder
        .Include(StoppedWorkingWithNewRelease)
        .Include(WorkaroundForNewRelease)
    ;

using var doc = builder.Build().Open(path + fileName);
doc.Process(input);

static object StoppedWorkingWithNewRelease(object value, string meta)
{
    if (meta == "format1")
    {
        var bitmap = (Bitmap)Image.FromFile(Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\..\\" + "template/img.bmp"));
        return bitmap;
    }

    return value;
}

static object WorkaroundForNewRelease(object value, string meta)
{
    if (meta == "format2")
    {
        var bitmap = (Bitmap)Image.FromFile(Path.Combine(Directory.GetCurrentDirectory() + "\\..\\..\\..\\" + "template/img.bmp"));

        var ms = new MemoryStream();
        bitmap.Save(ms, ImageFormat.Bmp);
        ms.Seek(0, SeekOrigin.Begin);

        return new ImageInfo(ms, ImageFormat.Jpeg.ToString(), bitmap.Width, bitmap.HorizontalResolution, bitmap.Height,
            bitmap.VerticalResolution);
    }

    return value;
}