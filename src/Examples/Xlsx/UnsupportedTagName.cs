namespace Examples.Xlsx;

public class UnsupportedTagName
{
    public static void Run()
    {
        Engine.Run("UnsupportedTag.xlsx", new Dictionary<string, object>()
        {
            { "Last_x0020_Name_x003a__x0020_Mid", new {Value = "some text"} }
        });
    }
}