using Newtonsoft.Json;
using Newtonsoft.Json.Linq;
using NGS.Templater;
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;

namespace EmbedImageError
{
    internal class Program
    {
        static async Task Main()
        {
            var fileName = "Template.docx";
            var root = Directory.GetCurrentDirectory();
            var path = Path.Combine(root + "\\..\\..\\..\\");
            File.Copy(path + $"template/{fileName}", path + fileName, true);
            var input = await ReadStrAsync<string>($"{path}template/data.json");

            using var json = new StringReader(input);
            var jsonData = Deserialize<IDictionary<string, object>>(json) ?? new Dictionary<string, object>();

            using var doc = Configuration.Builder
                .Include(new HideBlockProcessor().Handle)
                .Build()
                .Open(path + fileName);

            doc.Process(jsonData);
        }

        internal class HideBlockProcessor
        {
            private const string Name = "hide-block-if-empty";

            public Handled Handle(object value, string metadata, string property, int position, ITemplater templater)
            {
                if (!metadata.StartsWith(Name))
                {
                    return Handled.Nothing;
                }

                var isEmptyString = string.IsNullOrEmpty(value?.ToString());

                if (!isEmptyString)
                {
                    return Handled.Nothing;
                }

                var result = templater.Resize(property, 0);
                return result ? Handled.NestedTags : Handled.Nothing;

            }
        }

        private static T Deserialize<T>(StringReader reader)
        {
            var serializer = new Newtonsoft.Json.JsonSerializer();
            serializer.Converters.Add(new DictionaryConverter());

            using var jsonReader = new JsonTextReader(reader);
            return serializer.Deserialize<T>(jsonReader);
        }

        private static async Task<string> ReadStrAsync<T>(string filePath)
        {
            return await File. ReadAllTextAsync(filePath);
        }
    }   
}

public class DictionaryConverter : JsonConverter
{
    public override void WriteJson(JsonWriter writer, object value, JsonSerializer serializer)
    {
        WriteValue(writer, value);
    }

    private void WriteValue(JsonWriter writer, object value)
    {
        var t = JToken.FromObject(value);
        switch (t.Type)
        {
            case JTokenType.Object:
                WriteObject(writer, value);
                break;

            case JTokenType.Array:
                WriteArray(writer, value);
                break;

            default:
                writer.WriteValue(value);
                break;
        }
    }

    private void WriteObject(JsonWriter writer, object value)
    {
        writer.WriteStartObject();
        var obj = value as IDictionary<string, object>;
        foreach (var kvp in obj)
        {
            writer.WritePropertyName(kvp.Key);
            WriteValue(writer, kvp.Value);
        }

        writer.WriteEndObject();
    }

    private void WriteArray(JsonWriter writer, object value)
    {
        writer.WriteStartArray();
        var array = value as IEnumerable<object>;
        foreach (var o in array)
        {
            WriteValue(writer, o);
        }

        writer.WriteEndArray();
    }

    public override object ReadJson(JsonReader reader, Type objectType, object existingValue, JsonSerializer serializer)
    {
        return ReadValue(reader);
    }

    protected object ReadValue(JsonReader reader)
    {
        switch (reader.TokenType)
        {
            case JsonToken.StartObject:
                return ReadObject(reader);
            case JsonToken.StartArray:
                return ReadArray(reader);
            default:
                return reader.Value;
        }
    }

    private object ReadArray(JsonReader reader)
    {
        var list = new List<object>();

        while (reader.Read())
        {
            switch (reader.TokenType)
            {
                case JsonToken.Comment:
                    break;
                case JsonToken.EndArray:
                    return list;

                default:
                    var v = ReadValue(reader);
                    list.Add(v);
                    break;
            }
        }

        throw new JsonSerializationException("Unexpected end when reading IDictionary<string, object>");
    }

    protected virtual string PreparePropertyName(object value)
    {
        return value.ToString();
    }

    protected virtual object ReadObject(JsonReader reader)
    {
        var obj = new Dictionary<string, object>();

        while (reader.Read())
        {
            switch (reader.TokenType)
            {
                case JsonToken.PropertyName:
                    var propertyName = PreparePropertyName(reader.Value);

                    if (!reader.Read())
                    {
                        throw new JsonSerializationException(
                            "Unexpected end when reading IDictionary<string, object>");
                    }

                    var v = ReadValue(reader);

                    obj[propertyName] = v;
                    break;
                case JsonToken.Comment:
                    break;
                case JsonToken.EndObject:
                    return obj;
            }
        }

        throw new JsonSerializationException("Unexpected end when reading");
    }

    public override bool CanConvert(Type objectType)
    {
        return typeof(IDictionary<string, object>).IsAssignableFrom(objectType);
    }
}