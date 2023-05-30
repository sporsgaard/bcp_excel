// Static class ParseFormatOptions with a static method Parse that takes a string and returns a list of tupples: (string, integers).
// input string looks like
// "i,s20,u30,s,u"
// and should return a list of tupples:
// ("int", -1), ("string", 20), ("unicode", 30), ("string", -1), ("unicode", -1)

namespace AlarmPeople.Bcp;

public static class ParseFormatOptions
{
    public static List<(string, int)> Parse(string input)
    {
        var result = new List<(string, int)>();
        if (string.IsNullOrEmpty(input))
            return result;
        var parts = input.Split(',');
        foreach (var part in parts)
        {
            var (type, value) = ParsePart(part);
            result.Add((type, value));
        }
        return result;
    }

    private static (string, int) ParsePart(string part)
    {
        var type = part[0] switch
        {
            'i' => "int",
            's' => "string",
            'u' => "unicode",
            _ => throw new ArgumentException($"Invalid type: {part[0]}")
        };
        var value = part.Length > 1 ? int.Parse(part[1..]) : -1;
        return (type, value);
    }
}
