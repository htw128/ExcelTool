using System.Globalization;

namespace ExcelTool;

public static class StringExtensions
{
    /// <summary>
    /// 将 PascalCase 字符串转换为 camelCase
    /// </summary>
    public static string ToCamelCase(this string input)
    {
        if (string.IsNullOrEmpty(input))
        {
            return input;
        }

        // 如果第一个字符已经是小写，或者不是字母，直接返回（避免不必要的操作）
        if (!char.IsUpper(input[0]))
        {
            return input;
        }

        // 特殊情况：如果字符串只有一个字符且是大写，直接转小写
        if (input.Length == 1)
        {
            return input;
        }

        // 将第一个字符转小写，拼接剩余部分
        return char.ToLower(input[0], CultureInfo.InvariantCulture) + input.Substring(1);
    }
}