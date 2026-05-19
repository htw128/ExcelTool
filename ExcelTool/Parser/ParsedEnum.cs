using System.Collections.Generic;
using System.Text.RegularExpressions;
using System.Linq;

namespace ExcelTool.Parser
{
    public class ParsedEnumMember
    {
        public string Key
        {
            get;
            set { field = ParsedEnum.NormalizeIdentifier(value); }
        }

        /// <summary>显式值，为 null 时代码生成时自动递增（不写值）</summary>
        public int? Value { get; set; }

        /// <summary>行内注释，如 "主界面"</summary>
        public string Desc { get; set; }
    }

    public class ParsedEnum
    {
        /// <summary>A 列 Id，供外部系统通过 EnumIds 常量类引用</summary>
        public uint Id { get; set; }

        public string EnumName
        {
            get;
            set
            {
                field = NormalizeIdentifier(value);
            }
        }

        public static string NormalizeIdentifier(string input)
        {
            if (string.IsNullOrWhiteSpace(input))
            {
                return string.Empty;
            }

            string[] parts = Regex.Split(input, @"[\s_-]+")
                .Where(s => !string.IsNullOrWhiteSpace(s))
                .ToArray();

            for (int i = 0; i < parts.Length; i++)
            {
                string part = parts[i];

                if (part.Length == 1)
                {
                    parts[i] = char.ToUpperInvariant(part[0]).ToString();
                }
                else
                {
                    parts[i] = char.ToUpperInvariant(part[0]) + part.Substring(1);
                }
            }

            return string.Concat(parts);
        }

        public List<ParsedEnumMember> Members { get; set; } = [];
    }
}
