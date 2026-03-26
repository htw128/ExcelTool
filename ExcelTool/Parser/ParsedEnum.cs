using System.Collections.Generic;

namespace ExcelTool.Parser
{
    public class ParsedEnumMember
    {
        /// <summary>成员名，如 "Home"</summary>
        public string Key { get; set; }

        /// <summary>显式值，为 null 时代码生成时自动递增（不写值）</summary>
        public int? Value { get; set; }

        /// <summary>行内注释，如 "主界面"</summary>
        public string Desc { get; set; }
    }

    public class ParsedEnum
    {
        /// <summary>A 列 Id，供外部系统通过 EnumIds 常量类引用</summary>
        public uint Id { get; set; }

        /// <summary>枚举类型名，如 "GameState"</summary>
        public string EnumName { get; set; }

        public List<ParsedEnumMember> Members { get; set; } = new();
    }
}
