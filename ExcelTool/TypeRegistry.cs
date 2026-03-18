using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelTool
{
    /// <summary>
    /// 类型描述符：集中管理单个类型的「二进制写入」和「C# 代码生成」逻辑。
    /// 新增类型时只需在 TypeRegistry.Register() 里追加一条注册即可。
    /// </summary>
    public class TypeDescriptor
    {
        /// <summary>小写类型名，作为注册 key</summary>
        public string TypeName { get; init; }

        /// <summary>对应的 C# 属性类型字符串，例如 "int"、"List&lt;float&gt;"</summary>
        public string CSharpType { get; init; }

        /// <summary>将字符串值写入 BinaryWriter 的逻辑</summary>
        public Action<BinaryWriter, string> WriteBinary { get; init; }

        /// <summary>
        /// 生成 DeSerialize 方法体片段（变量名 → 代码行）
        /// 返回的字符串已含换行，调用方直接 sb.Append 即可
        /// </summary>
        public Func<string, string> GenDeserialize { get; init; }

        /// <summary>
        /// 生成 Serialize 方法体片段（变量名 → 代码行）
        /// </summary>
        public Func<string, string> GenSerialize { get; init; }
    }

    /// <summary>
    /// 类型注册表。所有支持的字段类型统一在这里注册，一处修改全局生效。
    /// </summary>
    public static class TypeRegistry
    {
        private static readonly Dictionary<string, TypeDescriptor> s_map = new();

        static TypeRegistry()
        {
            RegisterAll();
        }

        // ──────────────────────────────────────────────
        //  对外查询
        // ──────────────────────────────────────────────

        /// <summary>查找类型描述符，找不到返回 null</summary>
        public static TypeDescriptor Get(string typeName) =>
            s_map.TryGetValue(typeName.ToLower(), out var desc) ? desc : null;

        /// <summary>是否包含该类型</summary>
        public static bool Contains(string typeName) => s_map.ContainsKey(typeName.ToLower());

        // ──────────────────────────────────────────────
        //  注册入口：新增类型只需在此追加 Register(...)
        // ──────────────────────────────────────────────

        private static void RegisterAll()
        {
            // ---------- 基础值类型 ----------
            RegisterPrimitive("int",    "int",    "reader.ReadInt32()",    (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? 0              : Convert.ToInt32(v)))  ;
            RegisterPrimitive("uint",   "uint",   "reader.ReadUInt32()",   (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? 0u             : Convert.ToUInt32(v))) ;
            RegisterPrimitive("short",  "short",  "reader.ReadInt16()",    (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? (short)0       : Convert.ToInt16(v)))  ;
            RegisterPrimitive("ushort", "ushort", "reader.ReadUInt16()",   (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? (ushort)0      : Convert.ToUInt16(v))) ;
            RegisterPrimitive("sbyte",  "sbyte",  "reader.ReadSByte()",    (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? (sbyte)0       : Convert.ToSByte(v)))  ;
            RegisterPrimitive("byte",   "byte",   "reader.ReadByte()",     (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? (byte)0        : Convert.ToByte(v)))   ;
            RegisterPrimitive("float",  "float",  "reader.ReadSingle()",   (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? 0f             : Convert.ToSingle(v))) ;
            RegisterPrimitive("double", "double", "reader.ReadDouble()",   (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? 0d             : Convert.ToDouble(v))) ;
            RegisterPrimitive("long",   "long",   "reader.ReadInt64()",    (bw, v) => bw.Write(string.IsNullOrEmpty(v) ? 0L             : Convert.ToInt64(v)))  ;
            RegisterPrimitive("string", "string", "reader.ReadString()",   (bw, v) => bw.Write(v ?? ""))                                                        ;

            // bool 单独处理（解析逻辑稍特殊）
            Register(new TypeDescriptor
            {
                TypeName   = "bool",
                CSharpType = "bool",
                WriteBinary = (bw, v) =>
                {
                    if (string.IsNullOrEmpty(v)) { bw.Write(false); return; }
                    var s = v.Trim().ToLower();
                    bw.Write(s is "true" or "1");
                },
                GenDeserialize = name => $"\t\t{name} = reader.ReadBoolean();\n",
                GenSerialize   = name => $"\t\twriter.Write({name});\n",
            });

            // ---------- 集合类型 ----------
            // vector → List<float>（固定 3 分量，写入时先写分量数）
            Register(new TypeDescriptor
            {
                TypeName   = "vector",
                CSharpType = "List<float>",
                WriteBinary = (bw, v) =>
                {
                    if (string.IsNullOrEmpty(v)) { bw.Write(0); return; }
                    var parts = v.Replace("]", "").Replace("[", "").Split(',');
                    int count = 3;
                    bw.Write(count);
                    for (int i = 0; i < count; i++)
                        bw.Write(Convert.ToSingle(parts[i]));
                },
                GenDeserialize = name => GenListDeserialize(name, "float", "reader.ReadSingle()"),
                GenSerialize   = name => GenListSerialize(name),
            });

            // vectorlist → List<List<float>>
            Register(new TypeDescriptor
            {
                TypeName   = "vectorlist",
                CSharpType = "List<List<float>>",
                WriteBinary = (bw, v) =>
                {
                    if (string.IsNullOrEmpty(v)) { bw.Write(0); return; }
                    var parts = v.Replace("]", "").Replace("[", "").Split(',');
                    bw.Write(parts.Length / 3);
                    for (int i = 0; i < parts.Length; i++)
                    {
                        if (i % 3 == 0) bw.Write(3);
                        bw.Write(Convert.ToSingle(parts[i]));
                    }
                },
                GenDeserialize = name => GenVectorListDeserialize(name),
                GenSerialize   = name => GenVectorListSerialize(name),
            });

            // 简写集合类型（xxxlist）
            RegisterList("intlist",    "int",    "reader.ReadInt32()",  (bw, s) => bw.Write(Convert.ToInt32(s)));
            RegisterList("uintlist",   "uint",   "reader.ReadUInt32()", (bw, s) => bw.Write(Convert.ToUInt32(s)));
            RegisterList("shortlist",  "short",  "reader.ReadInt16()",  (bw, s) => bw.Write(Convert.ToInt16(s)));
            RegisterList("ushortlist", "ushort", "reader.ReadUInt16()", (bw, s) => bw.Write(Convert.ToUInt16(s)));
            RegisterList("sbytelist",  "sbyte",  "reader.ReadSByte()",  (bw, s) => bw.Write(Convert.ToSByte(s)));
            RegisterList("bytelist",   "byte",   "reader.ReadByte()",   (bw, s) => bw.Write(Convert.ToByte(s)));
            RegisterList("boollist",   "bool",   "reader.ReadBoolean()",(bw, s) => bw.Write(Convert.ToBoolean(s)));
            RegisterList("floatlist",  "float",  "reader.ReadSingle()", (bw, s) => bw.Write(Convert.ToSingle(s)));
            RegisterList("doublelist", "double", "reader.ReadDouble()", (bw, s) => bw.Write(Convert.ToDouble(s)));
            RegisterList("longlist",   "long",   "reader.ReadInt64()",  (bw, s) => bw.Write(Convert.ToInt64(s)));
            RegisterList("stringlist", "string", "reader.ReadString()", (bw, s) => bw.Write(s));
        }

        // ──────────────────────────────────────────────
        //  注册辅助
        // ──────────────────────────────────────────────

        public static void Register(TypeDescriptor desc) =>
            s_map[desc.TypeName.ToLower()] = desc;

        /// <summary>注册基础值类型的快捷方法</summary>
        private static void RegisterPrimitive(
            string typeName,
            string csType,
            string readExpr,
            Action<BinaryWriter, string> writeBinary)
        {
            Register(new TypeDescriptor
            {
                TypeName       = typeName,
                CSharpType     = csType,
                WriteBinary    = writeBinary,
                GenDeserialize = name => $"\t\t{name} = {readExpr};\n",
                GenSerialize   = name => $"\t\twriter.Write({name});\n",
            });
        }

        /// <summary>注册 xxxlist 简写集合类型的快捷方法</summary>
        private static void RegisterList(
            string typeName,
            string elemCsType,
            string readElemExpr,
            Action<BinaryWriter, string> writeElem)
        {
            Register(new TypeDescriptor
            {
                TypeName   = typeName,
                CSharpType = $"List<{elemCsType}>",
                WriteBinary = (bw, v) =>
                {
                    if (string.IsNullOrEmpty(v)) { bw.Write(0); return; }
                    var parts = v.Split(',');
                    bw.Write(parts.Length);
                    foreach (var p in parts) writeElem(bw, p);
                },
                GenDeserialize = name => GenListDeserialize(name, elemCsType, readElemExpr),
                GenSerialize   = name => GenListSerialize(name),
            });
        }

        // ──────────────────────────────────────────────
        //  公共代码生成片段（供 GenModels 中 list<T> 泛型语法共用）
        // ──────────────────────────────────────────────

        /// <summary>生成 List&lt;T&gt; 的 DeSerialize 片段</summary>
        public static string GenListDeserialize(string name, string elemCsType, string readElemExpr)
        {
            var camel = StringExtensions.ToCamelCase(name);
            return
                $"\t\tvar {camel}Count = reader.ReadInt32();\n" +
                $"\t\tif ({camel}Count > 0)\n" +
                "\t\t{\n" +
                $"\t\t\t{name} = new List<{elemCsType}>();\n" +
                $"\t\t\tfor (int i = 0; i < {camel}Count; i++)\n" +
                "\t\t\t{\n" +
                $"\t\t\t\t{name}.Add({readElemExpr});\n" +
                "\t\t\t}\n" +
                "\t\t}\n" +
                "\t\telse\n" +
                "\t\t{\n" +
                $"\t\t\t{name} = null;\n" +
                "\t\t}\n";
        }

        /// <summary>生成普通 List&lt;T&gt; 的 Serialize 片段（writer.Write 直接写元素）</summary>
        public static string GenListSerialize(string name)
        {
            return
                $"\t\tif ({name} == null || {name}.Count == 0)\n" +
                "\t\t{\n" +
                "\t\t\twriter.Write(0);\n" +
                "\t\t}\n" +
                "\t\telse\n" +
                "\t\t{\n" +
                $"\t\t\twriter.Write({name}.Count);\n" +
                $"\t\t\tfor (int i = 0; i < {name}.Count; i++)\n" +
                "\t\t\t{\n" +
                $"\t\t\t\twriter.Write({name}[i]);\n" +
                "\t\t\t}\n" +
                "\t\t}\n";
        }

        private static string GenVectorListDeserialize(string name)
        {
            var camel = StringExtensions.ToCamelCase(name);
            return
                $"\t\tvar {camel}Count = reader.ReadInt32();\n" +
                $"\t\tif ({camel}Count > 0)\n" +
                "\t\t{\n" +
                $"\t\t\t{name} = new List<List<float>>();\n" +
                $"\t\t\tfor (int i = 0; i < {camel}Count; i++)\n" +
                "\t\t\t{\n" +
                "\t\t\t\tvar tempList = new List<float>();\n" +
                "\t\t\t\tvar tempListCount = reader.ReadInt32();\n" +
                "\t\t\t\tfor (int j = 0; j < tempListCount; j++)\n" +
                "\t\t\t\t{\n" +
                "\t\t\t\t\ttempList.Add(reader.ReadSingle());\n" +
                "\t\t\t\t}\n" +
                $"\t\t\t\t{name}.Add(tempList);\n" +
                "\t\t\t}\n" +
                "\t\t}\n" +
                "\t\telse\n" +
                "\t\t{\n" +
                $"\t\t\t{name} = null;\n" +
                "\t\t}\n";
        }

        private static string GenVectorListSerialize(string name)
        {
            return
                $"\t\tif ({name} == null || {name}.Count == 0)\n" +
                "\t\t{\n" +
                "\t\t\twriter.Write(0);\n" +
                "\t\t}\n" +
                "\t\telse\n" +
                "\t\t{\n" +
                $"\t\t\twriter.Write({name}.Count);\n" +
                $"\t\t\tfor (int i = 0; i < {name}.Count; i++)\n" +
                "\t\t\t{\n" +
                $"\t\t\t\twriter.Write({name}[i].Count);\n" +
                $"\t\t\t\tfor (int j = 0; j < {name}[i].Count; j++)\n" +
                "\t\t\t\t{\n" +
                $"\t\t\t\t\twriter.Write({name}[i][j]);\n" +
                "\t\t\t\t}\n" +
                "\t\t\t}\n" +
                "\t\t}\n";
        }
    }
}
