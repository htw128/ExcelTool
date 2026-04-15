using System;
using System.Collections.Generic;
using System.IO;
using System.Text;

namespace ExcelTool
{
    /// <summary>
    /// 文件操作类
    /// </summary>
    public static class FileManager
    {
        /// <summary>将数据写入二进制文件（IBinarySerializable 版本）</summary>
        public static bool WriteBinaryDataToFile(string filePath, IBinarySerializable data)
        {
            if (string.IsNullOrEmpty(filePath)) return false;
            if (File.Exists(filePath)) File.Delete(filePath);

            using var fileStream = new FileStream(filePath, FileMode.Create);
            using var bw = new BinaryWriter(fileStream);
            data.Serialize(bw);
            bw.Flush();
            return true;
        }

        /// <summary>
        /// 将数据写入二进制文件。
        /// 类型名（Item1）须与 <see cref="TypeRegistry"/> 中注册的 key 一致（不区分大小写）。
        /// 额外支持 list&lt;T&gt; 泛型语法。
        /// </summary>
        public static bool WriteBinaryDatasToFile(string filePath, List<Tuple<string, string>> datas)
        {
            try
            {
                if (string.IsNullOrEmpty(filePath)) return false;
                if (File.Exists(filePath)) File.Delete(filePath);

                using var fileStream = new FileStream(filePath, FileMode.Create);
                using var bw = new BinaryWriter(fileStream);

                foreach (var (rawType, value) in datas)
                {
                    var type = rawType.ToLower();

                    // ── 优先从注册表查找 ──────────────────────────────
                    var desc = TypeRegistry.Get(type);
                    if (desc != null)
                    {
                        desc.WriteBinary(bw, value);
                        continue;
                    }

                    // ── list<T> 泛型语法（如 list<int>）─────────────
                    if (type.StartsWith("list<") && type.EndsWith(">"))
                    {
                        WriteGenericList(bw, type, value);
                        continue;
                    }

                    $"写入二进制文件：数据类型 \"{rawType}\" 未注册，请在 TypeRegistry 中添加".WriteErrorLine();
                    return false;
                }

                bw.Flush();
                return true;
            }
            catch (Exception ex)
            {
                ex.ToString().WriteErrorLine();
                return false;
            }
        }

        // ──────────────────────────────────────────────────────────────
        //  list<T> 泛型列表写入（T 必须是已在 TypeRegistry 注册的基础类型）
        // ──────────────────────────────────────────────────────────────
        private static void WriteGenericList(BinaryWriter bw, string fullType, string value)
        {
            // fullType 形如 "list<int>"
            var innerType = fullType[5..^1]; // 去掉 "list<" 和 ">"
            var elemDesc  = TypeRegistry.Get(innerType);
            if (elemDesc == null)
            {
                $"list<T> 的元素类型 \"{innerType}\" 未注册，请在 TypeRegistry 中添加".WriteErrorLine();
                return;
            }

            if (string.IsNullOrEmpty(value))
            {
                bw.Write(0);
                return;
            }

            var parts = value.Split('|');
            bw.Write(parts.Length);
            foreach (var p in parts)
                elemDesc.WriteBinary(bw, p);
        }

        // ──────────────────────────────────────────────────────────────
        //  其余文件工具方法（不涉及类型系统，保持不变）
        // ──────────────────────────────────────────────────────────────

        public static bool ReadBinaryDataFromBytes(byte[] bytes, ref IBinarySerializable data)
        {
            if (bytes == null) return false;
            using var ms = new MemoryStream(bytes);
            using var br = new BinaryReader(ms);
            data.DeSerialize(br);
            return true;
        }

        public static bool ReadBinaryDataFromFile(string filePath, ref IBinarySerializable data)
        {
            if (string.IsNullOrEmpty(filePath)) return false;
            using var fs = new FileStream(filePath, FileMode.Open);
            using var br = new BinaryReader(fs);
            data.DeSerialize(br);
            return true;
        }

        public static bool WriteBytesToFile(string filePath, byte[] data)
        {
            if (string.IsNullOrEmpty(filePath)) return false;
            if (File.Exists(filePath)) File.Delete(filePath);
            using Stream sw = new FileInfo(filePath).Create();
            sw.Write(data, 0, data.Length);
            sw.Flush();
            return true;
        }

        public static bool WriteToFile(string filePath, string context) =>
            WriteToFile(filePath, context, new UTF8Encoding(false));

        public static bool WriteToFile(string filePath, string context, Encoding encoding)
        {
            if (string.IsNullOrEmpty(filePath)) return false;
            if (File.Exists(filePath)) File.Delete(filePath);
            using var fs = new FileStream(filePath, FileMode.Create);
            var data = encoding.GetBytes(context);
            fs.Write(data, 0, data.Length);
            fs.Flush();
            return true;
        }

        public static string ReadAllByLine(string path)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return string.Empty;
            var sb = new StringBuilder();
            using var sr = new StreamReader(path, Encoding.Default);
            string line;
            while ((line = sr.ReadLine()) != null)
                sb.AppendLine(line);
            return sb.ToString();
        }

        public static byte[] ReadAllBytes(string path) =>
            string.IsNullOrEmpty(path) || !File.Exists(path) ? null : File.ReadAllBytes(path);

        public static void ReplaceContent(string path, string normalStr, string newStr)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return;
            File.WriteAllText(path, File.ReadAllText(path).Replace(normalStr, newStr));
        }

        public static void ReplaceContent(string path, string newStr, params string[] normalStrs)
        {
            if (string.IsNullOrEmpty(path) || !File.Exists(path)) return;
            string content = File.ReadAllText(path);
            foreach (var s in normalStrs) content = content.Replace(s, newStr);
            File.WriteAllText(path, content);
        }
    }
}
