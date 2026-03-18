using System;
using System.Collections.Generic;
using System.IO;

namespace ExcelTool.Parser
{
    public static class TableExcelExportBytes
    {
        public static bool ExportToFile(string fileName, string outputDir = null)
        {
            try
            {
                FileInfo fileInfo = new(fileName);
                if (string.IsNullOrEmpty(outputDir))
                {
                    outputDir = fileInfo.DirectoryName;
                }
                
                for (int sheetNum = 0; ; sheetNum++)
                {
                    var tableData = ExcelHelper.ExcelData(fileName, out string sheetName, out int sheetCount, sheetNum);
                    if (tableData == null || sheetNum >= sheetCount)
                        break;

                    // # 开头的 sheet 只跳过
                    if (sheetName.StartsWith("#"))
                        continue;

                    List<Tuple<string, string>> datas = [];
                    //先写入行数，然后每一行的数据一次写入  小写类型、字符串
                    Tuple<string, string> rowCount = new("int", tableData.RowCounts.ToString());
                    datas.Add(rowCount);
                    foreach (var row in tableData.Rows)
                    {
                        for (int i = 0; i < tableData.CollonCount; i++)
                        {
                            var type = tableData.Headers[i].FieldType.ToLower();
                            var data = row.StrList[i];

                            // 未在 TypeRegistry 注册的类型（且不是 list<T>）当枚举处理，写入 byte
                            if (!TypeRegistry.Contains(type) && !IsGenericList(type))
                            {
                                type = "byte";
                            }

                            datas.Add(new Tuple<string, string>(type, data));
                        }
                    }
                    var binaryFilePath = Path.Combine(outputDir, $"{sheetName}.bytes");
                    FileManager.WriteBinaryDatasToFile(binaryFilePath, datas);
                }
                return true;
            }
            catch (Exception ex)
            {
                ConsoleHelper.WriteErrorLine(ex.ToString());
                return false;
            }
        }

        private static bool IsGenericList(string type) =>
            type.StartsWith("list<") && type.EndsWith(">");
    }
}
