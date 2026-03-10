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
                    var tableData = ExcelHelper.ExcelDatas(fileName, out string sheetName, out int sheetCount, sheetNum);
                    if (tableData == null || sheetName.StartsWith($"#") || sheetNum > sheetCount)
                        break;

                    List<Tuple<string, string>> datas = new();
                    //先写入行数，然后每一行的数据一次写入  小写类型、字符串
                    Tuple<string, string> rowCount = new("int", tableData.RowCounts.ToString());
                    datas.Add(rowCount);
                    foreach (var row in tableData.Rows)
                    {
                        for (int i = 0; i < tableData.CollonCount; i++)
                        {
                            var type = tableData.Headers[i].FieldType.ToLower();
                            var data = row.StrList[i];
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
    }
}
