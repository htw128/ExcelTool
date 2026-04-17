#nullable enable
using System;
using System.Collections.Generic;
using System.IO;
using System.Threading.Tasks;
using System.CommandLine;
using ExcelTool.Parser;

namespace ExcelTool
{
    class Program
    {
        static async Task<int> Main(string[] args)
        {
            Option<FileInfo> tablesOption = new(
                name: "--tables",
                description: "__tables__.xlsx 的路径",
                getDefaultValue: () => new FileInfo(AppDomain.CurrentDomain.BaseDirectory)
                );

            RootCommand rootCommand = new();
            rootCommand.AddOption(tablesOption);
            
            rootCommand.SetHandler(ExcelProcess.Run, tablesOption);

            return await rootCommand.InvokeAsync(args);
        }
    }

    static class ExcelProcess
    {
        public static void Run(FileInfo tablesFile)
        {
            if(!tablesFile.Exists)
            {
                $"未在 {tablesFile} 处找到文件！".WriteErrorLine();
                return;
            }
            
            
            string path = tablesFile.FullName;
            List<TableEntry> tableEntries = TablesConfig.Parse(path);
            
            if (tableEntries.Count <= 0)
            {
                "__tables__.xlsx内没有数据".WriteErrorLine();
            }
            else
            {
                "==========================================================".WriteSuccessLine();
                "== 根据xlsx生成模板代码和二进制文件工具                 ==".WriteSuccessLine();
                "== 说明:将exe放在xlsx目录中或者exe或者传入根目录        ==".WriteSuccessLine();
                "==========================================================".WriteSuccessLine();

                //读取
                foreach (TableEntry tableEntry in tableEntries)
                {
                    if (!Path.Exists(tableEntry.InputFile))
                    {
                        $"{tableEntry.InputFile} 不存在！".WriteWarningLine();
                        continue;
                    }

                    if (!Path.Exists(tableEntry.OutputCodeDir))
                    {
                        Directory.CreateDirectory(tableEntry.OutputCodeDir);
                    }

                    if (!Path.Exists(tableEntry.OutputDataDir))
                    {
                        Directory.CreateDirectory(tableEntry.OutputDataDir);
                    }
                    
                    string fileName = Path.GetFileName(tableEntry.InputFile);
                    
                    List<ParsedSheet> sheets = ExcelHelper.ParseAllSheets(tableEntry.InputFile, out List<ParsedEnum>? enumSheets);

                    //生成CS文件
                    bool res = GenModels.GenCSharpModel(sheets, tableEntry.OutputCodeDir, tableEntry.Namespace);
                    if (res)
                    {
                        $"{fileName}CS模板生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{fileName}CS模板生成失败".WriteErrorLine();
                    }
                    
                    // 生成 AudioConsts.cs（合并枚举、枚举ID、AudioObject定义和CUE映射）
                    bool audioConstsRes = GenAudioConsts.Gen(sheets, enumSheets, tableEntry.OutputCodeDir, tableEntry.Namespace);
                    if (audioConstsRes)
                        $"{fileName} AudioConsts 生成成功".WriteSuccessLine();
                    else
                        $"{fileName} AudioConsts 生成失败".WriteErrorLine();

                    //生成二进制文件，如果list或者vector数据为空则写入0，要根据类型来读取csv的字段数据强转成对应的数据类型然后写入
                    res = TableExcelExportBytes.ExportToFile(sheets, tableEntry.OutputDataDir);
                    if (res)
                    {
                        $"{fileName}二进制数据生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{fileName}二进制数据生成失败".WriteErrorLine();
                    }
                }
            }
        }
    }
}

