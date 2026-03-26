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
            Option<DirectoryInfo> inputOption = new(
                name: "--input",
                description: "Excel 文件所在目录",
                getDefaultValue: () => new DirectoryInfo(AppDomain.CurrentDomain.BaseDirectory)
                );
            
            Option<DirectoryInfo> outputCodeDirOption = new(
                name: "--outputCodeDir",
                description: "CS 模板代码输出目录"
            );

            Option<DirectoryInfo> outputDataDirOption = new(
                name: "--outputDataDir",
                description: "二进制数据输出目录"
            );

            Option<string> namespaceOption = new(
                name: "--namespace",
                description: "生成代码的命名空间",
                getDefaultValue: () => ""
            );

            RootCommand rootCommand = new();
            rootCommand.AddOption(inputOption);
            rootCommand.AddOption(outputCodeDirOption);
            rootCommand.AddOption(outputDataDirOption);
            rootCommand.AddOption(namespaceOption);
            
            rootCommand.SetHandler(ExcelProcess.Run, inputOption, outputCodeDirOption, outputDataDirOption, namespaceOption);
            // TODO 单元测试

            return await rootCommand.InvokeAsync(args);
        }
    }

    static class ExcelProcess
    {
        public static void Run(DirectoryInfo inputDir, DirectoryInfo? outputCodeDir, DirectoryInfo? outputDataDir,
            string nameSpace)
        {
            string path = inputDir.FullName;
            string codeDir = outputCodeDir?.FullName ?? outputDataDir?.FullName ?? path;
            string dataDir = outputDataDir?.FullName ?? outputCodeDir?.FullName ?? path;

            DirectoryInfo dirInfo = new(path);
            FileInfo[] excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);
            if (excels.Length <= 0)
            {
                "当前exe目录或者目标目录没有excels文件,请重新设置目录".WriteErrorLine();
            }
            else
            {
                "==========================================================".WriteSuccessLine();
                "== 根据xlsx生成模板代码和二进制文件工具                 ==".WriteSuccessLine();
                "== 说明:将exe放在xlsx目录中或者exe或者传入根目录        ==".WriteSuccessLine();
                "==========================================================".WriteSuccessLine();

                excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);

                //读取
                foreach (FileInfo file in excels)
                {
                    if (file.Name.StartsWith("~$")) continue;

                    List<ParsedSheet> sheets = ExcelHelper.ParseAllSheets(file.FullName, out List<ParsedEnum>? enumSheets);

                    //生成CS文件
                    bool res = GenModels.GenCSharpModel(sheets, codeDir, nameSpace);
                    if (res)
                    {
                        $"{file.Name}CS模板生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{file.Name}CS模板生成失败".WriteErrorLine();
                    }
                    
                    // 生成CS枚举文件
                    if (enumSheets.Count > 0)
                    {
                        bool enumRes = GenEnums.GenCSharpEnum(enumSheets, codeDir, nameSpace);
                        if (enumRes)
                            $"{file.Name} 枚举代码生成成功".WriteSuccessLine();
                        else
                            $"{file.Name} 枚举代码生成失败".WriteErrorLine();

                        bool enumIdsRes = GenEnums.GenEnumIds(enumSheets, codeDir, nameSpace);
                        if (enumIdsRes)
                            $"{file.Name} EnumIds 生成成功".WriteSuccessLine();
                        else
                            $"{file.Name} EnumIds 生成失败".WriteErrorLine();
                    }

                    //生成二进制文件，如果list或者vector数据为空则写入0，要根据类型来读取csv的字段数据强转成对应的数据类型然后写入
                    res = TableExcelExportBytes.ExportToFile(sheets, dataDir);
                    if (res)
                    {
                        $"{file.Name}二进制数据生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{file.Name}二进制数据生成失败".WriteErrorLine();
                    }
                }
            }
        }
    }
}

