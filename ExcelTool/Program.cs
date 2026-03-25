using System;
using System.Collections.Generic;
using System.IO;
using ExcelTool.Parser;

namespace ExcelTool
{
    class Program
    {
        static void Main(string[] args)
        {
            string path = AppDomain.CurrentDomain.BaseDirectory;
            string outputCodeDir = "";
            string outputDataDir = "";
            string nameSpace = "";

            foreach (var arg in args)
            {
                if (arg.StartsWith("--input="))
                {
                    path = arg["--input=".Length..];
                }
                else if (arg.StartsWith("--outputCodeDir="))
                {
                    outputCodeDir = arg["--outputCodeDir=".Length..];

                }
                else if (arg.StartsWith("--outputDataDir="))
                {
                    outputDataDir = arg["--outputDataDir=".Length..];
                }
                else if (arg.StartsWith("--namespace="))
                {
                    nameSpace = arg["--namespace=".Length..];
                }
            }
            
            if (string.IsNullOrEmpty(outputCodeDir) && string.IsNullOrEmpty(outputDataDir))
            {
                outputDataDir = outputCodeDir = path;
            }else if (string.IsNullOrEmpty(outputCodeDir))
            {
                outputCodeDir = outputDataDir;
            }else if (string.IsNullOrEmpty(outputDataDir))
            {
                outputDataDir = outputCodeDir;
            }
            // TODO 使用System.CommandLine重构参数解析
            
            DirectoryInfo dirInfo = new(path);
            FileInfo[] excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);
            if (excels.Length <= 0)
            {
                "当前exe目录或者目标目录没有excels文件,请重新设置目录".WriteErrorLine();
            }
            else
            {
                "==========================================================".WriteSuccessLine();
                "== 根据xlsx生成模板代码和二进制文件工具             ==".WriteSuccessLine();
                "== 说明:将exe放在xlsx目录中或者exe或者传入根目录 ==".WriteSuccessLine();
                "==========================================================".WriteSuccessLine();

                excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);

                //读取
                foreach (FileInfo file in excels)
                {
                    if (file.Name.StartsWith("~$")) continue;

                    List<ParsedSheet> sheets = ExcelHelper.ParseAllSheets(file.FullName);
                    
                    //生成CS文件
                    bool res = GenModels.GenCSharpModel(sheets, outputCodeDir, nameSpace);
                    if (res)
                    {
                        $"{file.Name}CS模板生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{file.Name}CS模板生成失败".WriteErrorLine();
                    }

                    //生成二进制文件，如果list或者vector数据为空则写入0，要根据类型来读取csv的字段数据强转成对应的数据类型然后写入
                    res = TableExcelExportBytes.ExportToFile(sheets, outputDataDir);
                    if (res)
                    {
                        $"{file.Name}二进制数据生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{file.Name}二进制数据生成失败".WriteErrorLine();
                    }
                }
                
                // TODO 抽象 ExcelProcess 类 处理相关流程

                // TODO 单元测试
                //IBinarySerializable newavList = new avatarguideTestConfig();
                //var readOK = FileManager.ReadBinaryDataFromFile(Path.Combine(path, "avatarguideTest.bytes"), ref newavList);
                //if (readOK)
                //{
                //    ConsoleHelper.WriteInfoLine(newavList.ToString());
                //    var d = (newavList as avatarguideTestConfig).QueryById(1).ToList();
                //    ConsoleHelper.WriteInfoLine(d[0].gender);
                //}
                //else
                //{
                //    ConsoleHelper.WriteErrorLine("读取失败");
                //}
            }
        }
    }
}

