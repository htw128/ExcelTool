using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
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
            
            
            DirectoryInfo dirInfo = new(path);
            FileInfo[] csvs = dirInfo.GetFiles("*.csv", SearchOption.AllDirectories);
            FileInfo[] excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);
            if ((csvs.Length <= 0) && (excels.Length <= 0))
            {
                "当前exe目录或者目标目录没有csv文件或者excels文件,请重新设置目录".WriteErrorLine();
            }
            else
            {
                "==========================================================".WriteSuccessLine();
                "== 根据csv/xlsx生成模板代码和二进制文件工具             ==".WriteSuccessLine();
                "== 说明:将exe放在csv/xlsx目录中或者exe或者传入csv根目录 ==".WriteSuccessLine();
                "==========================================================".WriteSuccessLine();

                List<string> genExcels = [];
                foreach (FileInfo csv in csvs)
                {
                    //生成对应的xlsx文件
                    var tempPath = CsvHelper.CsvToXlsx(csv.FullName);
                    if (string.IsNullOrEmpty(tempPath))
                    {
                        $"csv:{csv.FullName}生成xlsx文件出错".WriteErrorLine();
                    }
                    else
                    {
                        genExcels.Add(tempPath);
                    }
                }
                excels = dirInfo.GetFiles("*.xlsx", SearchOption.AllDirectories);

                //读取
                foreach (var file in excels)
                {
                    if (file.Name.StartsWith("~$")) return;
                    
                    //生成CS文件
                    bool res = GenModels.GenCSharpModel(file.FullName, outputCodeDir, nameSpace);
                    if (res)
                    {
                        $"{file.Name}CS模板生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{file.Name}CS模板生成失败".WriteErrorLine();
                    }

                    //生成二进制文件，如果list或者vector数据为空则写入0，要根据类型来读取csv的字段数据强转成对应的数据类型然后写入
                    res = TableExcelExportBytes.ExportToFile(file.FullName, outputDataDir);
                    if (res)
                    {
                        $"{file.Name}二进制数据生成成功".WriteSuccessLine();
                    }
                    else
                    {
                        $"{file.Name}二进制数据生成失败".WriteErrorLine();
                    }
                }

                //删除生成的excels
                for (int i = genExcels.Count - 1; i >= 0; i--)
                {
                    File.Delete(genExcels[i]);
                }

                Dictionary<int, string> dics = new();
                new List<string>(dics.Values);

                //读取测试
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

