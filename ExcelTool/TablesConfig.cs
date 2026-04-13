using System;
using System.Collections.Generic;
using System.IO;
using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;

namespace ExcelTool;

public static class TablesConfig
{
    internal static List<TableEntry> Parse(string tablesXlsxPath)
    {
        List<TableEntry> tableEntries = [];
        
        string tablesDir = Path.GetDirectoryName(Path.GetFullPath(tablesXlsxPath))!;

        try
        {
            using FileStream fs = File.OpenRead(tablesXlsxPath);
            XSSFWorkbook workbook = new(fs);
            
            ISheet sheet = workbook.GetSheetAt(0);
            
            for (int i = 4; i <= sheet.LastRowNum; i++)
            {
                IRow row = sheet.GetRow(i);
                if (row == null || string.IsNullOrEmpty(ExcelHelper.GetCellValue(row.GetCell(0)))) continue;

                TableEntry tableEntry = new()
                {
                    InputFile     = ResolveDir(ExcelHelper.GetCellValue(row.GetCell(0))),
                    Namespace     = ExcelHelper.GetCellValue(row.GetCell(1)),
                    OutputCodeDir = ResolveDir(ExcelHelper.GetCellValue(row.GetCell(2))),
                    OutputDataDir = ResolveDir(ExcelHelper.GetCellValue(row.GetCell(3))),
                    
                };
                
                tableEntries.Add(tableEntry);
            }
            
        }
        catch (Exception exception)
        {
            exception.ToString().WriteErrorLine();
        }
        
        return tableEntries;

        string ResolveDir(string raw)
        {
            if (raw != null)
            {
                return Path.IsPathRooted(raw)
                    ? raw
                    : Path.GetFullPath(Path.Combine(tablesDir, raw));
            }
            return null;
        }
    }
}

public class TableEntry
{
    public string InputFile    { get; init; }  // e.g. "AudioConfigs.xlsx"
    public string Namespace    { get; init; }  // e.g. "OCES.Audio"
    public string OutputCodeDir { get; init; } 
    public string OutputDataDir { get; init; }
}
