using NPOI.SS.UserModel;
using NPOI.XSSF.UserModel;
using System;
using System.Collections.Generic;
using System.IO;
using ExcelTool.Parser;

namespace ExcelTool
{
    public static class ExcelHelper
    {
        public static List<TableExcelHeader> ExcelHeaders(string fileName, out string sheetName, out int sheetCount, int sheetNum = 0)
        {
            try
            {
                List<TableExcelHeader> headers = new();

                using FileStream fs = File.OpenRead(fileName);
                IWorkbook wk = new XSSFWorkbook(fs);
                
                sheetCount = wk.NumberOfSheets;
                if (sheetNum >= sheetCount)
                {
                    sheetName = "";
                    return null;
                }
                
                ISheet sheet = wk.GetSheetAt(sheetNum);
                sheetName = sheet.SheetName;
                
                IRow nameRow = sheet.GetRow(0);   // 字段名
                IRow typeRow = sheet.GetRow(1);   // 类型
                IRow descRow = sheet.GetRow(2);   // 注释

                for (int j = 0; j < nameRow.LastCellNum; j++)
                {
                    string fieldName = nameRow.GetCell(j)?.ToString() ?? "";
                    string fieldType = typeRow.GetCell(j)?.ToString() ?? "";
                    string fieldDesc = descRow.GetCell(j)?.ToString() ?? "";
                    
                    if (string.IsNullOrEmpty(fieldName))
                    {
                        $"列 {j} 字段名为空".WriteErrorLine();
                        continue;
                    }

                    headers.Add(new TableExcelHeader()
                    {
                        FieldName = fieldName,
                        FieldType = fieldType,
                        FieldDesc = fieldDesc,
                    });
                }

                return headers;
            }
            catch (Exception ex)
            {
                ex.ToString().WriteErrorLine();
                sheetName = null;
                sheetCount = 0;
                return null;
            }
        }

        public static TableExcelData ExcelDatas(string fileName, out string sheetName, out int sheetCount, int sheetNum = 0)
        {
            try
            {
                var excelHeader = ExcelHeaders(fileName, out sheetName, out sheetCount);
                var tableRows = new List<TableExcelRow>();

                using FileStream fs = File.OpenRead(fileName);
                IWorkbook wk = new XSSFWorkbook(fs);

                if (sheetNum >= sheetCount)
                {
                    return null;
                }
                ISheet sheet = wk.GetSheetAt(sheetNum);

                for (int i = 6; i <= sheet.LastRowNum; i++)
                {
                    IRow row = sheet.GetRow(i);
                    if (row == null) continue;

                    var tableExcelRow = new TableExcelRow();

                    for (int j = 0; j < excelHeader.Count; j++)
                    {
                        var cell = row.GetCell(j);
                        tableExcelRow.Add(GetCellValue(cell));
                    }

                    tableRows.Add(tableExcelRow);
                }

                return new TableExcelData(excelHeader, tableRows);
            }
            catch (Exception ex)
            {
                ex.ToString().WriteErrorLine();
                sheetName = null;
                sheetCount = 0;
                return null;
            }
        }
        private static string GetCellValue(ICell cell)
        {
            if (cell == null)
                return "";

            switch (cell.CellType)
            {
                case CellType.String:
                    return cell.StringCellValue;

                case CellType.Numeric:
                    if (DateUtil.IsCellDateFormatted(cell))
                    {
                        return cell.DateCellValue.ToString();
                    }
                    return cell.NumericCellValue.ToString();

                case CellType.Boolean:
                    return cell.BooleanCellValue ? "1" : "0";

                case CellType.Formula:
                    return cell.ToString();

                default:
                    return "";
            }
        }
    }
}
