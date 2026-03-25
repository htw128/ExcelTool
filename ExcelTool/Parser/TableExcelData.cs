using System.Collections.Generic;
using System.Linq;

namespace ExcelTool.Parser
{
    public class TableExcelData
    {
        public readonly int ColumnCount = 0;
        public readonly int RowCounts = 0;
        
        public List<TableExcelHeader> Headers { get; }
        public List<TableExcelRow> Rows { get; }

        public TableExcelData(IEnumerable<TableExcelHeader> headers, IEnumerable<TableExcelRow> rows)
        {
            Headers = headers.ToList();
            Rows = rows.ToList();
            this.ColumnCount = Headers.Count;
            this.RowCounts = Rows.Count;
        }


        //TODO:待检查数据类型的合法性
        //public bool CheckUnique(out string errorMsg)
        //{

        //}
    }
}
