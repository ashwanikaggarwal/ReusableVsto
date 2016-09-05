using System;
using IExcel = Microsoft.Office.Interop.Excel;

namespace Reusable.Excel.PivotTables
{
    public interface IPivotTableParam : IDisposable
    {
        Microsoft.Office.Interop.Excel.Range Target { get; set; }
        Microsoft.Office.Interop.Excel.PivotCache PivotCache { get; set; }
        
        String ColumnField { get; set; }
        String PageField { get; set; }
        String RowField { get; set; }

        IExcel.XlSortOrder RowSortOrder { get; set; }

        String DataField { get; set; }
        Microsoft.Office.Interop.Excel.XlConsolidationFunction Function { get; set; }
        
        Boolean ColumnGrand { get; set;}
        Boolean RowGrand { get; set; }
    }
}
