using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.QueryTables
{
    public interface IQueryTableParam : IDisposable
    {
        Boolean Format {get; set; }

        Boolean PreserveColumnInfo { get; set; }
        Boolean FillAdjacentFormulas { get; set; }
        Boolean PreserveFormatting { get; set; }

        Int32 FirstColumnToSortBy { get; set; }
        Int32 SecondColumnToSortBy { get; set; }
        Int32 ThirdColumnToSortBy { get; set; }

        Microsoft.Office.Interop.Excel.Worksheet TargetSheet { get; set; }
        Microsoft.Office.Interop.Excel.Range Target { get; set; }

        String ConnectionString { get; set;}
        String SqlStatement  { get; set;}
    }
}
