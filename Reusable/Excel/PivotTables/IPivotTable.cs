using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.PivotTables
{
    public interface IPivotTable : IDisposable
    {
        IExcel.PivotTable Pivot {get; set;}
        void Build();
    }
}
