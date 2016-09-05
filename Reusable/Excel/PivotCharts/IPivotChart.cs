using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.PivotCharts
{
    public interface IPivotChart : IDisposable
    {
        IExcel.ChartObject Chart { get;}
        void Build();
        void Update(IExcel.ChartObject cht);
    }
}
