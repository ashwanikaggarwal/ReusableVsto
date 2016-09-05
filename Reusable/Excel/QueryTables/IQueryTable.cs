using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.QueryTables
{
    public interface IQueryTable: IDisposable
    {
        Microsoft.Office.Interop.Excel.QueryTable Table { get; }

        void Build();
        void Update();
    }
}
