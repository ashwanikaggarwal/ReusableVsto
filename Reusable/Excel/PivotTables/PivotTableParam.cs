using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.PivotTables
{
    public class PivotTableParam : IDisposable, Reusable.Excel.PivotTables.IPivotTableParam
    {
        private bool disposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~PivotTableParam()
        {
            Dispose(false);
        }

        protected virtual void Dispose(bool disposing)
        {
            if (disposed)
                return;

            if (disposing)
            {
                // free other managed objects that implement
                // IDisposable only
            }

            // release any unmanaged objects
            // set the object references to null
            target = null;
            pivotCache = null;

            disposed = true;
        }

        IExcel.Range target;
        public IExcel.Range Target
        {
            get { return target; }
            set { target = value; }
        }

        IExcel.PivotCache pivotCache;
        public IExcel.PivotCache PivotCache
        {
            get { return pivotCache; }
            set { pivotCache = value; }
        }

        string pageField;
        public string PageField
        {
            get { return pageField; }
            set { pageField = value; }
        }

        string rowField;
        public string RowField
        {
            get { return rowField; }
            set { rowField = value; }
        }

        IExcel.XlSortOrder rowSortOrder = IExcel.XlSortOrder.xlAscending;
        public IExcel.XlSortOrder RowSortOrder
        {
            get { return rowSortOrder; }
            set { rowSortOrder = value; }
        }

        string columnField;
        public string ColumnField
        {
            get { return columnField; }
            set { columnField = value; }
        }

        string dataField;
        public string DataField
        {
            get { return dataField; }
            set { dataField = value; }
        }

        IExcel.XlConsolidationFunction function;
        public IExcel.XlConsolidationFunction Function
        {
            get { return function; }
            set { function = value; }
        }

        Boolean columnGrand;
        public Boolean ColumnGrand
        {
            get { return columnGrand; }
            set { columnGrand = value; }
        }

        Boolean rowGrand;
        public Boolean RowGrand
        {
            get { return rowGrand; }
            set { rowGrand = value; }
        }

        public PivotTableParam()
        {
        }
    }
}
