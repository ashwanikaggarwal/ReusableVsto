using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.PivotTables
{
    public class PivotTable : IPivotTable, IDisposable
    {
        private bool disposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~PivotTable()
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
                parameters.Dispose();
            }

            // release any unmanaged objects
            // set the object references to null
            pivot = null;

            disposed = true;
        }

        IExcel.PivotTable pivot;
        public IExcel.PivotTable Pivot
        {
            get { return pivot; }
            set { pivot = value; }
        }
        
        private IPivotTableParam parameters;
        public PivotTable(IPivotTableParam param)
        {
            parameters = param;
        }

        public void Build()
        {
            CreatePivotTable();
        }

        private void CreatePivotTable()
        {
            IExcel.PivotTable pvt = null;
            try
            {
                pvt = parameters.PivotCache.CreatePivotTable(parameters.Target, Type.Missing, Type.Missing, Type.Missing);

                if (parameters.PageField.Trim().Length > 0)
                {
                    IExcel.PivotField page = pvt.PivotFields(parameters.PageField);
                    page.Orientation = IExcel.XlPivotFieldOrientation.xlPageField;
                }

                if (parameters.RowField.Trim().Length > 0)
                {
                    IExcel.PivotField row = pvt.PivotFields(parameters.RowField);
                    row.Orientation = IExcel.XlPivotFieldOrientation.xlRowField;

                    row.AutoSort((int)parameters.RowSortOrder,parameters.RowField);
                }

                if (parameters.ColumnField.Trim().Length > 0)
                {
                    IExcel.PivotField column = pvt.PivotFields(parameters.ColumnField);
                    column.Orientation = IExcel.XlPivotFieldOrientation.xlColumnField;
                }

                if (parameters.DataField.Trim().Length > 0)
                {
                    IExcel.PivotField data = pvt.PivotFields(parameters.DataField);
                    data.Orientation = IExcel.XlPivotFieldOrientation.xlDataField;
                    data.Function = parameters.Function;
                }

                pvt.ColumnGrand = parameters.ColumnGrand;
                pvt.RowGrand = parameters.RowGrand;
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }

            this.pivot = pvt;
        }
    }
}
