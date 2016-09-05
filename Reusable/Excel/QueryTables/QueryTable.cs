using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.QueryTables
{
    public class QueryTable : IQueryTable
    {
        private bool disposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        protected virtual void Dispose(bool disposing)
        {
            // Check to see if Dispose has already been called. 
            if (!this.disposed)
            {
                // If disposing equals true, dispose all managed 
                // and unmanaged resources. 
                if (disposing)
                {
                    // Dispose managed resources.
                    parameters.Dispose();
                }
            }
            disposed = true;
        }

        IQueryTableParam parameters;
        public QueryTable(IQueryTableParam param)
        {
            parameters = param;
        }

        private Microsoft.Office.Interop.Excel.QueryTable table;
        public Microsoft.Office.Interop.Excel.QueryTable Table
        {
            get { return table; }
            set { table = value; }
        }

        public void Build()
        {
            try
            {
                IExcel.QueryTables qts = (IExcel.QueryTables)parameters.TargetSheet.QueryTables;
                IExcel.QueryTable qt = qts.Add(parameters.ConnectionString, parameters.Target, parameters.SqlStatement);

                SetQueryTableDisplayProperties(qt);

                Refresh(qt);

                this.Table = qt;
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }

        }

        public void Update()
        {
            IExcel.QueryTables qts = null;
            IExcel.QueryTable qt = null;

            try
            {
                qts = (IExcel.QueryTables)parameters.TargetSheet.QueryTables;
                qt = qts[1];
                qt.CommandText = parameters.SqlStatement;

                //on update replaces existing connction string
                qt.Connection = parameters.ConnectionString;

                SetQueryTableDisplayProperties(qt);

                Refresh(qt);

                this.Table = qt;
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }
        }
               
        private void Refresh(IExcel.QueryTable qt)
        {
            try
            {
                qt.Refresh();
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }
            finally
            {
            }
        }

        private void SetQueryTableDisplayProperties(IExcel.QueryTable qt)
        {
            qt.PreserveColumnInfo = parameters.PreserveColumnInfo;
            qt.FillAdjacentFormulas = parameters.FillAdjacentFormulas;
            qt.PreserveFormatting = parameters.PreserveFormatting;
        }
    }
}
