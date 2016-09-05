using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;

namespace Reusable.Excel.QueryTables
{
    public class QueryTableParam : IQueryTableParam
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
                    //parameters.Dispose();
                    target = null;
                    targetSheet = null;
                }
            }
            disposed = true;
        }

        private bool format;
        public bool Format
        {
            get { return format; }
            set { format = value; }
        }

        private Microsoft.Office.Interop.Excel.Worksheet targetSheet;
        public Microsoft.Office.Interop.Excel.Worksheet TargetSheet
        {
            get { return targetSheet; }
            set { targetSheet = value; }
        }

        private Microsoft.Office.Interop.Excel.Range target;
        public Microsoft.Office.Interop.Excel.Range Target
        {
            get { return target; }
            set { target = value; }
        }

        private string connectionString;
        public string ConnectionString
        {
            get { return connectionString; }
            set { connectionString = value; }
        }

        private string sqlStatement;
        public string SqlStatement
        {
            get { return sqlStatement; }
            set { sqlStatement = value; }
        }

        private bool preserveColumnInfo;
        public bool PreserveColumnInfo
        {
            get { return preserveColumnInfo; }
            set { preserveColumnInfo = value; }
        }

        private bool fillAdjacentFormulas;
        public bool FillAdjacentFormulas
        {
            get { return fillAdjacentFormulas; }
            set { fillAdjacentFormulas = value; }
        }

        private bool preserveFormatting;
        public bool PreserveFormatting
        {
            get { return preserveFormatting; }
            set { preserveFormatting = value; }
        }

        Int32 firstColumnToSortBy;
        public Int32 FirstColumnToSortBy
        {
            get { return firstColumnToSortBy; }
            set { firstColumnToSortBy = value; }
        }

        Int32 secondColumnToSortBy;
        public Int32 SecondColumnToSortBy
        {
            get { return secondColumnToSortBy; }
            set { secondColumnToSortBy = value; }
        }

        Int32 thirdColumnToSortBy;
        public Int32 ThirdColumnToSortBy
        {
            get { return thirdColumnToSortBy; }
            set { thirdColumnToSortBy = value; }
        }

        IQueryTableParam parameters;
        public QueryTableParam()
        {
        }
    }
}
