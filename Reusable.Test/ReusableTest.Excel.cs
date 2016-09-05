using System;
using Microsoft.VisualStudio.TestTools.UnitTesting;
using IExcel = Microsoft.Office.Interop.Excel;
using Reusable.Excel.PivotTables;

namespace Reusable.Test
{
    [TestClass]
    public class Excel
    {
        private IExcel.Application AppHelper()
        {
            IExcel.Application app = new IExcel.Application();
            app.Visible = true;

            IExcel.Workbook wkb = app.Workbooks.Add(Type.Missing);
            wkb.Worksheets.Add(Type.Missing, Type.Missing, 1, Type.Missing);

            return app;
        }


        private IExcel.PivotCache PivotCacheHelper(IExcel.Worksheet wks)
        {
            IExcel.Workbook wkb = wks.Parent as IExcel.Workbook;
            IExcel.PivotCaches pvcs = wkb.PivotCaches();
            IExcel.PivotCache pvc = pvcs.Create(IExcel.XlPivotTableSourceType.xlExternal);
            pvc.Connection = @"OLEDB;Provider=SQLOLEDB.1;Data Source=sql2k802.discountasp.net;Initial Catalog=SQL2008_600471_codedotnet;User Id=SQL2008_600471_codedotnet_user;Password=pauline;";
            pvc.CommandText = "SELECT * FROM [SQL2008_600471_codedotnet].[dbo].[Z_tblMonth]";
            pvc.CommandType = IExcel.XlCmdType.xlCmdSql;

            pvc.Refresh();

            return pvc;
        }

        [TestMethod]
        public void TestPivotTableCreate()
        {
            IExcel.Application app = AppHelper();
            IExcel.Workbook wkb = app.Workbooks[1];
            IExcel.Worksheet wks = wkb.Worksheets[1] as IExcel.Worksheet;

            IExcel.PivotCache pvc = PivotCacheHelper(wks);

            IExcel.Range pivotTarget = wks.get_Range("T1");

            IExcel.PivotTable pvt = null;
            using (IPivotTable pivotbuilder = new PivotTable(new PivotTableParam()
            {
                ColumnField = "",
                ColumnGrand = false,
                DataField = "",
                Function = IExcel.XlConsolidationFunction.xlAverage,
                PageField = string.Empty,
                PivotCache = pvc,
                RowField = "",
                RowSortOrder = IExcel.XlSortOrder.xlDescending,
                RowGrand = false,
                Target = pivotTarget.get_Offset(0, 0)
            }))
            {
                pivotbuilder.Build();
                pvt = pivotbuilder.Pivot;
            };

            Assert.IsNotNull(pvt);
        }

        [TestMethod]
        public void CreatePivotCacheTest()
        {
            IExcel.Application app = AppHelper();
            IExcel.Workbook wkb = app.Workbooks[1];
            IExcel.Worksheet wks = wkb.Worksheets[1] as IExcel.Worksheet;

            IExcel.PivotCache pvc = PivotCacheHelper(wks);

            Assert.IsNotNull(pvc);

            Assert.IsTrue(pvc.RecordCount > 0);

            wkb.Close(false, Type.Missing, Type.Missing);
            app.Quit();
        }
        
        [TestMethod]
        public void RunExcelTest()
        {
            IExcel.Application app = AppHelper();
            IExcel.Workbook wkb = app.Workbooks[1];
            IExcel.Worksheet wks = wkb.Worksheets[1] as IExcel.Worksheet;

            Assert.IsNotNull(wks);

            wkb.Close(false, Type.Missing, Type.Missing);
            app.Quit();
        }

    }
}
