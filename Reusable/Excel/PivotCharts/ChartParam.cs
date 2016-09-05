using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;

namespace Reusable.Excel.PivotCharts
{
    public class PivotChartParam : IPivotChartParam, IDisposable
    {
        private bool disposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~PivotChartParam()
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
            this.PivotTable = null;
            this.TargetSheet = null;

            disposed = true;
        }

        public CustomLayout CustomLayoutType { get; set; }
        public Double? ChartLowerBound {get; set;}
        public Double? ChartUpperBound {get; set;}
        public Boolean HasTitle {get; set;}
        public IExcel.PivotTable PivotTable {get; set;}
        public IExcel.Worksheet TargetSheet {get; set;}
        public Double Left {get; set;}
        public Double Top {get; set;}
        public IExcel.XlChartType ChartType {get; set;}
        public String Title {get; set;}
        public double Height {get; set;}
        public double Width {get; set;}
        public Double TickLabelSpacing {get; set;}
        public String TickLabelsNumberFormat { get; set; }
        public Double ChartTitleTop {get; set;}
        public Double ChartTitleLeft {get; set;}
        public Double PlotAreaTop {get; set;}
        public Double PlotAreaLeft {get; set;}
        public Double PlotAreaHeight {get; set;}
        public Double PlotAreaWidth {get; set;}
        public Double LegendLeft { get; set; }
        public Double LegendTop { get; set; }

        public PivotChartParam()
        {
        }
    }
}
