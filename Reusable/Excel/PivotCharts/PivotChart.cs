using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using Core = Microsoft.Office.Core;
using System.Reflection;

namespace Reusable.Excel.PivotCharts
{
    public class PivotChart : IPivotChart, IDisposable
    {
        private bool disposed;
        public void Dispose()
        {
            Dispose(true);
            GC.SuppressFinalize(this);
        }

        ~PivotChart()
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
            chart = null;

            disposed = true;
        }

        private IExcel.ChartObject chart = null;
        public IExcel.ChartObject Chart
        {
            get { return chart; }
        }

        private IPivotChartParam parameters;
        public PivotChart(IPivotChartParam param)
        {
            parameters = param;        
        }

        public void Build()
        {
            try
            {
                CreatePivotChart();
                FormatCharts();
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }
        }

        public void Update(IExcel.ChartObject cht)
        {
            try
            {
                chart = cht;
                FormatCharts();
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }
        }
        
        private void CreatePivotChart()
        {
            try
            {
                IExcel.ChartObjects charts = parameters.TargetSheet.ChartObjects();
                chart = charts.Add(parameters.Left, parameters.Top, parameters.Width, parameters.Height);
                chart.Chart.ChartType = parameters.ChartType;
                chart.Chart.SetSourceData(parameters.PivotTable.DataBodyRange);
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }
        }

        private void FormatCharts()
        {
            try
            {
                if (parameters.HasTitle == true)
                {
                    chart.Chart.HasTitle = true;
                    chart.Chart.ChartTitle.Caption = parameters.Title;
                }

                chart.Chart.Legend.Left = parameters.LegendLeft;
                chart.Chart.Legend.Top = parameters.LegendTop;

                chart.Chart.ChartTitle.Top = parameters.ChartTitleTop;
                chart.Chart.ChartTitle.Left = parameters.ChartTitleLeft;

                chart.Chart.PlotArea.Top = parameters.PlotAreaTop;
                chart.Chart.PlotArea.Left = parameters.PlotAreaLeft;
                chart.Chart.PlotArea.Height = parameters.PlotAreaHeight;
                chart.Chart.PlotArea.Width = parameters.PlotAreaWidth;

                if (parameters.TickLabelSpacing == 0)
                {
                    int countBound = 0;
                    foreach (IExcel.Series ser in chart.Chart.SeriesCollection())
                    {
                        if (ser.Points().Count > countBound)
                        {
                            countBound = ser.Points().Count;
                        }
                    }

                    double optimalTickSpacing = (double)Math.Ceiling((double)(countBound / 2));
                    chart.Chart.Axes(IExcel.XlAxisType.xlCategory).TickLabelSpacing = optimalTickSpacing;
                }
                else
                {
                    chart.Chart.Axes(IExcel.XlAxisType.xlCategory).TickLabelSpacing = parameters.TickLabelSpacing;
                }

                //does not work
                chart.Chart.Axes(IExcel.XlAxisType.xlCategory).TickLabels.NumberFormat = parameters.TickLabelsNumberFormat;

                //increase legend size to accomodate up to 10 items cleanly
                //chart.Chart.Legend.Height = chart.Chart.Legend.Height * 1.3;
                chart.Chart.Legend.Font.Size = 8;

                //sets lower bound, if more features needed, added to interface for parameters
                if (parameters.ChartLowerBound != null)
                {
                    chart.Chart.Axes(IExcel.XlAxisType.xlValue).MinimumScale = parameters.ChartLowerBound;
                    chart.Chart.Axes(IExcel.XlAxisType.xlValue).CrossesAt = parameters.ChartLowerBound;
                }
                //sets lower bound, if more features needed, added to interface for parameters
                if (parameters.ChartUpperBound != null)
                {
                    chart.Chart.Axes(IExcel.XlAxisType.xlValue).MaximumScale = parameters.ChartUpperBound;
                }
            }
            catch (Exception ex)
            {
                String errMessage = String.Format("{0}: {1}", MethodBase.GetCurrentMethod().Name, ex.Message);
                throw new Exception(errMessage, ex.InnerException);
            }
        }
    }
}
