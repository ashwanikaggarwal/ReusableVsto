using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using IExcel = Microsoft.Office.Interop.Excel;
using System.Reflection;
        
namespace Reusable.Excel.PivotCharts
{
    public enum CustomLayout
    { 
        None = 0,
        CusipChart = 1    
    }

    public interface IPivotChartParam : IDisposable
    {
        IExcel.PivotTable PivotTable {get;set;}
        IExcel.Worksheet TargetSheet { get; set; }

        IExcel.XlChartType ChartType { get; set; }

        CustomLayout CustomLayoutType { get; set; }

        Boolean HasTitle { get; set; }
        Double TickLabelSpacing { get; set; }
        String TickLabelsNumberFormat { get; set; }
        Double ChartTitleTop { get; set; }
        Double ChartTitleLeft { get; set; }

        Double Left { get; set; }
        Double Top { get; set; }
        String Title { get; set; }
        Double Height { get; set; }
        Double Width { get; set; }
        
        Double PlotAreaTop { get; set; }
        Double PlotAreaLeft { get; set; }
        Double PlotAreaHeight { get; set; }
        Double PlotAreaWidth { get; set; }

        Double? ChartLowerBound { get; set; }
        Double? ChartUpperBound { get; set; }

        Double LegendLeft { get; set; }
        Double LegendTop { get; set; }
    }
}
