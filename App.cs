using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Reflection;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media.Imaging;

namespace AreaCalculations
{
    internal class App : IExternalApplication
    {
        public Result OnStartup(UIControlledApplication application)
        {
            String assembName = Assembly.GetExecutingAssembly().Location;
            String path = System.IO.Path.GetDirectoryName(assembName);

            // create a tab
            String tabName = "&More";
            application.CreateRibbonTab(tabName);

            // creatte a panel
            RibbonPanel areaCalcPanel = application.CreateRibbonPanel(tabName, "Area Calculations");

            // create PushButon1
            PushButtonData butonData1 = new PushButtonData("Plot\nparameters", "Plot\nparameters", assembName, "AreaCalculations.SiteCalcs");
            butonData1.LargeImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            areaCalcPanel.AddItem(butonData1);
            areaCalcPanel.AddSeparator();

            // add tooltip
            butonData1.ToolTip = "This is a tooltip";
            butonData1.ToolTipImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            // create PushButon 2
            PushButtonData butonData2 = new PushButtonData("Area\ncoefficients", "Area\ncoefficients", assembName, "AreaCalculations.AreaCoefficients");
            butonData2.LargeImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            areaCalcPanel.AddItem(butonData2);
            areaCalcPanel.AddSeparator();

            // create PushButon 3
            PushButtonData butonData3 = new PushButtonData("Area\ncalculations", "Area\ncalculations", assembName, "AreaCalculations.CalculateAreaParameters");
            butonData2.LargeImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            areaCalcPanel.AddItem(butonData3);
            areaCalcPanel.AddSeparator();

            // add tooltip
            butonData3.ToolTip = "This is a tooltip";
            butonData3.ToolTipImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            // create PushButon 4
            PushButtonData butonData4 = new PushButtonData("Export to\nExcel", "Export to\nExcel", assembName, "AreaCalculations.ExportToExcel");
            butonData4.LargeImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            areaCalcPanel.AddItem(butonData4);
            areaCalcPanel.AddSeparator();

            // add tooltip
            butonData4.ToolTip = "This is a tooltip";
            butonData4.ToolTipImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Failed;
        }
    }
}