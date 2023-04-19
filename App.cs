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
            PushButtonData butonData2 = new PushButtonData("Area\ncoefficients", "Area\ncoefficients", assembName, "AreaCalculations.AreaCalcs.AreaCoefficients");
            butonData2.LargeImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            areaCalcPanel.AddItem(butonData2);
            areaCalcPanel.AddSeparator();

            // add tooltip
            butonData2.ToolTip = "This is a tooltip";
            butonData2.ToolTipImage = new BitmapImage(new Uri(path + @"\iconPlot.png"));

            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Failed;
        }
    }
}