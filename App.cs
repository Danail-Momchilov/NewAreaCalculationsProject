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
            String tabName = "MIPA";
            application.CreateRibbonTab(tabName);

            // creatte a panel
            RibbonPanel areaCalcPanel = application.CreateRibbonPanel(tabName, "Area Calculations");



            // 
            PushButtonData smallButton1 = new PushButtonData("About", "A", assembName, "AreaCalculations.AboutInfo");
            smallButton1.Image = new BitmapImage(new Uri(path + @"\info.png"));

            PushButtonData smallButton2 = new PushButtonData("Version", "V", assembName, "AreaCalculations.VersionInfo");
            smallButton2.Image = new BitmapImage(new Uri(path + @"\version.png"));

            PushButtonData smallButton3 = new PushButtonData("Resources", "R", assembName, "AreaCalculations.ResourceInfo");
            smallButton3.Image = new BitmapImage(new Uri(path + @"\education.png"));

            IList<RibbonItem> stackedButtons = areaCalcPanel.AddStackedItems(smallButton1, smallButton2, smallButton3);



            // create PushButon1
            PushButtonData butonData1 = new PushButtonData("Plot\nparameters", "Plot\nparameters", assembName, "AreaCalculations.SiteCalcs");
            butonData1.LargeImage = new BitmapImage(new Uri(path + @"\plotIcon.png"));

            areaCalcPanel.AddItem(butonData1);
            areaCalcPanel.AddSeparator();

            butonData1.ToolTip = "Натиснете този бутон, за да изчислите всички параметри за имота";
            butonData1.ToolTipImage = new BitmapImage(new Uri(path + @"\plotIcon.png"));



            // create PushButon 2
            PushButtonData butonData2 = new PushButtonData("Area\ncoefficients", "Area\ncoefficients", assembName, "AreaCalculations.AreaCoefficients");
            butonData2.LargeImage = new BitmapImage(new Uri(path + @"\areacIcon.png"));

            areaCalcPanel.AddItem(butonData2);
            areaCalcPanel.AddSeparator();

            butonData2.ToolTip = "Натиснете този бутон, за да попълните автоматично всички Area коефициенти";
            butonData2.ToolTipImage = new BitmapImage(new Uri(path + @"\areacIcon.png"));



            // create PushButon 3
            PushButtonData butonData3 = new PushButtonData("Area\ncalculations", "Area\ncalculations", assembName, "AreaCalculations.CalculateAreaParameters");
            butonData3.LargeImage = new BitmapImage(new Uri(path + @"\areaIcon.png"));

            areaCalcPanel.AddItem(butonData3);
            areaCalcPanel.AddSeparator();

            butonData3.ToolTip = "Натиснете този бутон, за да изчислите параметрите към всяка една Area";
            butonData3.ToolTipImage = new BitmapImage(new Uri(path + @"\areaIcon.png"));



            // create PushButon 4
            PushButtonData butonData4 = new PushButtonData("Export to\nExcel", "Export to\nExcel", assembName, "AreaCalculations.ExportToExcel");
            butonData4.LargeImage = new BitmapImage(new Uri(path + @"\excelIcon.png"));

            areaCalcPanel.AddItem(butonData4);
            areaCalcPanel.AddSeparator();

            butonData4.ToolTip = "Натиснете този бутон, за да експортнете всички обекти Area директно в Excel";
            butonData4.ToolTipImage = new BitmapImage(new Uri(path + @"\excelIcon.png"));



            return Result.Succeeded;
        }
        public Result OnShutdown(UIControlledApplication application)
        {
            return Result.Failed;
        }
    }
}