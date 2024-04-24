using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace AreaCalculations
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class AreaCoefficients : IExternalCommand
    {        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                AreaCoefficientsWindow window = new AreaCoefficientsWindow();

                window.ShowDialog();




                string totalTestString = "";
                totalTestString += window.areaCoefficients["ACOR"];
                totalTestString += window.areaCoefficients["ACLE"];
                totalTestString += window.areaCoefficients["ACLO"];
                totalTestString += window.areaCoefficients["ACHE"];
                totalTestString += window.areaCoefficients["ACRO"];
                totalTestString += window.areaCoefficients["ACSP"];
                totalTestString += window.areaCoefficients["ACZO"];
                totalTestString += window.areaCoefficients["ACCO"];
                totalTestString += window.areaCoefficients["ACCO"];
                totalTestString += window.areaCoefficients["ACSTS"];
                totalTestString += window.areaCoefficients["ACST"];
                totalTestString += window.areaCoefficients["ACGAP"];
                totalTestString += window.areaCoefficients["ACGA"];
                totalTestString += window.overrideBool.ToString();

                TaskDialog test = new TaskDialog("PENIS");
                test.MainInstruction = totalTestString;
                test.Show();





                AreaCollection areaUpdater = new AreaCollection(doc);

                int count;

                if (window.overrideBool)
                    count = areaUpdater.updateAreaCoefficientsOverride(window.areaCoefficients);
                else
                    count = areaUpdater.updateAreaCoefficients(window.areaCoefficients);

                TaskDialog report = new TaskDialog("Report");
                if (count > 0)
                    report.MainInstruction = $"Успешно бяха обновени коефициентите на {count} 'Area' обекти! За всички останали е обновен само коефициент 'A Coefficient Multiplied' :*";
                else
                    report.MainInstruction = "Не са открити обекти от тип 'Area' с непопълнени параметри за коефициентите. Обновен е единствено 'A Coefficient Multiplied'";
                report.Show();

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                TaskDialog exceptions = new TaskDialog("Съобщение за грешка");
                exceptions.MainInstruction = $"{e.Message}\n\n {e.ToString()}\n\n {e.InnerException} \n\n {e.GetBaseException()}";
                exceptions.Show();
                return Result.Failed;
            }
        }
    }
}
