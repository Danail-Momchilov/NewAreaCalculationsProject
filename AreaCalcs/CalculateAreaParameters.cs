using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.IO;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AreaCalculations
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class CalculateAreaParameters : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // reminder to use the previous commands prior to starting this one
                TaskDialog dialog = new TaskDialog("Напомняне");
                dialog.MainInstruction = "Моля, преди да пуснете инструмента, се уверете че имате успешно " +
                    "и коректно изчислени Plot Parameters и Area Coefficients!";
                dialog.Show();

                ProjInfoUpdater projInfo = new ProjInfoUpdater(doc);

                // check if all parameters are loaded in Project Info
                if (projInfo.CheckProjectInfoParameters() != "")
                {
                    TaskDialog projInfoParametrersError = new TaskDialog("Липсващи параметри");
                    projInfoParametrersError.MainInstruction = projInfo.CheckProjectInfoParameters();
                    projInfoParametrersError.Show();
                    string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "warnings.txt");
                    File.WriteAllText(path, projInfoParametrersError.MainInstruction);
                    return Result.Failed;
                }

                projInfo = new ProjInfoUpdater(doc.ProjectInformation, doc);

                // check whether Plot Type parameter is assigned correctly
                if (!projInfo.isPlotTypeCorrect)
                {
                    TaskDialog plotTypeError = new TaskDialog("Неправилно въведен Plot Type");
                    plotTypeError.MainInstruction = "За да продължите напред, моля попълнете параметър 'Plot Type' с една от четирите посочени опции: СТАНДАРТНО УПИ, ЪГЛОВО УПИ, УПИ В ДВЕ ЗОНИ, ДВЕ УПИ!";
                    plotTypeError.Show();
                    string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "warnings.txt");
                    File.WriteAllText(path, plotTypeError.MainInstruction);
                    return Result.Failed;
                }

                // area dictionary instance and additional plot parameters variables
                AreaDictionary areaDict = new AreaDictionary(doc);
                // calculate C1/C2 coefficients
                areaDict.calculateC1C2();
                // calculate common area percentages
                areaDict.calculateCommonAreaPerc();
                // calculate A Instance Common Area Parameter
                areaDict.calculateCommonArea();
                // calculate A Instance Building Permit %
                areaDict.calculateBuildingPercentPermit();
                // calculate A Instance RLP Area %
                areaDict.calculateRlpAreaPercent();
                // calculate A Instance RLP Area
                areaDict.calculateRlpArea();
                // calculate A Instance Common Area Special
                areaDict.calculateSpecialCommonAreas();                
                // after calculations are complete, redistribute the final surplus left
                areaDict.redistributeSurplus();
                // calculate A Instance Total Area
                areaDict.calculateTotalArea();
                // calculate A Instance Common Area Percentage
                string errrorMessage = areaDict.calculateInstancePropertyCommonAreaPercentage();
                errrorMessage += areaDict.errorMessage;

                if (errrorMessage != "")
                {
                    TaskDialog errorReport = new TaskDialog ("Открити грешки. Вземете им обърнете внимание");
                    errorReport.MainInstruction = errrorMessage;
                    errorReport.Show();
                    string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "warnings.txt");
                    File.WriteAllText(path, errrorMessage);
                    return Result.Failed;
                }
                else
                {
                    TaskDialog dialogReport = new TaskDialog("Репорт");
                    dialogReport.MainInstruction = "Успешно бяха преизчислени параметрите на всички Areas в обекта!";
                    dialogReport.Show();
                    return Result.Succeeded;
                }                
            }
            catch (Exception e)
            {
                TaskDialog exceptions = new TaskDialog("Съобщение за грешка");
                exceptions.MainInstruction = $"{e.Message}\n\n {e.ToString()} \n\n {e.StackTrace}";
                exceptions.Show();
                return Result.Failed;
            }
        }
    }
}
