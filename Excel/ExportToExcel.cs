using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;

namespace AreaCalculations
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class ExportToExcel : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                // reminder to use the previous commands prior to starting this one
                TaskDialog dialog = new TaskDialog("Напомняне");
                dialog.MainInstruction = "Моля, преди да пуснете инструмента, се уверете че имате успешно и коректно изчислени Plot Parameters, Area Coefficients и Area Parameters!";
                dialog.Show();

                // define a ProjectInfo Updater object
                ProjInfoUpdater projInfo = new ProjInfoUpdater(doc.ProjectInformation, doc);

                // check if all parameters are loaded in Project Info
                if (projInfo.CheckProjectInfoParameters() != "")
                {
                    TaskDialog projInfoParametrersError = new TaskDialog("Липсващи параметри");
                    projInfoParametrersError.MainInstruction = projInfo.CheckProjectInfoParameters();
                    projInfoParametrersError.Show();
                    return Result.Failed;
                }
                // check whether Plot Type parameter is assigned correctly
                if (!projInfo.isPlotTypeCorrect)
                {
                    TaskDialog plotTypeError = new TaskDialog("Неправилно въведен Plot Type");
                    plotTypeError.MainInstruction = "За да продължите напред, моля попълнете параметър 'Plot Type' с една от четирите посочени опции: СТАНДАРТНО УПИ, ЪГЛОВО УПИ, УПИ В ДВЕ ЗОНИ, ДВЕ УПИ!";
                    plotTypeError.Show();
                    return Result.Failed;
                }

                // area dictionary instance and additional plot parameters variables
                AreaDictionary areaDict = new AreaDictionary(doc);

                string filePath = "C:\\Users\\Danail.Momchilov\\Desktop\\Test.xlsx";

                areaDict.exportToExcel(filePath);

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                return Result.Failed;
                throw e;
            }
        }
    }
}
