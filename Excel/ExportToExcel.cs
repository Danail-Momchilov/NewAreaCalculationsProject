using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using System;
using System.Windows.Forms;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Microsoft.Win32;
using System.IO;

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

                // define a ProjectInfo Updater object
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

                // TODO: develop this check further and also apply it to any other method that is currently constructing Area Dictionary
                // TODO
                // TODO
                // TaskDialog.Show("Total Number of Areas in the Dictionary", areaDict.areasCount.ToString());
                // TaskDialog.Show("Total Number of Areas that were skipped", areaDict.missingAreasCount.ToString() + "\nThey are either not placed or have their plot or group name not filled in\n" + areaDict.missingAreasData);
                // TODO
                // TODO
                // TODO

                System.Windows.Forms.OpenFileDialog openFileDialog = new System.Windows.Forms.OpenFileDialog();
                openFileDialog.Filter = "Excel Files|*.xlsx";
                openFileDialog.Multiselect = false;

                SheetNameWindow window = new SheetNameWindow();
                window.ShowDialog();

                if (openFileDialog.ShowDialog() == DialogResult.OK)
                {
                    string filePath = openFileDialog.FileName;
                    string errormessage = areaDict.exportToExcel(filePath, window.sheetName.Text);

                    if (errormessage != "")
                    {
                        TaskDialog exportError = new TaskDialog("Открити грешки");
                        exportError.MainInstruction = errormessage;
                        exportError.Show();
                        string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "warnings.txt");
                        File.WriteAllText(path, exportError.MainInstruction);
                    }
                }

                TaskDialog.Show("Export", "Експортът завърши успешно!");

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                TaskDialog exceptions = new TaskDialog("Съобщение за грешка");
                exceptions.MainInstruction = $"{e.Message}\n\n {e}\n\n {e.InnerException} \n\n {e.GetBaseException()}";
                exceptions.Show();
                return Result.Failed;
            }
        }
    }
}
