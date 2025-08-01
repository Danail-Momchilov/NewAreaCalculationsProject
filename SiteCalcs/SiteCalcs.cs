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
    public class SiteCalcs : IExternalCommand
    {
        public Result Execute (ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                ProjInfoUpdater ProjInfo = new ProjInfoUpdater(doc);

                // check if all parameters are loaded in Project Info
                if (ProjInfo.CheckProjectInfoParameters() != "")
                {
                    TaskDialog projInfoParametrersError = new TaskDialog("Липсващи параметри");
                    projInfoParametrersError.MainInstruction = ProjInfo.CheckProjectInfoParameters();
                    projInfoParametrersError.Show();
                    string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "warnings.txt");
                    File.WriteAllText(path, projInfoParametrersError.MainInstruction);
                    return Result.Failed;
                }

                ProjInfo = new ProjInfoUpdater(doc.ProjectInformation, doc);
                
                // check whether Plot Type parameter is assigned correctly
                if (!ProjInfo.isPlotTypeCorrect)
                {
                    TaskDialog plotTypeError = new TaskDialog("Неправилно въведен Plot Type");
                    plotTypeError.MainInstruction = "За да продължите напред, моля попълнете параметър 'Plot Type' " +
                        "с една от четирите посочени опции: СТАНДАРТНО УПИ, ЪГЛОВО УПИ, УПИ В ДВЕ ЗОНИ, ДВЕ УПИ!";
                    plotTypeError.Show();
                    string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "warnings.txt");
                    File.WriteAllText(path, plotTypeError.MainInstruction);
                    return Result.Failed;
                }

                // area calculation instance and additional plot parameters variables
                AreaCollection areaCalcs = new AreaCollection(doc, ProjInfo.plotNames);
                List<double> kint = new List<double>();
                List<double> density = new List<double>();

                // Greenery object definition
                Greenery greenery = new Greenery(doc, ProjInfo.plotNames, ProjInfo.plotAreas);

                // define output report string
                OutputReport output = new OutputReport();

                // check if all the information in the Areas and Project info is set correctly
                string errors = ProjInfo.CheckProjectInfo() + areaCalcs.CheckAreasParameters(ProjInfo.plotNames, doc.ProjectInformation) + greenery.errorReport;
                if (errors != "")
                {
                    TaskDialog errorReport = new TaskDialog("Открити грешки");
                    errorReport.MainInstruction = errors;
                    errorReport.Show();
                    string path = Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Desktop), "warnings.txt");
                    File.WriteAllText(path, errors);
                    return Result.Failed;
                }
                
                // determine plot type and calculate general parameters
                switch (doc.ProjectInformation.LookupParameter("Plot Type").AsString())
                {
                    case "СТАНДАРТНО УПИ":
                        output.addString("Тип на УПИ: Стандартно\n");
                        density.Add(Math.Round(areaCalcs.build[0] / ProjInfo.plotAreas[0], 2, MidpointRounding.AwayFromZero));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / ProjInfo.plotAreas[0], 2, MidpointRounding.AwayFromZero));
                        ProjInfo.SetAchievedStandard(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0], 
                            greenery.greenArea, greenery.achievedPercentage);
                        break;

                    case "ЪГЛОВО УПИ":
                        output.addString("Тип на УПИ: Ъглово\n");
                        density.Add(Math.Round(areaCalcs.build[0] / ProjInfo.plotAreas[0], 2, MidpointRounding.AwayFromZero));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / ProjInfo.plotAreas[0], 2, MidpointRounding.AwayFromZero));
                        ProjInfo.SetAchievedStandard(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0], 
                            greenery.greenArea, greenery.achievedPercentage);
                        ProjInfo.SetRequired(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0]);
                        break;

                    case "УПИ В ДВЕ ЗОНИ":
                        output.addString("Тип на УПИ: Един имот в две устройствени зони\n");
                        output.addString("Отделните параметри 1st и 2nd бяха сумирани\n");
                        density.Add(Math.Round(areaCalcs.build[0] / ProjInfo.plotAreas[0], 2, MidpointRounding.AwayFromZero));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / ProjInfo.plotAreas[0], 2, MidpointRounding.AwayFromZero));
                        ProjInfo.SetAllTwoZones(ProjInfo.plotAreas[0], areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0], 
                            greenery.greenArea, greenery.achievedPercentage);
                        break;

                    case "ДВЕ УПИ":
                        output.addString("Тип на УПИ: Две отделни УПИ\n");
                        density.Add(Math.Round(areaCalcs.build[0] / ProjInfo.plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / ProjInfo.plotAreas[0], 2));
                        density.Add(Math.Round(areaCalcs.build[1] / ProjInfo.plotAreas[1], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[1] / ProjInfo.plotAreas[1], 2));
                        ProjInfo.SetAchievedTwoPlots(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], 
                            density[0], areaCalcs.build[1], areaCalcs.totalBuild[1], kint[1], density[1], greenery.greenArea1, 
                            greenery.greenArea2, greenery.achievedPercentage1, greenery.achievedPercentage2);
                        break;
                }

                // output report
                output.updateFinalOutput (ProjInfo.plotAreas, ProjInfo.plotNames, areaCalcs.build, density, 
                    areaCalcs.totalBuild, kint, greenery.greenAreas, greenery.achievedPercentages);

                TaskDialog testDialog = new TaskDialog("Report");
                testDialog.MainInstruction = output.outputString;
                testDialog.Show();
                
                return Result.Succeeded;
            }

            catch (Exception e)
            {
                TaskDialog exceptions = new TaskDialog("Съобщение за грешка");
                exceptions.MainInstruction = $"{e.Message}\n\n {e.ToString()}";
                exceptions.Show();
                return Result.Failed;
            }
        }
    }
}