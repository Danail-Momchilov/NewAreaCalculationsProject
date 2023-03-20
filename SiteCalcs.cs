using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace AreaCalculations
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class SiteCalcs : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                // get all areas and project info
                FilteredElementCollector allAreas = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();
                ProjectInfo projInfo = doc.ProjectInformation;

                // define a ProjectInfo Updater object
                ProjInfoUpdater ProjInfo = new ProjInfoUpdater(projInfo, doc);

                // get plot numbers from project and define a list for all the plot areas
                List<double> plotAreas = new List<double>();
                List<string> plotNames = new List<string>();

                if (projInfo.LookupParameter("Plot Type").AsString() == "ДВЕ УПИ")
                {
                    plotNames.Add(projInfo.LookupParameter("Plot Number 1st").AsString());
                    plotNames.Add(projInfo.LookupParameter("Plot Number 2nd").AsString());
                }
                else
                {
                    plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                }

                // area calculation instance and additional plot parameters variables
                AreaCollection areaCalcs = new AreaCollection(allAreas, plotNames);
                List<double> kint = new List<double>();
                List<double> density = new List<double>();

                // area conversion variable
                double areaConvert = 10.763914692;

                // define output report string
                OutputReport output = new OutputReport();

                TaskDialog errors = new TaskDialog("Ужас, смрад, безобразие");

                
                // determine plot type and calculate general parameters
                switch (projInfo.LookupParameter("Plot Type").AsString())
                {
                    case "СТАНДАРТНО УПИ":
                        output.addString("Тип на УПИ: Стандартно\n");
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
                        ProjInfo.SetAchievedStandard(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0]);
                        break;

                    case "ЪГЛОВО УПИ":
                        output.addString("Тип на УПИ: Ъглово\n");
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
                        ProjInfo.SetAchievedStandard(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0]);
                        ProjInfo.SetRequired(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0]);
                        break;

                    case "УПИ В ДВЕ ЗОНИ":
                        output.addString("Тип на УПИ: Един имот в две устройствени зони\n");
                        output.addString("Отделните параметри 1st и 2nd бяха сумирани\n");
                        double plotAr = Math.Round((projInfo.LookupParameter("Zone Area 1st").AsDouble() / areaConvert) + (projInfo.LookupParameter("Zone Area 2nd").AsDouble() / areaConvert), 2);
                        plotAreas.Add(plotAr);
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
                        ProjInfo.SetAllTwoZones(plotAr, areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0]);
                        break;

                    case "ДВЕ УПИ":
                        output.addString("Тип на УПИ: Две отделни УПИ\n");               
                        plotAreas.Add(Math.Round(projInfo.LookupParameter("Plot Area 1st").AsDouble() / areaConvert, 2));
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
                        plotAreas.Add(Math.Round(projInfo.LookupParameter("Plot Area 2nd").AsDouble() / areaConvert, 2));
                        density.Add(Math.Round(areaCalcs.build[1] / plotAreas[1], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[1] / plotAreas[1], 2));
                        ProjInfo.SetAchievedTwoPlots(areaCalcs.build[0], areaCalcs.totalBuild[0], kint[0], density[0], areaCalcs.build[1], areaCalcs.totalBuild[1], kint[1], density[1]);
                        break;
                }

                // output report
                output.updateFinalOutput(plotAreas, plotNames, areaCalcs.build, density, areaCalcs.totalBuild, kint);

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