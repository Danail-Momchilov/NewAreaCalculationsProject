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
                FilteredElementCollector allAreas = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();
                ProjectInfo projInfo = doc.ProjectInformation;

                List<Area> areasZP = new List<Area>();
                List<double> plotAreas = new List<double>();
                List<string> plotNames = new List<string>();
                List<string> areaLevels = new List<string>();
                List<double> kint = new List<double>();
                List<double> density = new List<double>();

                double areaConvert = 10.763914692;
                double buildArea = 0;
                double totalBuildArea = 0;

                Transaction T = new Transaction(doc, "Update Project Info");

                string teststring = "Проектните параметри бяха обновени успешно!\n";

                // achieved area calculations
                foreach (Area area in allAreas)
                {
                    if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                    {
                        areasZP.Add(area);
                        buildArea += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                    else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                    {
                        totalBuildArea += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                }
                
                // determine plot type and calculate general parameters
                switch (projInfo.LookupParameter("Plot Type").AsString())
                {
                    case "СТАНДАРТНО УПИ":
                        teststring += "Тип на УПИ: Стандартно\n";
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                        density.Add(Math.Round(buildArea / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea / plotAreas[0], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "ЪГЛОВО УПИ":
                        teststring += "Тип на УПИ: Ъглово\n";
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                        density.Add(Math.Round(buildArea / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea / plotAreas[0], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
                        projInfo.LookupParameter("Required Build up Area").Set(buildArea);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea);
                        projInfo.LookupParameter("Required Gross External Area").Set(totalBuildArea);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Required Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        projInfo.LookupParameter("Required Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "УПИ В ДВЕ ЗОНИ":
                        teststring += "Тип на УПИ: Един имот в две устройствени зони\n";
                        double plotAr = Math.Round( (projInfo.LookupParameter("Zone Area 1st").AsDouble() / areaConvert) + (projInfo.LookupParameter("Zone Area 2nd").AsDouble() / areaConvert), 2 );
                        plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                        plotAreas.Add(plotAr);
                        density.Add(Math.Round(buildArea / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea / plotAreas[0], 2));
                        T.Start();
                        projInfo.LookupParameter("Plot Area").Set(plotAr);
                        projInfo.LookupParameter("Required Built up Density")
                            .Set(projInfo.LookupParameter("Required Built up Density 1st").AsDouble() + projInfo.LookupParameter("Required Built up Density 2nd").AsDouble());
                        projInfo.LookupParameter("Required Built up Area")
                            .Set(projInfo.LookupParameter("Required Built up Area 1st").AsDouble() + projInfo.LookupParameter("Required Built up Area 2nd").AsDouble());
                        projInfo.LookupParameter("Required Area Intensity")
                            .Set(projInfo.LookupParameter("Required Area Intensity 1st").AsDouble() + projInfo.LookupParameter("Required Area Intensity 2nd").AsDouble());
                        projInfo.LookupParameter("Required Gross External Area")
                            .Set(projInfo.LookupParameter("Required Gross External Area 1st").AsDouble() + projInfo.LookupParameter("Required Gross External Area 2nd").AsDouble());
                        projInfo.LookupParameter("Required Green Area Percentage")
                            .Set(projInfo.LookupParameter("Required Green Area Percentage 1st").AsDouble() + projInfo.LookupParameter("Required Green Area Percentage 2nd").AsDouble());
                        projInfo.LookupParameter("Required Green Area")
                            .Set(projInfo.LookupParameter("Required Green Area 1st").AsDouble() + projInfo.LookupParameter("Required Green Area 2nd").AsDouble());
                        teststring += "Отделните параметри 1st и 2nd бяха сумирани\n";
                        projInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "ДВЕ УПИ В ЕДНА ЗОНА":
                        teststring += "Тип на УПИ: Две УПИ в една зона\n";
                        // just a test underneath
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                        density.Add(Math.Round(buildArea / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea / plotAreas[0], 2));
                        // to be continued
                        break;

                    default:
                        TaskDialog error = new TaskDialog("Възникнала грешка");
                        error.MainInstruction = "Моля, попълнете параметъра 'Plot Type' с една от следните опции: СТАНДАРТНО УПИ, ЪГЛОВО УПИ, УПИ В ДВЕ ЗОНИ, ДВЕ УПИ В ЕДНА ЗОНА";
                        error.Show();
                        Environment.Exit(0);
                        break;
                }

                // output report
                for (int i = 0; i < plotAreas.Count; i++)
                {
                    teststring = teststring + $"Площ на имот {i}: " + plotAreas[i].ToString() + "\n" + $"Име на имот {i}: " + plotNames[i] + "\n";
                }

                teststring += $"Постигнато ЗП = {buildArea}\n";
                teststring += $"Постигната плътност = {density[0]}\n";
                teststring += $"Постигнато РЗП = {totalBuildArea}\n";
                teststring += $"Постигнат КИНТ = {kint[0]}\n";

                TaskDialog testDialog = new TaskDialog("Report");
                testDialog.MainInstruction = teststring;
                testDialog.Show();

                return Result.Succeeded;
            }

            catch (Exception e)
            {
                return Result.Failed;
            }
        }
    }
}