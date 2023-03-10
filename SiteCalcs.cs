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

                Transaction T = new Transaction(doc, "Update Project Info");

                // define output report string
                string teststring = "Проектните параметри бяха обновени успешно!\n";

                TaskDialog errors = new TaskDialog("Ужас, смрад, безобразие");
                bool errorsExist = false;

                // determine plot type and calculate general parameters
                switch (projInfo.LookupParameter("Plot Type").AsString())
                {
                    case "СТАНДАРТНО УПИ":
                        teststring += "Тип на УПИ: Стандартно\n";
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area").Set(areaCalcs.build[0]);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(areaCalcs.totalBuild[0]);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "ЪГЛОВО УПИ":
                        teststring += "Тип на УПИ: Ъглово\n";
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area").Set(areaCalcs.build[0]);
                        projInfo.LookupParameter("Required Build up Area").Set(areaCalcs.build[0]);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(areaCalcs.totalBuild[0]);
                        projInfo.LookupParameter("Required Gross External Area").Set(areaCalcs.totalBuild[0]);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Required Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        projInfo.LookupParameter("Required Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "УПИ В ДВЕ ЗОНИ":
                        teststring += "Тип на УПИ: Един имот в две устройствени зони\n";
                        double plotAr = Math.Round((projInfo.LookupParameter("Zone Area 1st").AsDouble() / areaConvert) + (projInfo.LookupParameter("Zone Area 2nd").AsDouble() / areaConvert), 2);
                        plotAreas.Add(plotAr);
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
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
                        projInfo.LookupParameter("Achieved Built up Area").Set(areaCalcs.build[0]);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(areaCalcs.totalBuild[0]);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "ДВЕ УПИ":
                        teststring += "Тип на УПИ: Две отделни УПИ\n";                  
                        plotAreas.Add(Math.Round(projInfo.LookupParameter("Plot Area 1st").AsDouble() / areaConvert, 2));     
                        density.Add(Math.Round(areaCalcs.build[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[0] / plotAreas[0], 2));
                        plotAreas.Add(Math.Round(projInfo.LookupParameter("Plot Area 2nd").AsDouble() / areaConvert, 2));                  
                        density.Add(Math.Round(areaCalcs.build[1] / plotAreas[1], 2));
                        kint.Add(Math.Round(areaCalcs.totalBuild[1] / plotAreas[1], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area 1st").Set(areaCalcs.build[0]);
                        projInfo.LookupParameter("Achieved Gross External Area 1st").Set(areaCalcs.totalBuild[0]);
                        projInfo.LookupParameter("Achieved Area Intensity 1st").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density 1st").Set(density[0]);
                        projInfo.LookupParameter("Achieved Built up Area 2nd").Set(areaCalcs.build[1]);
                        projInfo.LookupParameter("Achieved Gross External Area 2nd").Set(areaCalcs.totalBuild[1]);
                        projInfo.LookupParameter("Achieved Area Intensity 2nd").Set(kint[1]);
                        projInfo.LookupParameter("Achieved Built up Density 2nd").Set(density[1]);
                        T.Commit();
                        break;

                    default:
                        errors.MainInstruction += "Моля, попълнете параметъра 'Plot Type' с една от следните опции: СТАНДАРТНО УПИ, ЪГЛОВО УПИ, УПИ В ДВЕ ЗОНИ, ДВЕ УПИ\n";
                        errorsExist = true;
                        break;
                }

                // output report
                if (errorsExist)
                    errors.Show();
                else
                {
                    for (int i = 0; i < plotAreas.Count; i++)
                    {
                        teststring = teststring + $"Площ на имот {i}: " + plotAreas[i].ToString() + "\n" + $"Име на имот {i}: " + plotNames[i] + "\n";
                    }
                    if (plotAreas.Count == 1)
                    {
                        teststring += $"Постигнато ЗП = {areaCalcs.build[0]}\n";
                        teststring += $"Постигната плътност = {density[0]}\n";
                        teststring += $"Постигнато РЗП = {areaCalcs.totalBuild[0]}\n";
                        teststring += $"Постигнат КИНТ = {kint[0]}\n";
                    }
                    else
                    {
                        teststring += $"Постигнато ЗП за имот 1 = {areaCalcs.build[0]}\n";
                        teststring += $"Постигната плътност за имот 1 = {density[0]}\n";
                        teststring += $"Постигнато РЗП за имот 1 = {areaCalcs.totalBuild[0]}\n";
                        teststring += $"Постигнат КИНТ за имот 1 = {kint[0]}\n";
                        teststring += $"Постигнато ЗП за имот 2 = {areaCalcs.build[1]}\n";
                        teststring += $"Постигната плътност за имот 2 = {density[1]}\n";
                        teststring += $"Постигнато РЗП за имот 2 = {areaCalcs.totalBuild[1]}\n";
                        teststring += $"Постигнат КИНТ за имот 2 = {kint[1]}\n";
                    }

                    TaskDialog testDialog = new TaskDialog("Report");
                    testDialog.MainInstruction = teststring;
                    testDialog.Show();
                }
                
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