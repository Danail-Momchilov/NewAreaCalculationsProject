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

                List<double> plotAreas = new List<double>();
                List<string> plotNames = new List<string>();
                List<string> areaLevels = new List<string>();
                List<double> kint = new List<double>();
                List<double> density = new List<double>();

                List<double> buildArea = new List<double>();
                buildArea.Add(0);
                buildArea.Add(0);

                List<double> totalBuildArea = new List<double>();
                totalBuildArea.Add(0);
                totalBuildArea.Add(0);

                double areaConvert = 10.763914692;

                Transaction T = new Transaction(doc, "Update Project Info");

                string teststring = "Проектните параметри бяха обновени успешно!\n";

                TaskDialog errors = new TaskDialog("Ужас, смрад, безобразие");
                bool errorsExist = false;

                // determine plot type and calculate general parameters
                switch (projInfo.LookupParameter("Plot Type").AsString())
                {
                    case "СТАНДАРТНО УПИ":
                        // achieved area calculations
                        foreach (Area area in allAreas)
                        {
                            if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                                buildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                            else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                                totalBuildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                        teststring += "Тип на УПИ: Стандартно\n";
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                        density.Add(Math.Round(buildArea[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea[0] / plotAreas[0], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area").Set(buildArea[0]);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea[0]);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "ЪГЛОВО УПИ":
                        // achieved area calculations
                        foreach (Area area in allAreas)
                        {
                            if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                                buildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                            else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                                totalBuildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                        teststring += "Тип на УПИ: Ъглово\n";
                        plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                        plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                        density.Add(Math.Round(buildArea[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea[0] / plotAreas[0], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area").Set(buildArea[0]);
                        projInfo.LookupParameter("Required Build up Area").Set(buildArea[0]);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea[0]);
                        projInfo.LookupParameter("Required Gross External Area").Set(totalBuildArea[0]);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Required Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        projInfo.LookupParameter("Required Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "УПИ В ДВЕ ЗОНИ":
                        // achieved area calculations
                        foreach (Area area in allAreas)
                        {
                            if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                                buildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                            else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                                totalBuildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                        teststring += "Тип на УПИ: Един имот в две устройствени зони\n";
                        double plotAr = Math.Round((projInfo.LookupParameter("Zone Area 1st").AsDouble() / areaConvert) + (projInfo.LookupParameter("Zone Area 2nd").AsDouble() / areaConvert), 2);
                        plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                        plotAreas.Add(plotAr);
                        density.Add(Math.Round(buildArea[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea[0] / plotAreas[0], 2));
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
                        projInfo.LookupParameter("Achieved Built up Area").Set(buildArea[0]);
                        projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea[0]);
                        projInfo.LookupParameter("Achieved Area Intensity").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density").Set(density[0]);
                        T.Commit();
                        break;

                    case "ДВЕ УПИ":
                        teststring += "Тип на УПИ: Две отделни УПИ\n";
                        plotNames.Add(projInfo.LookupParameter("Plot Number 1st").AsString());
                        plotNames.Add(projInfo.LookupParameter("Plot Number 2nd").AsString());
                        foreach (Area area in allAreas)
                        {
                            if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                            {
                                if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                                    buildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                                else if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                                    buildArea[1] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                                else
                                    errors.MainInstruction += "Открита е партерна или полуподземна Area, която не е сътонесена към нито един от двата въведени имота!\n";
                            }
                            else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                            {
                                if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                                    totalBuildArea[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                                else if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                                    totalBuildArea[1] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                                else
                                {
                                    errors.MainInstruction += "Открита е надземна Area, която не е сътонесена към нито един от двата въведени имота!\n";
                                    errorsExist = true;
                                }                                    
                            }
                        }                        
                        plotAreas.Add(Math.Round(projInfo.LookupParameter("Plot Area 1st").AsDouble() / areaConvert, 2));     
                        density.Add(Math.Round(buildArea[0] / plotAreas[0], 2));
                        kint.Add(Math.Round(totalBuildArea[0] / plotAreas[0], 2));
                        plotAreas.Add(Math.Round(projInfo.LookupParameter("Plot Area 2nd").AsDouble() / areaConvert, 2));                  
                        density.Add(Math.Round(buildArea[1] / plotAreas[1], 2));
                        kint.Add(Math.Round(totalBuildArea[1] / plotAreas[1], 2));
                        T.Start();
                        projInfo.LookupParameter("Achieved Built up Area 1st").Set(buildArea[0]);
                        projInfo.LookupParameter("Achieved Gross External Area 1st").Set(totalBuildArea[0]);
                        projInfo.LookupParameter("Achieved Area Intensity 1st").Set(kint[0]);
                        projInfo.LookupParameter("Achieved Built up Density 1st").Set(density[0]);
                        projInfo.LookupParameter("Achieved Built up Area 2nd").Set(buildArea[1]);
                        projInfo.LookupParameter("Achieved Gross External Area 2nd").Set(totalBuildArea[1]);
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
                        teststring += $"Постигнато ЗП = {buildArea}\n";
                        teststring += $"Постигната плътност = {density[0]}\n";
                        teststring += $"Постигнато РЗП = {totalBuildArea}\n";
                        teststring += $"Постигнат КИНТ = {kint[0]}\n";
                    }
                    else
                    {
                        teststring += $"Постигнато ЗП за имот 1 = {buildArea[0]}\n";
                        teststring += $"Постигната плътност за имот 1 = {density[0]}\n";
                        teststring += $"Постигнато РЗП за имот 1 = {totalBuildArea[0]}\n";
                        teststring += $"Постигнат КИНТ за имот 1 = {kint[0]}\n";
                        teststring += $"Постигнато ЗП за имот 2 = {buildArea}\n";
                        teststring += $"Постигната плътност за имот 2 = {density[1]}\n";
                        teststring += $"Постигнато РЗП за имот 2 = {totalBuildArea}\n";
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