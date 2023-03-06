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

                double areaConvert = 10.763914692;
                double buildArea = 0;
                double totalBuildArea = 0;
                double density;
                double kint;


                // achieved area calculations
                foreach (Area area in allAreas)
                {
                    if (area.LookupParameter("A Instance Area Location").AsValueString() == "НАДЗЕМНО" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНО")
                    {
                        areasZP.Add(area);
                        buildArea += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                    else if (area.LookupParameter("A Instance Area Location").AsValueString() == "НАЗЕМНО" || area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНО")
                    {
                        totalBuildArea += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                }

                // determine plot area
                if (projInfo.LookupParameter("Zone Area 1st") != null && projInfo.LookupParameter("Zone Area 2nd") != null)
                {
                    double area1 = projInfo.LookupParameter("Plot Area 1st").AsDouble() / areaConvert;
                    double area2 = projInfo.LookupParameter("Plot Area 2nd").AsDouble() / areaConvert;
                    plotAreas.Add(area1 + area2);
                }
                else
                {
                    plotAreas.Add(projInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                }
                
                // calculate plot parameters
                plotNames.Add(projInfo.LookupParameter("Plot Number").AsString());
                density = Math.Round(buildArea / plotAreas[0], 2);
                kint = Math.Round(totalBuildArea / plotAreas[0], 2);







                /*
                // info test
                Transaction T = new Transaction(doc, "Update Project Info");
                T.Start();

                projInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
                projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea);
                projInfo.LookupParameter("Achieved Area Intensity").Set(kint);
                projInfo.LookupParameter("Achieved Built up Density").Set(density);

                T.Commit();
                */





                // test
                string teststring = "";

                //teststring += projInfo.LookupParameter("Plot Area").AsDouble();
                
                for (int i = 0; i < plotAreas.Count; i++)
                {
                    teststring = teststring + $"PLot {i} area: " + plotAreas[i].ToString() + "\n" + $"PLot {i} name: " + plotNames[i] + "\n";
                }
                

                teststring += $"Number of Areas included in the build area = {areasZP.Count}\n";
                teststring += $"BuildArea = {buildArea}\n";
                //teststring += $"Density = {density}\n";
                teststring += $"Total number of Areas in the model = {allAreas.GetElementCount()}\n";
                teststring += $"TBA = {totalBuildArea}\n";
                //teststring += $"Kint = {kint}\n";

                foreach (string level in areaLevels)
                {
                    teststring += "level\n";
                }

                TaskDialog testDialog = new TaskDialog("Report");
                testDialog.MainInstruction = teststring;
                testDialog.Show();
                // end of test

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                return Result.Failed;
            }
        }
    }
}