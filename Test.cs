using System;
using System.Collections.Generic;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace AreaCalculations
{
    [TransactionAttribute(TransactionMode.Manual)]
    public class Test : IExternalCommand
    {
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            UIDocument uidoc = commandData.Application.ActiveUIDocument;
            Document doc = uidoc.Document;

            try
            {
                FilteredElementCollector PropertyLines = new FilteredElementCollector(doc).OfClass(typeof(PropertyLine));
                FilteredElementCollector allAreas = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();
                ProjectInfo projInfo = doc.ProjectInformation;

                List<Area> areasFP01 = new List<Area>();
                List<double> plotAreas = new List<double>();
                List<string> plotNames = new List<string>();
                List<string> areaLevels = new List<string>();

                double areaConvert = 10.763914692;
                double buildArea = 0;
                double totalBuildArea = 0;
                double density;
                double kint;

                foreach (Area area in allAreas)
                {
                    if (area.LookupParameter("Level").AsValueString() == "AR-FP-01" && area.LookupParameter("A Instance Area Group").AsString() != "ЗЕМЯ")
                    {
                        areasFP01.Add(area);
                        buildArea += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                    else if (!area.LookupParameter("Level").AsValueString().Contains("BP") && area.LookupParameter("A Instance Area Group").AsString() != "ЗЕМЯ")
                    {
                        totalBuildArea += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                }

                foreach (PropertyLine property in PropertyLines)
                {
                    plotAreas.Add(Math.Round(property.GetParameters("Area")[0].AsDouble()/areaConvert, 2));
                    plotNames.Add(property.GetParameters("Name")[0].AsString());
                }

                density = Math.Round(buildArea / plotAreas[0], 2);
                kint = Math.Round(totalBuildArea / plotAreas[0], 2);





                
                // info test
                Transaction T = new Transaction(doc, "Update Project Info");
                T.Start();

                projInfo.LookupParameter("Plot Number").Set(plotNames[0]);
                projInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
                projInfo.LookupParameter("Achieved Gross External Area").Set(totalBuildArea);
                projInfo.LookupParameter("Achieved Area Intensity").Set(kint);
                projInfo.LookupParameter("Achieved Built up Density").Set(density);

                T.Commit();






                // test
                string teststring = "";
                for (int i = 0; i < plotAreas.Count; i++)
                {
                    teststring = teststring + $"PLot {i} area: " + plotAreas[i].ToString() + "\n" + $"PLot {i} name: " + plotNames[i] + "\n";
                }

                teststring += $"Number of Areas on the ground floor = {areasFP01.Count}\n";
                teststring += $"BuildArea = {buildArea}\n";
                teststring += $"Density = {density}\n";
                teststring += $"Total number of Areas in the model = {allAreas.GetElementCount()}\n";
                teststring += $"TBA = {totalBuildArea}\n";
                teststring += $"Kint = {kint}\n";

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