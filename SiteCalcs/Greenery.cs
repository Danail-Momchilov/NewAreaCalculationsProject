using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Media;

namespace AreaCalculations
{
    internal class Greenery
    {
        public double greenArea { get; set; }
        public double greenArea1 { get; set; }
        public double greenArea2 { get; set; }
        public double achievedPercentage { get; set; }
        public double achievedPercentage1 { get; set; }
        public double achievedPercentage2 { get; set; }
        public List<double> greenAreas { get; set; } = new List<double>();
        public List<double> achievedPercentages { get; set; } = new List<double>();
        private double semiSmartRound(Element element, string parameterName)
        {
            double result = Math.Round(element.LookupParameter(parameterName).AsDouble() / areaConvert, 2, MidpointRounding.AwayFromZero);

            return result;
        }
        public string errorReport = "";
        double areaConvert = 10.7639104167096;
        double lengthConvert = 30.48;
        public Greenery(Document doc, List<string> plotNames, List<double> plotAreas)
        {
            try
            {
                FilteredElementCollector allFloors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
                FilteredElementCollector allWalls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();
                FilteredElementCollector allRailings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRailing).WhereElementIsNotElementType();

                if (plotNames.Count == 1)
                {
                    foreach (Floor floor in allFloors)
                        if (floor.FloorType.LookupParameter("Green Area").AsInteger() == 1)
                            greenArea += semiSmartRound(floor, "Area");

                    foreach (Wall wall in allWalls)
                    {
                        if (wall.WallType.LookupParameter("Green Area").AsInteger() == 1)
                        {
                            if ((wall.LookupParameter("Unconnected Height").AsDouble() * lengthConvert) <= 200)
                                greenArea += semiSmartRound(wall, "Area");
                            else
                                greenArea += semiSmartRound(wall, "Length") / 100 * 2;
                        }
                    }

                    foreach (Railing railing in allRailings)
                    {
                        ElementId railingTypeId = railing.GetTypeId();
                        ElementType railingType = doc.GetElement(railingTypeId) as ElementType;

                        if (railingType.LookupParameter("Green Area").AsInteger() == 1)
                        {
                            if (semiSmartRound(railingType, "Railing Height") <= 200)
                                greenArea += Math.Round(semiSmartRound(railing, "Length") / 100 * semiSmartRound(railingType, "Railing Height") / 100, 
                                    2, MidpointRounding.AwayFromZero);
                            else
                                greenArea += semiSmartRound(railing, "Length") / 100;
                        }
                    }

                    achievedPercentage = Math.Round(((greenArea * 100) / plotAreas[0]), 2);
                    greenAreas.Add(greenArea);
                    achievedPercentages.Add(achievedPercentage);
                }

                else if (plotNames.Count == 2)
                {
                    foreach (Floor floor in allFloors)
                    {
                        if (floor.FloorType.LookupParameter("Green Area").AsInteger() == 1)
                        {
                            if (floor.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                                greenArea1 += semiSmartRound(floor, "Area");
                            else if (floor.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                                greenArea2 += semiSmartRound(floor, "Area");
                            else
                                errorReport += $"Плоча с id: {floor.Id} има попълнен параметър A Instance Area Plot, чиято стойност не отговаря на нито едно от двете въведени имена за УПИ\n";
                        }
                    }

                    foreach (Wall wall in allWalls)
                    {
                        if (wall.WallType.LookupParameter("Green Area").AsInteger() == 1)
                        {
                            double wallArea = 0;
                            if (wall.WallType.LookupParameter("Green Area").AsInteger() == 1)
                            {
                                if (wall.LookupParameter("Unconnected Height").AsDouble() * lengthConvert <= 200)
                                    wallArea += semiSmartRound(wall, "Area");
                                else
                                    wallArea += semiSmartRound(wall, "Length") / 100 * 2;
                            }
                            if (wall.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                                greenArea1 += wallArea;
                            else if (wall.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                                greenArea2 += wallArea;
                            else
                                errorReport += $"Стена с id: {wall.Id} има попълнен параметър A Instance Area Plot, чиято стойност не отговаря на нито едно от двете въведени имена за УПИ\n";
                        }
                    }

                    foreach (Railing railing in allRailings)
                    {
                        ElementId railingTypeId = railing.GetTypeId();
                        ElementType railingType = doc.GetElement(railingTypeId) as ElementType;

                        if (railingType.LookupParameter("Green Area").AsInteger() == 1)
                        {
                            double railingArea = 0;
                            if (railingType.LookupParameter("Green Area").AsInteger() == 1)
                            {
                                if (semiSmartRound(railingType, "Railing Height") / 100 <= 2)
                                    railingArea += Math.Round(semiSmartRound(railing, "length") / 100
                                        * semiSmartRound(railingType, "Railing Height") / 100, 2, MidpointRounding.AwayFromZero);
                                else
                                    railingArea += Math.Round(railing.LookupParameter("Length").AsDouble() * lengthConvert / 100 * 2, 2, MidpointRounding.AwayFromZero);
                            }
                            if (railing.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                                greenArea1 += railingArea;
                            else if (railing.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                                greenArea2 += railingArea;
                            else
                                errorReport += $"Парапет с id: {railing.Id} има попълнен параметър A Instance Area Plot, чиято стойност не отговаря на нито едно от двете въведени имена за УПИ\n";
                        }
                    }

                    achievedPercentage1 = Math.Round(((greenArea1 * 100) / plotAreas[0]), 2);
                    achievedPercentage2 = Math.Round(((greenArea1 * 100) / plotAreas[1]), 2);
                    achievedPercentages.Add(achievedPercentage1);
                    achievedPercentages.Add(achievedPercentage2);
                    greenAreas.Add(greenArea1);
                    greenAreas.Add(greenArea2);
                }
            }
            catch (Exception e)
            {
                TaskDialog.Show("Warning", e.Message);
            }
        }
    }
}
