using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
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
        public string errorReport = "";

        double areaConvert = 10.763914692;

        public Greenery(Autodesk.Revit.DB.Document doc, List<string> plotNames, List<double> plotAreas)
        {
            FilteredElementCollector allFloors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
            FilteredElementCollector allWalls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();
            FilteredElementCollector allRailings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRailing).WhereElementIsNotElementType();

            if (plotNames.Count == 1)
            {
                foreach (Floor floor in allFloors)
                    if (floor.FloorType.LookupParameter("Green Area").AsInteger() == 1) { Math.Round(greenArea += floor.LookupParameter("Area").AsDouble() / areaConvert, 2); }

                foreach (Wall wall in allWalls)
                {
                    if (wall.WallType.LookupParameter("Green Area").AsInteger() == 1)
                    {
                        if ((wall.LookupParameter("Unconnected Height").AsDouble() / areaConvert) <= 200)
                            greenArea += Math.Round(wall.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        else
                            greenArea += Math.Round(((wall.LookupParameter("Length").AsDouble() / areaConvert) * 200), 2);
                    }
                }
                
                foreach (Railing railing in allRailings)
                {
                    ElementId railingTypeId = railing.GetTypeId();
                    ElementType railingType = doc.GetElement(railingTypeId) as ElementType;

                    if (railingType.LookupParameter("Green Area").AsInteger() == 1)
                    {
                        if (railingType.LookupParameter("Railing Height").AsDouble() / areaConvert  <= 200)
                            greenArea += Math.Round(((railing.LookupParameter("Length").AsDouble() / areaConvert) * (railingType.LookupParameter("Railing Height").AsDouble() / areaConvert)), 2);
                        else
                            greenArea += Math.Round(((railing.LookupParameter("Length").AsDouble() / areaConvert) * 200), 2);
                    }
                }

                achievedPercentage = Math.Round(((greenArea * 100) / plotAreas[0]), 2);
            }

            else if (plotNames.Count == 2)
            {
                foreach (Floor floor in allFloors)
                {
                    if (floor.FloorType.LookupParameter("Green Area").AsInteger() == 1)
                    {
                        if (floor.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                            greenArea1 += Math.Round(floor.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        else if (floor.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                            greenArea2 += Math.Round(floor.LookupParameter("Area").AsDouble() / areaConvert, 2);
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
                            if (wall.LookupParameter("Unconnected Height").AsDouble() / areaConvert <= 200)
                                wallArea += Math.Round((wall.LookupParameter("Area").AsDouble() / areaConvert), 2);
                            else
                                wallArea += Math.Round((wall.LookupParameter("Length").AsDouble() * 200), 2);
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
                            if ((railingType.LookupParameter("Railing Height").AsDouble() / areaConvert) <= 200)
                                railingArea += Math.Round(((railing.LookupParameter("Length").AsDouble() / areaConvert) * (railingType.LookupParameter("Railing Height").AsDouble() / areaConvert)), 2);
                            else
                                railingArea += Math.Round((railing.LookupParameter("Length").AsDouble() * 200), 2);
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
            }
        }
    }
}
