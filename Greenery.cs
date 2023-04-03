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
        public int railingsCount { get; set; }

        double areaConvert = 10.763914692;

        public Greenery(Autodesk.Revit.DB.Document doc, List<string> plotNames)
        {
            FilteredElementCollector allFloors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
            FilteredElementCollector allWalls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();
            FilteredElementCollector allRailings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_StairsRailing).WhereElementIsNotElementType();

            railingsCount = allRailings.Count();

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
            }

            else if (plotNames.Count == 2)
            {
                foreach (Floor floor in allFloors)
                {
                    if (floor.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                        greenArea1 += Math.Round(floor.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    else
                        greenArea2 += Math.Round(floor.LookupParameter("Area").AsDouble() / areaConvert, 2);
                }

                foreach (Wall wall in allWalls)
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
                    else
                        greenArea2 += wallArea;
                }

                foreach (Railing railing in allRailings)
                {
                    ElementId railingTypeId = railing.GetTypeId();
                    ElementType railingType = doc.GetElement(railingTypeId) as ElementType;

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
                    else
                        greenArea2 += railingArea;
                }
            }
        }
    }
}
