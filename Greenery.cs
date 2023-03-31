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
        double greenArea = 0;
        double greenArea1 = 0;
        double greenArea2 = 0;

        public Greenery(Autodesk.Revit.DB.Document doc, List<string> plotNames)
        {
            FilteredElementCollector allFloors = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Floors).WhereElementIsNotElementType();
            FilteredElementCollector allWalls = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Walls).WhereElementIsNotElementType();
            FilteredElementCollector allRailings = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Railings).WhereElementIsNotElementType();

            if (plotNames.Count == 1)
            {
                foreach (Floor floor in allFloors)
                    if (floor.FloorType.LookupParameter("Green Area").AsInteger() == 1) { greenArea += floor.LookupParameter("Area").AsDouble(); }

                foreach (Wall wall in allWalls)
                {
                    if (wall.WallType.LookupParameter("Green Area").AsInteger() == 1)
                    {
                        if (wall.LookupParameter("Unconnected Height").AsDouble() <= 6.56167979)
                            greenArea += wall.LookupParameter("Area").AsDouble();
                        else
                            greenArea += (wall.LookupParameter("Length").AsDouble() * 6.56167979);
                    }
                }

                foreach (Railing railing in allRailings)
                {
                    if (railing.LookupParameter("Green Area").AsInteger() == 1)
                    {
                        if (railing.LookupParameter("Railing Height").AsDouble() <= 6.56167979)
                            greenArea += (railing.LookupParameter("Length").AsDouble() * railing.LookupParameter("Railing Height").AsDouble());
                        else
                            greenArea += (railing.LookupParameter("Length").AsDouble() * 6.56167979);
                    }
                }
            }

            else if (plotNames.Count == 2)
            {
                foreach (Floor floor in allFloors)
                {
                    if (floor.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                        greenArea1 += floor.LookupParameter("Area").AsDouble();
                    else
                        greenArea2 += floor.LookupParameter("Area").AsDouble();
                }

                foreach (Wall wall in allWalls)
                {
                    double wallArea = 0;
                    if (wall.WallType.LookupParameter("Green Area").AsInteger() == 1)
                    {
                        if (wall.LookupParameter("Unconnected Height").AsDouble() <= 6.56167979)
                            wallArea += wall.LookupParameter("Area").AsDouble();
                        else
                            wallArea += (wall.LookupParameter("Length").AsDouble() * 6.56167979);
                    }
                    if (wall.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                        greenArea1 += wallArea;
                    else
                        greenArea2 += wallArea;
                }

                foreach (Railing railing in allRailings)
                {
                    double railingArea = 0;
                    if (railing.LookupParameter("Green Area").AsInteger() == 1)
                    {
                        if (railing.LookupParameter("Railing Height").AsDouble() <= 6.56167979)
                            railingArea += (railing.LookupParameter("Length").AsDouble() * railing.LookupParameter("Railing Height").AsDouble());
                        else
                            railingArea += (railing.LookupParameter("Length").AsDouble() * 6.56167979);
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
