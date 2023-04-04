using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;

namespace AreaCalculations
{
    internal class AreaCollection
    {
        public List<double> build = new List<double>();
        public List<double> totalBuild = new List<double>();

        double areaConvert = 10.763914692;

        private static bool hasValue(Parameter param)
        {
            if (param.HasValue)
                return true;
            else
                return false;
        }

        public string CheckAreaParameters(List<string> plotNames, FilteredElementCollector Areas)
        {
            string errorMessage = "";

            List<string> AreaCategoryValues = new List<string> { "ИЗКЛЮЧЕНА ОТ ОЧ", "НЕПРИЛОЖИМО", "ОБЩА ЧАСТ", "САМОСТОЯТЕЛЕН ОБЕКТ" };
            List<string> AreaLocationValues = new List<string> { "НАДЗЕМНА", "НАЗЕМНА", "НЕПРИЛОЖИМО", "ПОДЗЕМНА", "ПОЛУПОДЗЕМНА" };

            foreach (Area area in Areas)
            {
                if (!hasValue(area.LookupParameter("A Instance Area Group"))) { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Group \n"; }
                if (!hasValue(area.LookupParameter("A Instance Area Category"))) { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Group\n"; }
                else if (!AreaCategoryValues.Contains(area.LookupParameter("A Instance Area Category").AsString())) 
                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Group. Допустими стойности: ИЗКЛЮЧЕНА ОТ ОЧ, НЕПРИЛОЖИМО, ОБЩА ЧАСТ, САМОСТОЯТЕЛЕН ОБЕКТ\n"; }
                if (!hasValue(area.LookupParameter("A Instance Area Location"))) { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Location\n"; }
                else if (!AreaLocationValues.Contains(area.LookupParameter("A Instance Area Location").AsString()))
                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Location. Допустими стойности: НАДЗЕМНА, НАЗЕМНА, НЕПРИЛОЖИМО, ПОДЗЕМНА, ПОЛУПОДЗЕМНА\n"; }
                if (!hasValue(area.LookupParameter("A Instance Area Entrance"))) { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Entrance\n"; }
                if (plotNames.Count == 2)
                {
                    if (!plotNames.Contains(area.LookupParameter("A Instance Area Plot").AsString()))
                        { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Plot. Допустими стойности: {plotNames[0]} и {plotNames[1]}\n"; }
                }
            }

            return errorMessage;
        }

        public AreaCollection(FilteredElementCollector Areas, List<string> plotNames)
        {
            this.build.Add(0);
            this.build.Add(0);

            this.totalBuild.Add(0);
            this.totalBuild.Add(0);

            foreach (Area area in Areas)
            {
                if (plotNames.Count == 1)
                {
                    if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                        this.build[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                        this.totalBuild[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                }
                else if (plotNames.Count == 2)
                {
                    if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                    {
                        if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                            this.build[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        else if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                            this.build[1] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                    else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                    {
                        if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                            this.totalBuild[0] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        else if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                            this.totalBuild[1] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                    }
                }
                else
                {
                    this.build[0] = plotNames.Count;
                    this.totalBuild[0] = 1;
                }
            }
        }
    }
}
