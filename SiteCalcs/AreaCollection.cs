using System;
using System.Collections.Generic;
using System.Linq;
using System.Net;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.Creation;
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using Document = Autodesk.Revit.DB.Document;

namespace AreaCalculations
{
    internal class AreaCollection
    {
        public List<double> build { get; set; }
        public List<double> totalBuild { get; set; }
        public FilteredElementCollector areasCollector { get; set; }
        public Document doc { get; set; }
        Transaction transaction { get; set; }

        private double areaConvert = 10.763914692;

        private bool hasValue(Parameter param)
        {
            if (param.HasValue)
                return true;
            else
                return false;
        }

        private bool updateIfNoValue(Parameter param, double value)
        {
            if (param.HasValue)
                return false;
            else
            {
                param.Set(value);
                return true;
            }
        }

        private object belongsToArea(Area area)
        {
            foreach (Area mainArea in areasCollector)
            {
                if (area.LookupParameter("A Instance Area Entrance").AsString() == mainArea.LookupParameter("Number").AsString())
                    return mainArea;
                break;
            }

            return null;
        }        

        public AreaCollection(Document document)
        {
            this.doc = document;

            this.areasCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();

            this.transaction = new Transaction(doc, "Update Areas");
        }

        public AreaCollection(Document document, List<string> plotNames)
        {
            this.doc = document;

            this.areasCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();

            this.transaction = new Transaction(document, "Update Areas");

            this.build = new List<double>();
            this.totalBuild = new List<double>();

            this.build.Add(0);
            this.build.Add(0);

            this.totalBuild.Add(0);
            this.totalBuild.Add(0);

            foreach (Area area in areasCollector)
            {
                if (area.LookupParameter("Area").AsString() != "Not Placed")
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

        public string CheckAreasParameters(List<string> plotNames)
        {
            string errorMessage = "";

            List<string> AreaCategoryValues = new List<string> { "ИЗКЛЮЧЕНА ОТ ОЧ", "НЕПРИЛОЖИМО", "ОБЩА ЧАСТ", "САМОСТОЯТЕЛЕН ОБЕКТ" };
            List<string> AreaLocationValues = new List<string> { "НАДЗЕМНА", "НАЗЕМНА", "НЕПРИЛОЖИМО", "ПОДЗЕМНА", "ПОЛУПОДЗЕМНА" };

            foreach (Area area in areasCollector)
            {
                if (area.Area != 0)
                {
                    if (area.LookupParameter("A Instance Area Group").AsString() == "") { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Group \n"; }
                    else if ((!new List<string> { "ТРАФ", "ЗЕМЯ" }.Contains(area.LookupParameter("A Instance Area Group").AsString())) && (area.LookupParameter("A Instance Area Category").AsString() == "НЕПРИЛОЖИМО"))
                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Category. Индивидуален обект със зададена стойност за 'A Instance Area Group', различна от 'ТРАФ' и 'ЗЕМЯ', не може да приеме стойност 'НЕПРИЛОЖИМО' за 'A Instance Area Category'\n"; }

                    if (area.LookupParameter("A Instance Area Category").AsString() == "") { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Category\n"; }
                    else if (!AreaCategoryValues.Contains(area.LookupParameter("A Instance Area Category").AsString()))
                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Category. Допустими стойности: ИЗКЛЮЧЕНА ОТ ОЧ, НЕПРИЛОЖИМО, ОБЩА ЧАСТ, САМОСТОЯТЕЛЕН ОБЕКТ\n"; }

                    if (area.LookupParameter("A Instance Area Location").AsString() == "") { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Location\n"; }
                    else if (!AreaLocationValues.Contains(area.LookupParameter("A Instance Area Location").AsString()))
                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Location. Допустими стойности: НАДЗЕМНА, НАЗЕМНА, НЕПРИЛОЖИМО, ПОДЗЕМНА, ПОЛУПОДЗЕМНА\n"; }

                    if (area.LookupParameter("A Instance Area Category").AsString() == "НЕПРИЛОЖИМО" && area.LookupParameter("A Instance Area Location").AsString() != "НЕПРИЛОЖИМО" && area.LookupParameter("A Instance Area Location").AsString() != "")
                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Location. Когато за параметър 'A Instance Area Category' е попълнена стойност 'НЕПРИЛОЖИМО', то за 'A Instance Area Location' трябва да бъде зададена същата стойност\n"; }

                    if (area.LookupParameter("A Instance Area Entrance").AsString() == "") { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Entrance\n"; }

                    if (plotNames.Count == 2)
                    {
                        if (!plotNames.Contains(area.LookupParameter("A Instance Area Plot").AsString()))
                        { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Параметър: A Instance Area Plot. Допустими стойности: {plotNames[0]} и {plotNames[1]}\n"; }
                    }
                }
            }

            return errorMessage;
        }

        public int updateAreaCoefficients()
        {
            int i = 0;

            transaction.Start();

            foreach (var area in areasCollector)
            {
                bool wasUpdated = false;

                if (area.LookupParameter("Area").AsString() != "Not Placed")
                {
                    double totalMultiplied = 0;

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Orientation (Ки)"), 1);
                    totalMultiplied = area.LookupParameter("A Coefficient Orientation (Ки)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Level (Кв)"), 1);
                    totalMultiplied *= area.LookupParameter("A Coefficient Level (Кв)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Location (Км)"), 1);
                    totalMultiplied *= area.LookupParameter("A Coefficient Location (Км)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Height (Кив)"), 1);
                    totalMultiplied *= area.LookupParameter("A Coefficient Height (Кив)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Roof (Кпп)"), 1);
                    totalMultiplied *= area.LookupParameter("A Coefficient Roof (Кпп)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Special (Кок)"), 1);
                    totalMultiplied *= area.LookupParameter("A Coefficient Special (Кок)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Zones (Кк)"), 1);
                    totalMultiplied *= area.LookupParameter("A Coefficient Zones (Кк)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Correction"), 1);
                    totalMultiplied *= area.LookupParameter("A Coefficient Correction").AsDouble();

                    if (new List<string> { "склад", "мазе" }.Contains(area.LookupParameter("Name").AsString().ToLower()))
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Storage (Ксп)"), 0.3);
                    else
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Storage (Ксп)"), 1);                        

                    totalMultiplied *= area.LookupParameter("A Coefficient Storage (Ксп)").AsDouble();

                    if (new List<string> { "гараж", "паркинг" }.Contains(area.LookupParameter("Name").AsString().ToLower()))
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Garage (Кпг)"), 0.8);
                    else
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Garage (Кпг)"), 1);

                    totalMultiplied *= area.LookupParameter("A Coefficient Garage (Кпг)").AsDouble();

                    area.LookupParameter("A Coefficient Multiplied").Set(totalMultiplied);
                }
                if (wasUpdated)
                    i += 1;
            }

            transaction.Commit();
            return i;
        }
    }
}
