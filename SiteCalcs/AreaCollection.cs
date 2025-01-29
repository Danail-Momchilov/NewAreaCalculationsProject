﻿using System;
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
        private double areaConvert = 10.7639104167096;
        private double semiRoundArea(Area area, string parameterName)
        {
            double result = Math.Round(area.LookupParameter(parameterName).AsDouble() / areaConvert, 2, MidpointRounding.AwayFromZero);

            return result;
        }
        private bool updateIfNoValue(Parameter param, double value)
        {
            if (param.HasValue && param.AsValueString() != "" && param.AsDouble() != 0)
                return false;
            else
            {
                param.Set(value);
                return true;
            }
        }     
        public AreaCollection(Document document)
        {
            this.doc = document;

            ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ROOM_AREA));
            FilterNumericRuleEvaluator evaluator = new FilterNumericGreater();
            double epsilon = 0.0001;
            ElementParameterFilter filter = new ElementParameterFilter(new FilterDoubleRule(provider, evaluator, epsilon, 1E-6));

            this.areasCollector = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().WherePasses(filter);

            this.transaction = new Transaction(doc, "Update Areas");
        }
        public AreaCollection(Document document, List<string> plotNames)
        {
            this.doc = document;

            ParameterValueProvider provider = new ParameterValueProvider(new ElementId(BuiltInParameter.ROOM_AREA));
            FilterNumericRuleEvaluator evaluator = new FilterNumericGreater();
            double epsilon = 0.0001;
            ElementParameterFilter filter = new ElementParameterFilter(new FilterDoubleRule(provider, evaluator, epsilon, 1E-6));

            this.areasCollector = new FilteredElementCollector(document).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().WherePasses(filter);

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
                            this.build[0] += semiRoundArea(area, "Area");
                        if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА")
                            this.totalBuild[0] += semiRoundArea(area, "Area");
                    }
                    else if (plotNames.Count == 2)
                    {
                        if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА" || area.LookupParameter("A Instance Area Location").AsString() == "ПОЛУПОДЗЕМНА")
                        {
                            if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                                this.build[0] += semiRoundArea(area, "Area");
                            else if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                                this.build[1] += semiRoundArea(area, "Area");
                        }
                        else if (area.LookupParameter("A Instance Area Location").AsString() == "НАДЗЕМНА")
                        {
                            if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[0])
                                this.totalBuild[0] += semiRoundArea(area, "Area");
                            else if (area.LookupParameter("A Instance Area Plot").AsString() == plotNames[1])
                                this.totalBuild[1] += semiRoundArea(area, "Area");
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
        public string CheckAreasParameters(List<string> plotNames, ProjectInfo projInfo)
        {
            string errorMessage = "";

            List<string> AreaCategoryValues = new List<string> { "ИЗКЛЮЧЕНА ОТ ОЧ", "НЕПРИЛОЖИМО", "ОБЩА ЧАСТ", "САМОСТОЯТЕЛЕН ОБЕКТ" };
            List<string> AreaLocationValues = new List<string> { "НАДЗЕМНА", "НАЗЕМНА", "НЕПРИЛОЖИМО", "ПОДЗЕМНА", "ПОЛУПОДЗЕМНА" };

            List<string> plotTypesValues = new List<string>();

            if (new List<string> { "СТАНДАРТНО УПИ", "ЪГЛОВО УПИ", "УПИ В ДВЕ ЗОНИ" }.Contains(projInfo.LookupParameter("Plot Type").AsString()))
            {
                plotTypesValues.Add(projInfo.LookupParameter("Plot Number").AsString());
            }
            else if (projInfo.LookupParameter("Plot Type").AsString() == "ДВЕ УПИ")
            {
                plotTypesValues.Add(projInfo.LookupParameter("Plot Number 1st").AsString());
                plotTypesValues.Add(projInfo.LookupParameter("Plot Number 2nd").AsString());
            }

            foreach (Area area in areasCollector)
            {
                if (area.Area != 0)
                {
                    if (area.LookupParameter("A Instance Area Group").AsString() == "") { errorMessage += $"Грешка: " +
                            $"Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Group \n"; }
                    else if ((!new List<string> { "ТРАФ", "ЗЕМЯ" }.Contains(area.LookupParameter("A Instance Area Group").AsString())) && 
                        (area.LookupParameter("A Instance Area Category").AsString() == "НЕПРИЛОЖИМО"))

                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} " +
                            $"/ Параметър: A Instance Area Category. Индивидуален обект със зададена стойност за 'A Instance Area Group', " +
                            $"различна от 'ТРАФ' и 'ЗЕМЯ', не може да приеме стойност 'НЕПРИЛОЖИМО' за 'A Instance Area Category'\n"; }

                    if (area.LookupParameter("A Instance Area Category").AsString() == "") { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} " +
                            $"/ id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Category\n"; }
                    else if (!AreaCategoryValues.Contains(area.LookupParameter("A Instance Area Category").AsString()))

                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} " +
                            $"/ Параметър: A Instance Area Category. Допустими стойности: ИЗКЛЮЧЕНА ОТ ОЧ, НЕПРИЛОЖИМО, ОБЩА ЧАСТ, САМОСТОЯТЕЛЕН ОБЕКТ\n"; }

                    if (area.LookupParameter("A Instance Area Location").AsString() == "") { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} " +
                            $"/ id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Location\n"; }

                    else if (!AreaLocationValues.Contains(area.LookupParameter("A Instance Area Location").AsString()))
                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} " +
                            $"/ Параметър: A Instance Area Location. Допустими стойности: НАДЗЕМНА, НАЗЕМНА, НЕПРИЛОЖИМО, ПОДЗЕМНА, ПОЛУПОДЗЕМНА\n"; }

                    if (area.LookupParameter("A Instance Area Category").AsString() == "НЕПРИЛОЖИМО" && 
                        area.LookupParameter("A Instance Area Location").AsString() != "НЕПРИЛОЖИМО" && area.LookupParameter("A Instance Area Location").AsString() != "")

                    { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} " +
                            $"/ Параметър: A Instance Area Location. Когато за параметър 'A Instance Area Category' е попълнена стойност 'НЕПРИЛОЖИМО', " +
                            $"то за 'A Instance Area Location' трябва да бъде зададена същата стойност\n"; }

                    if (area.LookupParameter("A Instance Area Entrance").AsString() == "") { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} " +
                            $"/ id: {area.Id.ToString()} / Непопълнен параметър: A Instance Area Entrance\n"; }

                    if (plotNames.Count == 2)
                    {
                        if (!plotNames.Contains(area.LookupParameter("A Instance Area Plot").AsString()))

                        { errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} " +
                                $"/ Параметър: A Instance Area Plot. Допустими стойности: {plotNames[0]} и {plotNames[1]}\n"; }
                    }
                    
                    if (!plotTypesValues.Contains(area.LookupParameter("A Instance Area Plot").AsString()))
                    {
                        string allPlotsStr = "";

                        foreach (string plot in plotTypesValues)
                            allPlotsStr += $"| {plot} | ";

                        errorMessage += $"Грешка: Area {area.LookupParameter("Number").AsString()} / id: {area.Id.ToString()} " +
                            $"/ Параметър: A Instance Area Plot. Допустими стойности: {allPlotsStr}\n";
                    }                    
                }
            }

            return errorMessage;
        }
        public int updateAreaCoefficients(Dictionary<string, double> areaCoefficientsDict)
        {
            int i = 0;

            double ACOR = areaCoefficientsDict["ACOR"];
            double ACLE = areaCoefficientsDict["ACLE"];
            double ACLO = areaCoefficientsDict["ACLO"];
            double ACHE = areaCoefficientsDict["ACHE"];
            double ACRO = areaCoefficientsDict["ACRO"];
            double ACSP = areaCoefficientsDict["ACSP"];
            double ACZO = areaCoefficientsDict["ACZO"];
            double ACCO = areaCoefficientsDict["ACCO"];
            double ACSTS = areaCoefficientsDict["ACSTS"];
            double ACST = areaCoefficientsDict["ACST"];
            double ACGAP = areaCoefficientsDict["ACGAP"];
            double ACGA = areaCoefficientsDict["ACGA"];

            transaction.Start();

            foreach (var area in areasCollector)
            {
                bool wasUpdated = false;

                if (area.LookupParameter("Area").AsString() != "Not Placed")
                {
                    double totalMultiplied = 0;

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Orientation (Ки)"), ACOR);
                    totalMultiplied = area.LookupParameter("A Coefficient Orientation (Ки)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Level (Кв)"), ACLE);
                    totalMultiplied *= area.LookupParameter("A Coefficient Level (Кв)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Location (Км)"), ACLO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Location (Км)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Height (Кив)"), ACHE);
                    totalMultiplied *= area.LookupParameter("A Coefficient Height (Кив)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Roof (Кпп)"), ACRO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Roof (Кпп)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Special (Кок)"), ACSP);
                    totalMultiplied *= area.LookupParameter("A Coefficient Special (Кок)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Zones (Кк)"), ACZO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Zones (Кк)").AsDouble();

                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Correction"), ACCO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Correction").AsDouble();

                    if (new List<string> { "склад", "мазе" }.Contains(area.LookupParameter("Name").AsString().ToLower()))
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Storage (Ксп)"), ACSTS);
                    else
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Storage (Ксп)"), ACST);                        

                    totalMultiplied *= area.LookupParameter("A Coefficient Storage (Ксп)").AsDouble();

                    if (new List<string> { "гараж", "паркинг" }.Contains(area.LookupParameter("Name").AsString().ToLower()))
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Garage (Кпг)"), ACGAP);
                    else
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Garage (Кпг)"), ACGA);

                    totalMultiplied *= area.LookupParameter("A Coefficient Garage (Кпг)").AsDouble();

                    area.LookupParameter("A Coefficient Multiplied").Set(totalMultiplied);
                }
                if (wasUpdated)
                    i += 1;
            }

            transaction.Commit();
            return i;
        }
        public int updateAreaCoefficientsOverride(Dictionary<string, double> areaCoefficientsDict)
        {
            int i = 0;

            double ACOR = areaCoefficientsDict["ACOR"];
            double ACLE = areaCoefficientsDict["ACLE"];
            double ACLO = areaCoefficientsDict["ACLO"];
            double ACHE = areaCoefficientsDict["ACHE"];
            double ACRO = areaCoefficientsDict["ACRO"];
            double ACSP = areaCoefficientsDict["ACSP"];
            double ACZO = areaCoefficientsDict["ACZO"];
            double ACCO = areaCoefficientsDict["ACCO"];
            double ACSTS = areaCoefficientsDict["ACSTS"];
            double ACST = areaCoefficientsDict["ACST"];
            double ACGAP = areaCoefficientsDict["ACGAP"];
            double ACGA = areaCoefficientsDict["ACGA"];

            transaction.Start();

            foreach (var area in areasCollector)
            {
                if (area.LookupParameter("Area").AsString() != "Not Placed")
                {
                    double totalMultiplied = 0;

                    area.LookupParameter("A Coefficient Orientation (Ки)").Set(ACOR);
                    totalMultiplied = area.LookupParameter("A Coefficient Orientation (Ки)").AsDouble();

                    area.LookupParameter("A Coefficient Level (Кв)").Set(ACLE);
                    totalMultiplied *= area.LookupParameter("A Coefficient Level (Кв)").AsDouble();

                    area.LookupParameter("A Coefficient Location (Км)").Set(ACLO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Location (Км)").AsDouble();

                    area.LookupParameter("A Coefficient Height (Кив)").Set(ACHE);
                    totalMultiplied *= area.LookupParameter("A Coefficient Height (Кив)").AsDouble();

                    area.LookupParameter("A Coefficient Roof (Кпп)").Set(ACRO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Roof (Кпп)").AsDouble();

                    area.LookupParameter("A Coefficient Special (Кок)").Set(ACSP);
                    totalMultiplied *= area.LookupParameter("A Coefficient Special (Кок)").AsDouble();

                    area.LookupParameter("A Coefficient Zones (Кк)").Set(ACZO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Zones (Кк)").AsDouble();

                    area.LookupParameter("A Coefficient Correction").Set(ACCO);
                    totalMultiplied *= area.LookupParameter("A Coefficient Correction").AsDouble();

                    if (new List<string> { "склад", "мазе" }.Any(keyword => area.LookupParameter("Name").AsString().ToLower().Contains(keyword)))
                        area.LookupParameter("A Coefficient Storage (Ксп)").Set(ACSTS);
                    else
                        area.LookupParameter("A Coefficient Storage (Ксп)").Set(ACST);

                    totalMultiplied *= area.LookupParameter("A Coefficient Storage (Ксп)").AsDouble();

                    if (new List<string> { "гараж", "паркинг" }.Any(keyword => area.LookupParameter("Name").AsString().ToLower().Contains(keyword)))
                        area.LookupParameter("A Coefficient Garage (Кпг)").Set(ACGAP);
                    else
                        area.LookupParameter("A Coefficient Garage (Кпг)").Set(ACGA);

                    totalMultiplied *= area.LookupParameter("A Coefficient Garage (Кпг)").AsDouble();

                    area.LookupParameter("A Coefficient Multiplied").Set(totalMultiplied);
                }
                i += 1;
            }

            transaction.Commit();
            return i;
        }
    }
}
