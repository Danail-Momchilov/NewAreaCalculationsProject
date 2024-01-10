using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Document = Autodesk.Revit.DB.Document;

namespace AreaCalculations
{
    internal class AreaDictionary
    {
        public Dictionary<string, Dictionary<string, List<Area>>> AreasOrganizer { get; set; }
        public List<string> plotNames { get; set; }
        public Dictionary<string, List<string>> plotProperties { get; set; }
        public Document doc { get; set; }
        public Transaction transaction { get; set; }
        public AreaDictionary(Document activeDoc)
        {
            this.doc = activeDoc;
            this.AreasOrganizer = new Dictionary<string, Dictionary<string, List<Area>>>();
            this.plotNames = new List<string>();
            this.plotProperties = new Dictionary<string, List<string>>();
            this.transaction = new Transaction(activeDoc, "Calculate and Update Area Parameters");

            FilteredElementCollector areasCollector = new FilteredElementCollector(activeDoc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();

            foreach (Area area in areasCollector)
            {
                if (area != null)
                {
                    string plotName = area.LookupParameter("A Instance Area Plot").AsString();
                    string groupName = area.LookupParameter("A Instance Area Group").AsString();

                    if (!string.IsNullOrEmpty(plotName) && !string.IsNullOrEmpty(groupName))
                    {
                        if (!AreasOrganizer.ContainsKey(plotName))
                        {
                            this.AreasOrganizer.Add(plotName, new Dictionary<string, List<Area>>());
                            this.plotNames.Add(plotName);
                            this.plotProperties.Add(plotName, new List<string>());
                        }
                    
                        if (!AreasOrganizer[plotName].ContainsKey(groupName))
                        {
                            this.AreasOrganizer[plotName].Add(groupName, new List<Area>());
                            this.plotProperties[plotName].Add(groupName);
                        }
                    
                        this.AreasOrganizer[plotName][groupName].Add(area);
                    }

                    else
                    {
                        // TODO: to be continued
                        // TODO: check whether an Area has a plotName, that is not relative to the ones in the Project Info (is it actually needed?)
                    }
                }
            }
        }

        private double areaConvert = 10.763914692;
        public string calculatePrimaryArea()
        {
            string errorMessage = "";
            List<string> missingNumbers = new List<string>();

            transaction.Start();

            foreach (Area mainArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType())
            {
                mainArea.LookupParameter("A Instance Gross Area").Set(mainArea.LookupParameter("Area").AsDouble());
            }
            
            foreach (Area secondaryArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType())
            {
                bool wasfound = false;
                
                if (secondaryArea.LookupParameter("A Instance Area Primary").HasValue && secondaryArea.LookupParameter("A Instance Area Primary").AsString() != "")
                {
                    foreach (Area mainArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType())
                    {
                        if (secondaryArea.LookupParameter("A Instance Area Primary").AsString() == mainArea.LookupParameter("Number").AsString())
                        {
                            double sum = mainArea.LookupParameter("A Instance Gross Area").AsDouble() + secondaryArea.LookupParameter("Area").AsDouble();

                            bool wasSet = mainArea.LookupParameter("A Instance Gross Area").Set(sum);

                            wasfound = true;

                            // test

                            errorMessage += $"Добавяме площта: {secondaryArea.LookupParameter("Area").AsDouble()/areaConvert} към тази на Area: {mainArea.Number}, която към момента е: {mainArea.Area/areaConvert}! Общата сума е {secondaryArea.LookupParameter("Area").AsDouble() / areaConvert + mainArea.Area / areaConvert}\nsum = {sum/areaConvert}\nsum in imperial = {sum}\nwasSet = {wasSet}\n\n";

                            // test
                        }
                    }

                    if (!wasfound && !missingNumbers.Contains(secondaryArea.LookupParameter("Number").AsString()))
                    {
                        missingNumbers.Add(secondaryArea.LookupParameter("Number").AsString());
                        errorMessage += $"Грешка: Area {secondaryArea.LookupParameter("Number").AsString()} / id: {secondaryArea.Id} / Посочената Area е зададена като подчинена на такава с несъществуващ номер. Моля, проверете го и стартирайте апликацията отново\n";
                    }
                }
            }

            transaction.Commit();

            return errorMessage;
        }
    }
}
