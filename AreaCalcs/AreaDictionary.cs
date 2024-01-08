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
        public AreaDictionary(Document activeDoc)
        {
            this.doc = activeDoc;
            this.AreasOrganizer = new Dictionary<string, Dictionary<string, List<Area>>>();
            this.plotNames = new List<string>();
            this.plotProperties = new Dictionary<string, List<string>>();

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
        public void calculatePrimaryArea()
        {
            string errorMessage = "";

            foreach (string plot in plotNames)
            {
                foreach (string property in plotProperties[plot])
                {
                    // set main Area, based only on Area parameter first
                    foreach (Area mainArea in AreasOrganizer[plot][property])
                    {
                        mainArea.LookupParameter("A Instance Gross Area").Set(mainArea.LookupParameter("Area").AsDouble());
                    }

                    foreach (Area secondaryArea in AreasOrganizer[plot][property])
                    {
                        if (secondaryArea.LookupParameter("A Instance Area Primary").HasValue && secondaryArea.LookupParameter("A Instance Area Primary").AsString() != "")
                        {
                            bool wasfound = false;

                            foreach(Area mainArea in AreasOrganizer[plot][property])
                            {
                                if (secondaryArea.LookupParameter("A Instance Area Primary").AsString() == mainArea.LookupParameter("Number").AsString())
                                {
                                    mainArea.LookupParameter("A Instance Gross Area").Set(mainArea.LookupParameter("A Instance Gross Area").AsDouble() + secondaryArea.LookupParameter("Area").AsDouble());
                                }
                            }

                            if (!wasfound)
                            {

                            }
                        }
                    }
                }
            }
        }
    }
}
