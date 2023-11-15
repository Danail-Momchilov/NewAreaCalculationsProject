using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;
using System.Windows.Controls;

namespace AreaCalculations.AreaCalcs
{
    internal class AreaUpdater
    {
        private FilteredElementCollector allAreas { get; set; }
        public Document document { get; set; }

        public AreaUpdater(Document doc)
        {
            this.allAreas = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();
            this.document = doc;
        }
        private static bool hasValue(Parameter param)
        {
            if (param.HasValue)
                return true;
            else
                return false;
        }

        private static bool updateIfNoValue(Parameter param, double value)
        {
            if (param.HasValue)
                return false;
            else
            {
                param.Set(value);
                return true;
            }                
        }

        public int updateAreaCoefficients()
        {
            int i = 0;

            Transaction transaction = new Transaction(document, "Update Areas");
            transaction.Start();

            foreach (var area in this.allAreas)
            {
                bool wasUpdated = false;
                if (area.LookupParameter("Area").AsString() != "Not Placed")
                {
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Orientation (Ки)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Level (Кв)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Location (Км)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Height (Кив)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Roof (Кпп)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Roof (Кпп)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Special (Кок)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Zones (Кк)"), 1);
                    wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Correction"), 1);

                    if (area.LookupParameter("Name").AsString().ToLower().Contains("склад"))
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Storage (Ксп)"), 0.3);
                    else
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Storage (Ксп)"), 1);

                    if (area.LookupParameter("Name").AsString().ToLower().Contains("гараж"))
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Garage (Кпг)"), 0.8);
                    else
                        wasUpdated = updateIfNoValue(area.LookupParameter("A Coefficient Garage (Кпг)"), 1);
                }
                if (wasUpdated)
                    i += 1;
            }

            transaction.Commit();
            return i;
        }

    }
}
