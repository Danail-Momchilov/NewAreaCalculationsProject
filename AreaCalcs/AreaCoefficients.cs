﻿using System;
using System.IO;
using System.Collections.Generic;
using System.Linq;
using Autodesk.Revit.Attributes;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.Architecture;
using Autodesk.Revit.UI;

namespace AreaCalculations
{
    [TransactionAttribute(TransactionMode.Manual)]
    internal class AreaCoefficients : IExternalCommand
    {        
        public Result Execute(ExternalCommandData commandData, ref string message, ElementSet elements)
        {
            try
            {
                UIDocument uidoc = commandData.Application.ActiveUIDocument;
                Document doc = uidoc.Document;

                AreaCollection areaUpdater = new AreaCollection(doc);
                int count = areaUpdater.updateAreaCoefficients();

                TaskDialog report = new TaskDialog("Report");
                if (count > 0)
                    report.MainInstruction = $"Успешно бяха обновени коефициентите на {count} 'Area' обекти!";
                else
                    report.MainInstruction = "Не са открити обекти от тип 'Area' с непопълнени параметри за коефициентите";
                report.Show();

                return Result.Succeeded;
            }
            catch (Exception e)
            {
                TaskDialog exceptions = new TaskDialog("Съобщение за грешка");
                exceptions.MainInstruction = $"{e.Message}\n\n {e.ToString()}";
                exceptions.Show();
                return Result.Failed;
            }
        }
    }
}
