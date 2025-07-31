using Autodesk.Revit.DB;
using TaskDialog = Autodesk.Revit.UI.TaskDialog;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using System.Windows.Controls;
using System.Windows.Forms;

namespace AreaCalculations
{
    internal class SmartRound
    {
        public readonly double areaConvert = 10.7639104167096;
        public readonly double lengthConvert = 30.48;
        private readonly Document doc;
        public SmartRound(Document document)
        {
            this.doc = document;
        }
        public double sqFeetToSqMeters(double valueFeet)
        {
            FormatOptions customFormat = new FormatOptions(UnitTypeId.SquareMeters);
            customFormat.Accuracy = 0.01;
            customFormat.UseDefault = false;

            FormatValueOptions formatOptions = new FormatValueOptions();
            formatOptions.AppendUnitSymbol = false;
            formatOptions.SetFormatOptions(customFormat);

            return double.Parse(UnitFormatUtils.Format(doc.GetUnits(), SpecTypeId.Area, valueFeet, true, formatOptions));
        }
        public double feetToCentimeters(double valueFeet)
        {
            FormatOptions customFormat = new FormatOptions(UnitTypeId.Centimeters);
            customFormat.Accuracy = 0.01;
            customFormat.UseDefault = false;

            FormatValueOptions formatOptions = new FormatValueOptions();
            formatOptions.AppendUnitSymbol = false;
            formatOptions.SetFormatOptions(customFormat);

            return double.Parse(UnitFormatUtils.Format(doc.GetUnits(), SpecTypeId.Length, valueFeet, true, formatOptions));
        }
    }
}
