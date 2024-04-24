
using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Windows;

namespace AreaCalculations
{
    /// <summary>
    /// Interaction logic for AreaCoefficientsWindow.xaml
    /// </summary>
    public partial class AreaCoefficientsWindow : Window
    {
        public Dictionary<string, double> areaCoefficients = new Dictionary<string, double>
        {
            {"ACOR", 1 },
            {"ACLE", 1 },
            {"ACLO", 1 },
            {"ACHE", 1 },
            {"ACRO", 1 },
            {"ACSP", 1 },
            {"ACZO", 1 },
            {"ACCO", 1 },
            {"ACSTS", 0.3 },
            {"ACST", 1 },
            {"ACGAP", 0.8 },
            {"ACGA", 1 }
        };

        public bool overrideBool = false;

        public AreaCoefficientsWindow()
        {
            InitializeComponent();
        }

        private void SetAreaCoefficients(object sender, EventArgs e)
        {
            try
            {
                this.areaCoefficients["ACOR"] = Convert.ToDouble(ACOR.Text);
                this.areaCoefficients["ACLE"] = Convert.ToDouble(ACLE.Text);
                this.areaCoefficients["ACLO"] = Convert.ToDouble(ACLO.Text);
                this.areaCoefficients["ACHE"] = Convert.ToDouble(ACHE.Text);
                this.areaCoefficients["ACRO"] = Convert.ToDouble(ACRO.Text);
                this.areaCoefficients["ACSP"] = Convert.ToDouble(ACSP.Text);
                this.areaCoefficients["ACZO"] = Convert.ToDouble(ACZO.Text);
                this.areaCoefficients["ACCO"] = Convert.ToDouble(ACCO.Text);
                this.areaCoefficients["ACCO"] = Convert.ToDouble(ACCO.Text);
                this.areaCoefficients["ACSTS"] = Convert.ToDouble(ACSTS.Text);
                this.areaCoefficients["ACST"] = Convert.ToDouble(ACST.Text);
                this.areaCoefficients["ACGAP"] = Convert.ToDouble(ACGAP.Text);
                this.areaCoefficients["ACGA"] = Convert.ToDouble(ACGA.Text);

                if (overrideCoefficients.IsChecked != null)
                    this.overrideBool = (bool)overrideCoefficients.IsChecked;
            }
            catch { }

            this.Close();
        }
    }
}
