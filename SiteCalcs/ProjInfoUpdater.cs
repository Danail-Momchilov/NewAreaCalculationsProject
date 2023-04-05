using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;

namespace AreaCalculations
{
    internal class ProjInfoUpdater
    {
        public ProjectInfo ProjectInfo { get; set; }
        Document doc { get; set; }
        Transaction T { get; set; }
        public List<double> plotAreas { get; set; } = new List<double>();
        public List<string> plotNames { get; set; } = new List<string>();

        public bool isPlotTypeCorrect = true;
        double areaConvert = 10.763914692;

        private static bool hasValue(Parameter param)
        {
            if (param.HasValue)
                return true;
            else
                return false;
        }

        public string CheckProjectInfo()
        {
            string errorMessage = "";

            switch (ProjectInfo.LookupParameter("Plot Type").AsString())
            {
                 case "СТАНДАРТНО УПИ":
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Number"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Plot Number'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Urban Index"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Urban Index'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Urban Building Height"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Urban Building Height'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Permit Building Height"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Permit Building Height'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Density"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Built up Density' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Area"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Built up Area' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Area Intensity"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Area Intensity' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Gross External Area"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Gross External Area' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Green Area Percentage"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Green Area Percantage'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Green Area"))) { errorMessage += "При въведена опция 'СТАНДАРТНО УПИ' е нужно да попълните параметър 'Required Green Area' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Area"))) { errorMessage += "Айде въведи я тая 'Plot Area' де..."; }
                     break;

                 case "ЪГЛОВО УПИ":
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Number"))) { errorMessage += "При въведена опция 'ЪГЛОВО УПИ' се попълват точно 3 параметъра... 'Plot Number' е един от тях...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Urban Index"))) { errorMessage += "При въведена опция 'ЪГЛОВО УПИ'се попълват точно 3 параметъра... 'Urban Index' например...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Area"))) { errorMessage += "'Plot Area' въведи поне... you had one job...\n"; }
                     break;

                 case "УПИ В ДВЕ ЗОНИ":
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Number"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' все пак е нужно да попълните параметър 'Plot Number'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Zone Area 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Zone Area 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Zone Area 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Zone Area 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Urban Index 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Urban Index 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Urban Index 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Urban Index 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Urban Building Height 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Urban Building Height 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Urban Building Height 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Urban Building Height 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Permit Building Height 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Permit Building Height 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Permit Building Height 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Permit Building Height 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Density 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Built up Density 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Density 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Built up Density 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Area 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Built up Area 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Area 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Built up Area 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Area Intensity 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Area Intensity 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Area Intensity 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Area Intensity 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Gross External Area 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Gross External Area 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Gross External Area 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Gross External Area 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Green Area Percentage 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Green Area Percantage 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Green Area 1st"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Green Area 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Green Area 2nd"))) { errorMessage += "При въведена опция 'УПИ В ДВЕ ЗОНИ' е нужно да попълните параметър 'Required Green Area 2nd'...\n"; }
                     break;

                 case "ДВЕ УПИ":
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Area"))) { errorMessage += "'Plot Area въведи поне'... you had one job...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Number 1st"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Plot Number 1st'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Plot Number 2nd"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Plot Number 2nd'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Urban Building Height"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Urban Building Height'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Permit Building Height"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Permit Building Height'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Density"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Built up Density' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Built up Area"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Built up Area' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Area Intensity"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Area Intensity' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Gross External Area"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Gross External Area' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Green Area Percentage"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Green Area Percantage'...\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Required Green Area"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Required Green Area' !!!\n"; }
                     if (!hasValue(ProjectInfo.LookupParameter("Urban Index"))) { errorMessage += "При въведена опция 'ДВЕ УПИ' е нужно да попълните параметър 'Urban Index'...\n"; }
                     break;
             }

            return errorMessage;
        }
            

        public ProjInfoUpdater(ProjectInfo projectInfo, Document Doc)
        {
            this.ProjectInfo = projectInfo;
            this.doc = Doc;
            Transaction transaction = new Transaction(doc, "Update Project Info");
            T = transaction;

            switch (ProjectInfo.LookupParameter("Plot Type").AsString())
            {
                case "СТАНДАРТНО УПИ":
                    plotNames.Add(projectInfo.LookupParameter("Plot Number").AsString());
                    plotAreas.Add(projectInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                    break;
                case "ЪГЛОВО УПИ":
                    plotNames.Add(projectInfo.LookupParameter("Plot Number").AsString());
                    plotAreas.Add(projectInfo.LookupParameter("Plot Area").AsDouble() / areaConvert);
                    break;
                case "УПИ В ДВЕ ЗОНИ":
                    double plotAr = Math.Round((projectInfo.LookupParameter("Zone Area 1st").AsDouble() / areaConvert) + (projectInfo.LookupParameter("Zone Area 2nd").AsDouble() / areaConvert), 2);
                    plotAreas.Add(plotAr);
                    plotNames.Add(projectInfo.LookupParameter("Plot Number").AsString());
                    break;
                case "ДВЕ УПИ":
                    plotNames.Add(projectInfo.LookupParameter("Plot Number 1st").AsString());
                    plotNames.Add(projectInfo.LookupParameter("Plot Number 2nd").AsString());
                    plotAreas.Add(Math.Round(projectInfo.LookupParameter("Plot Area 1st").AsDouble() / areaConvert, 2));
                    plotAreas.Add(Math.Round(projectInfo.LookupParameter("Plot Area 2nd").AsDouble() / areaConvert, 2));
                    break;
                default:
                    isPlotTypeCorrect = false;
                    break;
            }
        }

        public void SetAchievedStandard(double buildArea, double grossArea, double intensity, double density, double greenArea, double achievedPercentage)
        {
            T.Start();
            ProjectInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
            ProjectInfo.LookupParameter("Achieved Gross External Area").Set(grossArea);
            ProjectInfo.LookupParameter("Achieved Area Intensity").Set(grossArea);
            ProjectInfo.LookupParameter("Achieved Built up Density").Set(density);
            ProjectInfo.LookupParameter("Achieved Green Area").Set(greenArea);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage").Set(achievedPercentage);
            T.Commit();
        }
        
        public void SetRequired(double buildArea, double grossArea, double intensity, double density)
        {
            T.Start();
            ProjectInfo.LookupParameter("Required Built up Area").Set(buildArea);
            ProjectInfo.LookupParameter("Required Gross External Area").Set(grossArea);
            ProjectInfo.LookupParameter("Required Area Intensity").Set(intensity);
            ProjectInfo.LookupParameter("Required Built up Density").Set(density);
            T.Commit();
        }

        public void SetAllTwoZones(double plotArea, double buildArea, double totalBuild, double kint, double density, double greenArea, double achievedPercentage)
        {
            T.Start();
            ProjectInfo.LookupParameter("Plot Area").Set(plotArea);
            ProjectInfo.LookupParameter("Required Built up Density")
                .Set((ProjectInfo.LookupParameter("Required Built up Density 1st").AsDouble() + ProjectInfo.LookupParameter("Required Built up Density 2nd").AsDouble()) / 2);
            ProjectInfo.LookupParameter("Required Built up Area")
                .Set(ProjectInfo.LookupParameter("Required Built up Area 1st").AsDouble() + ProjectInfo.LookupParameter("Required Built up Area 2nd").AsDouble());
            ProjectInfo.LookupParameter("Required Area Intensity")
                .Set((ProjectInfo.LookupParameter("Required Area Intensity 1st").AsDouble() + ProjectInfo.LookupParameter("Required Area Intensity 2nd").AsDouble() / 2));
            ProjectInfo.LookupParameter("Required Gross External Area")
                .Set(ProjectInfo.LookupParameter("Required Gross External Area 1st").AsDouble() + ProjectInfo.LookupParameter("Required Gross External Area 2nd").AsDouble());
            ProjectInfo.LookupParameter("Required Green Area Percentage")
                .Set((ProjectInfo.LookupParameter("Required Green Area Percentage 1st").AsDouble() + ProjectInfo.LookupParameter("Required Green Area Percentage 2nd").AsDouble()) / 2);
            ProjectInfo.LookupParameter("Required Green Area")
                .Set(ProjectInfo.LookupParameter("Required Green Area 1st").AsDouble() + ProjectInfo.LookupParameter("Required Green Area 2nd").AsDouble());
            ProjectInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
            ProjectInfo.LookupParameter("Achieved Gross External Area").Set(totalBuild);
            ProjectInfo.LookupParameter("Achieved Area Intensity").Set(kint);
            ProjectInfo.LookupParameter("Achieved Built up Density").Set(density);
            // check this with the Simo
            ProjectInfo.LookupParameter("Achieved Green Area").Set(greenArea);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage").Set(achievedPercentage);
            // check this with the Simo
            T.Commit();
        }

        public void SetAchievedTwoPlots(double buildArea1, double totalBuild1, double kint1, double density1, double buildArea2, double totalBuild2, double kint2, double density2, double greenArea1, double greenArea2, double percentage1, double percentage2)
        {
            T.Start();
            ProjectInfo.LookupParameter("Achieved Built up Area 1st").Set(buildArea1);
            ProjectInfo.LookupParameter("Achieved Gross External Area 1st").Set(totalBuild1);
            ProjectInfo.LookupParameter("Achieved Area Intensity 1st").Set(kint1);
            ProjectInfo.LookupParameter("Achieved Built up Density 1st").Set(density1);
            ProjectInfo.LookupParameter("Achieved Built up Area 2nd").Set(buildArea2);
            ProjectInfo.LookupParameter("Achieved Gross External Area 2nd").Set(totalBuild2);
            ProjectInfo.LookupParameter("Achieved Area Intensity 2nd").Set(kint2);
            ProjectInfo.LookupParameter("Achieved Built up Density 2nd").Set(density2);
            ProjectInfo.LookupParameter("Achieved Green Area 1st").Set(greenArea1);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage 1st").Set(percentage1);
            ProjectInfo.LookupParameter("Achieved Green Area 2nd").Set(greenArea2);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage 2nd").Set(percentage2);
            T.Commit();
        }
    }
}
