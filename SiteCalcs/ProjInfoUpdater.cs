using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;
using Autodesk.Revit.DB.ExtensibleStorage;
using Autodesk.Revit.UI;

namespace AreaCalculations
{
    internal class ProjInfoUpdater
    {
        public ProjectInfo ProjectInfo { get; set; }
        Document doc { get; set; }
        Transaction transaction { get; set; }
        public List<double> plotAreas { get; set; } = new List<double>();
        public List<string> plotNames { get; set; } = new List<string>();
        public bool isPlotTypeCorrect = true;
        double areaConvert = 10.7639104167096;
        private SmartRound smartRounder { get; set; }
        public ProjInfoUpdater(ProjectInfo projectInfo, Document Doc)
        {
            this.ProjectInfo = projectInfo;
            this.doc = Doc;
            this.transaction = new Transaction(doc, "Update Project Info");
            this.smartRounder = new SmartRound(doc);

            switch (ProjectInfo.LookupParameter("Plot Type").AsString())
            {
                case "СТАНДАРТНО УПИ":
                    plotNames.Add(projectInfo.LookupParameter("Plot Number").AsString());
                    plotAreas.Add(smartRounder.sqFeetToSqMeters(projectInfo.LookupParameter("Plot Area").AsDouble()));
                    break;

                case "ЪГЛОВО УПИ":
                    plotNames.Add(projectInfo.LookupParameter("Plot Number").AsString());
                    plotAreas.Add(smartRounder.sqFeetToSqMeters(projectInfo.LookupParameter("Plot Area").AsDouble()));
                    break;

                case "УПИ В ДВЕ ЗОНИ":
                    double plotAr = smartRounder.sqFeetToSqMeters(projectInfo.LookupParameter("Zone Area 1st").AsDouble()) 
                        + smartRounder.sqFeetToSqMeters(projectInfo.LookupParameter("Zone Area 2nd").AsDouble());
                    plotAreas.Add(plotAr);
                    plotNames.Add(projectInfo.LookupParameter("Plot Number").AsString());
                    break;

                case "ДВЕ УПИ":
                    plotNames.Add(projectInfo.LookupParameter("Plot Number 1st").AsString());
                    plotNames.Add(projectInfo.LookupParameter("Plot Number 2nd").AsString());
                    plotAreas.Add(smartRounder.sqFeetToSqMeters(projectInfo.LookupParameter("Plot Area 1st").AsDouble()));
                    plotAreas.Add(smartRounder.sqFeetToSqMeters(projectInfo.LookupParameter("Plot Area 2nd").AsDouble()));
                    break;

                default:
                    isPlotTypeCorrect = false;
                    break;
            }
        }
        public ProjInfoUpdater(Document doc)
        {
            this.ProjectInfo = doc.ProjectInformation;
        }
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
        public string CheckProjectInfoParameters()
        {
            try
            {
                string missingParameters = "";

                if (ProjectInfo.LookupParameter("Plot Type") == null) { missingParameters += "Липсва параметър 'Plot Type', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Plot Area") == null) { missingParameters += "Липсва параметър 'Plot Area', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Plot Area 1st") == null) { missingParameters += "Липсва параметър 'Plot Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Plot Area 2nd") == null) { missingParameters += "Липсва параметър 'Plot Area 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Zone Area 1st") == null) { missingParameters += "Липсва параметър 'Zone Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Zone Area 2nd") == null) { missingParameters += "Липсва параметър 'Zone Area 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Plot Number") == null) { missingParameters += "Липсва параметър 'Plot Number', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Plot Number 1st") == null) { missingParameters += "Липсва параметър 'Plot Number 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Plot Number 2nd") == null) { missingParameters += "Липсва параметър 'Plot Number 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Urban Index") == null) { missingParameters += "Липсва параметър 'Urban Index', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Urban Index 1st") == null) { missingParameters += "Липсва параметър 'Urban Index 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Urban Index 2nd") == null) { missingParameters += "Липсва параметър 'Urban Index 2nd', моля заредете го и опитайте отново...\n"; }
                /*
                if (ProjectInfo.LookupParameter("Required Urban Building Height") == null) { missingParameters += "Липсва параметър 'Required Urban Building Height', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Urban Building Height 1st") == null) { missingParameters += "Липсва параметър 'Required Urban Building Height 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Urban Building Height 2nd") == null) { missingParameters += "Липсва параметър 'Required Urban Building Height 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Permit Building Height") == null) { missingParameters += "Липсва параметър 'Required Permit Building Height', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Permit Building Height 1st") == null) { missingParameters += "Липсва параметър 'Required Permit Building Height 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Permit Building Height 2nd") == null) { missingParameters += "Липсва параметър 'Required Permit Building Height 2nd', моля заредете го и опитайте отново...\n"; }
                */
                if (ProjectInfo.LookupParameter("Required Built up Density") == null) { missingParameters += "Липсва параметър 'Required Built up Density', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Built up Density 1st") == null) { missingParameters += "Липсва параметър 'Required Built up Density 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Built up Density 2nd") == null) { missingParameters += "Липсва параметър 'Required Built up Density 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Built up Area") == null) { missingParameters += "Липсва параметър 'Required Built up Area', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Built up Area 1st") == null) { missingParameters += "Липсва параметър 'Required Built up Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Built up Area 2nd") == null) { missingParameters += "Липсва параметър 'Required Built up Area 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Area Intensity") == null) { missingParameters += "Липсва параметър 'Required Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Area Intensity 1st") == null) { missingParameters += "Липсва параметър 'Required Area Intensity 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Area Intensity 2nd") == null) { missingParameters += "Липсва параметър 'Required Area Intensity 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Gross External Area") == null) { missingParameters += "Липсва параметър 'Required Gross External Area', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Gross External Area 1st") == null) { missingParameters += "Липсва параметър 'Required Gross External Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Gross External Area 2nd") == null) { missingParameters += "Липсва параметър 'Required Gross External Area 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Green Area Percentage") == null) { missingParameters += "Липсва параметър 'Required Green Area Percentage', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Green Area Percentage 1st") == null) { missingParameters += "Липсва параметър 'Required Green Area Percentage 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Green Area Percentage 2nd") == null) { missingParameters += "Липсва параметър 'Required Green Area Percentage 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Green Area") == null) { missingParameters += "Липсва параметър 'Required Green Area', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Green Area 1st") == null) { missingParameters += "Липсва параметър 'Required Green Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Green Area 2nd") == null) { missingParameters += "Липсва параметър 'Required Green Area 2nd', моля заредете го и опитайте отново...\n"; }
                /*
                if (ProjectInfo.LookupParameter("Required Way of Construction") == null) { missingParameters += "Липсва параметър 'Required Way of Construction', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Way of Construction 1st") == null) { missingParameters += "Липсва параметър 'Required Way of Construction 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Way of Construction 2nd") == null) { missingParameters += "Липсва параметър 'Required Way of Construction 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Specifics") == null) { missingParameters += "Липсва параметър 'Required Specifics', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Specifics 1st") == null) { missingParameters += "Липсва параметър 'Required Specifics 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Specifics 2nd") == null) { missingParameters += "Липсва параметър 'Required Specifics 2nd', моля заредете го и опитайте отново...\n"; }
                */
                if (ProjectInfo.LookupParameter("Achieved Building Height") == null) { missingParameters += "Липсва параметър 'Achieved Building Height', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Building Height 1st") == null) { missingParameters += "Липсва параметър 'Achieved Building Height 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Building Height 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Building Height 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Built up Density") == null) { missingParameters += "Липсва параметър 'Achieved Built up Density', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Built up Density 1st") == null) { missingParameters += "Липсва параметър 'Achieved Built up Density 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Built up Density 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Built up Density 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Built up Area") == null) { missingParameters += "Липсва параметър 'Achieved Built up Area', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Built up Area 1st") == null) { missingParameters += "Липсва параметър 'Achieved Built up Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Built up Area 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Built up Area 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity 1st") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Gross External Area") == null) { missingParameters += "Липсва параметър 'Achieved Gross External Area', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Gross External Area 1st") == null) { missingParameters += "Липсва параметър 'Achieved Gross External Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Gross External Area 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Gross External Area 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Green Area Percentage") == null) { missingParameters += "Липсва параметър 'Achieved Green Area Percentage', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Green Area Percentage 1st") == null) { missingParameters += "Липсва параметър 'Achieved Green Area Percentage 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Green Area Percentage 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Green Area Percentage 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Green Area") == null) { missingParameters += "Липсва параметър 'Achieved Green Area', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Green Area 1st") == null) { missingParameters += "Липсва параметър 'Achieved Green Area 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Green Area 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Green Area 2nd', моля заредете го и опитайте отново...\n"; }
                /*
                if (ProjectInfo.LookupParameter("Required Motor Vehicle Places") == null) { missingParameters += "Липсва параметър 'Required Motor Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Motor Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Required Motor Vehicle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Motor Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Required Motor Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                */
                if (ProjectInfo.LookupParameter("Required Electrical Vehicle Places") == null) { missingParameters += "Required Electrical Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Electrical Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Required Electrical Vehicle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Electrical Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Required Electrical Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Disabled Vehicle Places") == null) { missingParameters += "Липсва параметър 'Required Disabled Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Disabled Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Required Disabled Vehicle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Disabled Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Required Disabled Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Total Vehicle Places") == null) { missingParameters += "Липсва параметър 'Required Total Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Total Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Required Total Vehicle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Total Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Required Total Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Bicycle Class 1 Places") == null) { missingParameters += "Липсва параметър 'Required Bicycle Class 1 Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Bicycle Class 1 Places 1st") == null) { missingParameters += "Липсва параметър 'Required Bicycle Class 1 Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Bicycle Class 1 Places 2nd") == null) { missingParameters += "Липсва параметър 'Required Bicycle Class 1 Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Bicycle Class 2 Places") == null) { missingParameters += "Липсва параметър 'Required Bicycle Class 2 Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Bicycle Class 2 Places 1st") == null) { missingParameters += "Липсва параметър 'Required Bicycle Class 2 Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Bicycle Class 2 Places 2nd") == null) { missingParameters += "Липсва параметър 'Required Bicycle Class 2 Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Area Intensity") == null) { missingParameters += "Липсва параметър 'Achieved Area Intensity', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Total Bicycle Places") == null) { missingParameters += "Липсва параметър 'Required Total Bicycle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Total Bicycle Places 1st") == null) { missingParameters += "Липсва параметър 'Required Total Bicycle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Required Total Bicycle Places 2nd") == null) { missingParameters += "Липсва параметър 'Required Total Bicycle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Motor Vehicle Places") == null) { missingParameters += "Липсва параметър 'Achieved Motor Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Motor Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Achieved Motor Vehicle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Motor Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Motor Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Electrical Vehicle Places") == null) { missingParameters += "Липсва параметър 'Achieved Electrical Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Electrical Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Achieved Electrical Vehicle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Electrical Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Electrical Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Disabled Vehicle Places") == null) { missingParameters += "Липсва параметър 'Achieved Disabled Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Disabled Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Achieved Disabled Vehicle Places 1st, моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Disabled Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Disabled Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Total Vehicle Places") == null) { missingParameters += "Липсва параметър 'Achieved Total Vehicle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Total Vehicle Places 1st") == null) { missingParameters += "Липсва параметър 'Achieved Total Vehicle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Total Vehicle Places 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Total Vehicle Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Bicycle Class 1 Places") == null) { missingParameters += "Липсва параметър 'Achieved Bicycle Class 1 Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Bicycle Class 1 Places 1st") == null) { missingParameters += "Липсва параметър 'Achieved Bicycle Class 1 Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Bicycle Class 1 Places 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Bicycle Class 1 Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Bicycle Class 2 Places") == null) { missingParameters += "Липсва параметър 'Achieved Bicycle Class 2 Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Bicycle Class 2 Places 1st") == null) { missingParameters += "Липсва параметър 'Achieved Bicycle Class 2 Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Bicycle Class 2 Places 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Bicycle Class 2 Places 2nd', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Total Bicycle Places") == null) { missingParameters += "Липсва параметър 'Achieved Total Bicycle Places', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Total Bicycle Places 1st") == null) { missingParameters += "Липсва параметър 'Achieved Total Bicycle Places 1st', моля заредете го и опитайте отново...\n"; }
                if (ProjectInfo.LookupParameter("Achieved Total Bicycle Places 2nd") == null) { missingParameters += "Липсва параметър 'Achieved Total Bicycle Places 2nd', моля заредете го и опитайте отново...\n"; }

                return missingParameters;
            }
            catch (Exception e)
            {
                throw e;
            }
        }
        public void SetAchievedStandard(double buildArea, double grossArea, double intensity, double density, double greenArea, double achievedPercentage)
        {
            transaction.Start();
            ProjectInfo.LookupParameter("Achieved Built up Area").Set(buildArea * areaConvert);
            ProjectInfo.LookupParameter("Achieved Gross External Area").Set(grossArea * areaConvert);
            ProjectInfo.LookupParameter("Achieved Area Intensity").Set(intensity);
            ProjectInfo.LookupParameter("Achieved Built up Density").Set(density);
            ProjectInfo.LookupParameter("Achieved Green Area").Set(greenArea * areaConvert);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage").Set(achievedPercentage);
            transaction.Commit();
        }        
        public void SetRequired(double buildArea, double grossArea, double intensity, double density)
        {
            transaction.Start();
            ProjectInfo.LookupParameter("Required Built up Area").Set(buildArea * areaConvert);
            ProjectInfo.LookupParameter("Required Gross External Area").Set(grossArea * areaConvert);
            ProjectInfo.LookupParameter("Required Area Intensity").Set(intensity);
            ProjectInfo.LookupParameter("Required Built up Density").Set(density);
            transaction.Commit();
        }
        public void SetAllTwoZones(double plotArea, double buildArea, double totalBuild, double kint, double density, double greenArea, double achievedPercentage)
        {
            transaction.Start();
            ProjectInfo.LookupParameter("Plot Area").Set(plotArea * areaConvert);
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
            ProjectInfo.LookupParameter("Achieved Built up Area").Set(buildArea * areaConvert);
            ProjectInfo.LookupParameter("Achieved Gross External Area").Set(totalBuild * areaConvert);
            ProjectInfo.LookupParameter("Achieved Area Intensity").Set(kint);
            ProjectInfo.LookupParameter("Achieved Built up Density").Set(density);
            // check this with the Simo
            ProjectInfo.LookupParameter("Achieved Green Area").Set(greenArea * areaConvert);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage").Set(achievedPercentage);
            // check this with the Simo
            transaction.Commit();
        }
        public void SetAchievedTwoPlots(double buildArea1, double totalBuild1, double kint1, double density1, double buildArea2, double totalBuild2, double kint2, double density2, double greenArea1, double greenArea2, double percentage1, double percentage2)
        {
            transaction.Start();
            ProjectInfo.LookupParameter("Achieved Built up Area 1st").Set(buildArea1 * areaConvert);
            ProjectInfo.LookupParameter("Achieved Gross External Area 1st").Set(totalBuild1 * areaConvert);
            ProjectInfo.LookupParameter("Achieved Area Intensity 1st").Set(kint1);
            ProjectInfo.LookupParameter("Achieved Built up Density 1st").Set(density1);
            ProjectInfo.LookupParameter("Achieved Built up Area 2nd").Set(buildArea2 * areaConvert);
            ProjectInfo.LookupParameter("Achieved Gross External Area 2nd").Set(totalBuild2 * areaConvert);
            ProjectInfo.LookupParameter("Achieved Area Intensity 2nd").Set(kint2);
            ProjectInfo.LookupParameter("Achieved Built up Density 2nd").Set(density2);
            ProjectInfo.LookupParameter("Achieved Green Area 1st").Set(greenArea1 * areaConvert);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage 1st").Set(percentage1);
            ProjectInfo.LookupParameter("Achieved Green Area 2nd").Set(greenArea2 * areaConvert);
            ProjectInfo.LookupParameter("Achieved Green Area Percentage 2nd").Set(percentage2);
            transaction.Commit();
        }
    }
}
