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

        public string CheckProjectInfo()
        {
            string errorMessage = "";

            if (ProjectInfo.LookupParameter("Plot Type").AsString() != "СТАНДАРТНО УПИ" && ProjectInfo.LookupParameter("Plot Type").AsString() != "ЪГЛОВО УПИ"
                && ProjectInfo.LookupParameter("Plot Type").AsString() != "УПИ В ДВЕ ЗОНИ" && ProjectInfo.LookupParameter("Plot Type").AsString() != "ДВЕ УПИ")
            { errorMessage += "Моля попълнете параметър 'Plot Type' с една от четирите посочени опции: СТАНДАРТНО УПИ, ЪГЛОВО УПИ, УПИ В ДВЕ ЗОНИ, ДВЕ УПИ!\n"; }
            else
            {
                // fhfghgh
                switch (ProjectInfo.LookupParameter("Plot Type").AsString())
                {
                    case "СТАНДАРТНО УПИ":

                        break;

                    case "ЪГЛОВО УПИ":

                        break;

                    case "УПИ В ДВЕ ЗОНИ":

                        break;

                    case "ДВЕ УПИ":

                        break;
                }

                // fhfthfth
            }
            return errorMessage;
        }

        public ProjInfoUpdater(ProjectInfo projectInfo, Document Doc)
        {
            this.ProjectInfo = projectInfo;
            this.doc = Doc;
            Transaction transaction = new Transaction(doc, "Update Project Info");
            T = transaction;
        }

        public void SetAchievedStandard(double buildArea, double grossArea, double intensity, double density)
        {
            T.Start();
            ProjectInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
            ProjectInfo.LookupParameter("Achieved Gross External Area").Set(grossArea);
            ProjectInfo.LookupParameter("Achieved Area Intensity").Set(grossArea);
            ProjectInfo.LookupParameter("Achieved Built up Density").Set(density);
            T.Commit();
        }
        
        public void SetRequired(double buildArea, double grossArea, double intensity, double density)
        {
            T.Start();
            ProjectInfo.LookupParameter("Required Build up Area").Set(buildArea);
            ProjectInfo.LookupParameter("Required Gross External Area").Set(grossArea);
            ProjectInfo.LookupParameter("Required Area Intensity").Set(intensity);
            ProjectInfo.LookupParameter("Required Built up Density").Set(density);
            T.Commit();
        }

        public void SetAllTwoZones(double plotArea, double buildArea, double totalBuild, double kint, double density)
        {
            T.Start();
            ProjectInfo.LookupParameter("Plot Area").Set(plotArea);
            ProjectInfo.LookupParameter("Required Built up Density")
                .Set(ProjectInfo.LookupParameter("Required Built up Density 1st").AsDouble() + ProjectInfo.LookupParameter("Required Built up Density 2nd").AsDouble());
            ProjectInfo.LookupParameter("Required Built up Area")
                .Set(ProjectInfo.LookupParameter("Required Built up Area 1st").AsDouble() + ProjectInfo.LookupParameter("Required Built up Area 2nd").AsDouble());
            ProjectInfo.LookupParameter("Required Area Intensity")
                .Set(ProjectInfo.LookupParameter("Required Area Intensity 1st").AsDouble() + ProjectInfo.LookupParameter("Required Area Intensity 2nd").AsDouble());
            ProjectInfo.LookupParameter("Required Gross External Area")
                .Set(ProjectInfo.LookupParameter("Required Gross External Area 1st").AsDouble() + ProjectInfo.LookupParameter("Required Gross External Area 2nd").AsDouble());
            ProjectInfo.LookupParameter("Required Green Area Percentage")
                .Set(ProjectInfo.LookupParameter("Required Green Area Percentage 1st").AsDouble() + ProjectInfo.LookupParameter("Required Green Area Percentage 2nd").AsDouble());
            ProjectInfo.LookupParameter("Required Green Area")
                .Set(ProjectInfo.LookupParameter("Required Green Area 1st").AsDouble() + ProjectInfo.LookupParameter("Required Green Area 2nd").AsDouble());
            ProjectInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
            ProjectInfo.LookupParameter("Achieved Gross External Area").Set(totalBuild);
            ProjectInfo.LookupParameter("Achieved Area Intensity").Set(kint);
            ProjectInfo.LookupParameter("Achieved Built up Density").Set(density);
            T.Commit();
        }

        public void SetAchievedTwoPlots(double buildArea1, double totalBuild1, double kint1, double density1, double buildArea2, double totalBuild2, double kint2, double density2)
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
            T.Commit();
        }
    }
}
