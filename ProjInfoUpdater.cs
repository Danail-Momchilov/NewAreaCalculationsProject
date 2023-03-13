using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Autodesk.Revit.DB;

namespace AreaCalculations
{
    internal class ProjInfoUpdater
    {
        public ProjectInfo ProjectInfo { get; set; }
        Document doc { get; set; }

        Transaction T { get; set; }

        public ProjInfoUpdater(ProjectInfo projectInfo, Document Doc)
        {
            this.ProjectInfo = projectInfo;
            this.doc = Doc;
            Transaction transaction = new Transaction(doc, "Update Project Info");
            T = transaction;
        }

        public void setAchievedStandart(double buildArea, double grossArea, double intensity, double density)
        {
            T.Start();
            ProjectInfo.LookupParameter("Achieved Built up Area").Set(buildArea);
            ProjectInfo.LookupParameter("Achieved Gross External Area").Set(grossArea);
            ProjectInfo.LookupParameter("Achieved Area Intensity").Set(grossArea);
            ProjectInfo.LookupParameter("Achieved Built up Density").Set(density);
            T.Commit();
        }
    }
}
