using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AreaCalculations
{
    internal class OutputReport
    {
        string outputString { get; set; }
        public OutputReport() 
        { outputString = "Проектните параметри бяха обновени успешно!\n"; }

        public void addString(string message)
        { outputString += message; }

        public void addPlotNamesAreas(List<double> plotAreas, List<string> plotNames)
        {

        }
    }
}
