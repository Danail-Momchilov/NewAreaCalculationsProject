using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;

namespace AreaCalculations
{
    internal class OutputReport
    {
        public string outputString { get; set; }
        public OutputReport() 
        { outputString = "Проектните параметри бяха обновени успешно!\n"; }

        public void addString(string message)
        { outputString += message; }

        public void updateFinalOutput(List<double> plotAreas, List<string> plotNames, List<double> build, List<double> density, List<double> totalBuild, List<double> kint, List<double> greenAreas, List<double> achievedPercentages)
        {
            for (int i = 0; i < plotAreas.Count; i++)
            {
                outputString += $"Площ на имот {i + 1}: " + plotAreas[i].ToString() + "  кв.м.\n" + $"Име на имот {i + 1}: " + plotNames[i] + "\n";
            }
            if (plotAreas.Count == 1)
            {
                outputString += $"Постигнато ЗП = {build[0]} кв.м.\n";
                outputString += $"Постигната плътност = {density[0]} %\n";
                outputString += $"Постигнато РЗП = {totalBuild[0]} кв.м.\n";
                outputString += $"Постигнат КИНТ = {kint[0]}\n";
                outputString += $"Постигнато озеленяване = {greenAreas[0]} кв.м.\n";
                outputString += $"Постигнат процент озеленяване = {achievedPercentages[0]} %\n";
            }
            else
            {
                outputString += $"Постигнато ЗП за имот 1 = {build[0]} кв.м.\n";
                outputString += $"Постигната плътност за имот 1 = {density[0]} %\n";
                outputString += $"Постигнато РЗП за имот 1 = {totalBuild[0]} кв.м.\n";
                outputString += $"Постигнат КИНТ за имот 1 = {kint[0]}\n";
                outputString += $"Постигнато озеленяване за имот 1 = {greenAreas[0]}  кв.м.\n";
                outputString += $"Постигнат процент озеленяване за имот 1 = {achievedPercentages[0]} %\n";
                outputString += $"Постигнато ЗП за имот 2 = {build[1]} кв.м.\n";
                outputString += $"Постигната плътност за имот 2 = {density[1]} %\n";
                outputString += $"Постигнато РЗП за имот 2 = {totalBuild[1]} кв.м.\n";
                outputString += $"Постигнат КИНТ за имот 2 = {kint[1]}\n";
                outputString += $"Постигнато озеленяване за имот 2 = {greenAreas[1]} кв.м.\n";
                outputString += $"Постигнат процент озеленяване за имот 2 = {achievedPercentages[1]} %\n";
            }
        }
    }
}
