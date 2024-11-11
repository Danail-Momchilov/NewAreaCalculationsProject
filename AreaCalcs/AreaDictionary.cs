using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Document = Autodesk.Revit.DB.Document;
using System.Runtime.InteropServices;
using System.Drawing;
using static System.Net.Mime.MediaTypeNames;
using Regex = System.Text.RegularExpressions.Regex;
using System.Windows.Input;
using System.Globalization;
using Autodesk.Revit.DB.Architecture;

namespace AreaCalculations
{
    internal class AreaDictionary
    {
        private readonly double areaConvert = 10.763914692;
        private readonly double lengthConvert = 30.48;
        public Dictionary<string, Dictionary<string, List<Area>>> AreasOrganizer { get; set; }
        public List<string> plotNames { get; set; }
        public Dictionary<string, Dictionary<string, double>> propertyCommonAreas { get; set; }
        public Dictionary<string, Dictionary<string, double>> propertyCommonAreasSpecial { get; set; }
        public Dictionary<string, Dictionary<string, double>> propertyCommonAreasAll { get; set; }
        public Dictionary<string, Dictionary<string, double>> propertyIndividualAreas { get; set; }
        public Dictionary<string, double> plotAreasImp { get; set; } // IMPORTANT!!! STORING DATA IN IMPERIAL
        public Dictionary<string, List<string>> plotProperties { get; set; }
        public Dictionary<string, double> plotBuildAreas { get; set; }
        public Dictionary<string, double> plotTotalBuild { get; set; }
        public Dictionary<string, double> plotUndergroundAreas { get; set; }
        public Dictionary<string, double> plotIndividualAreas { get; set; }
        public Dictionary<string, double> plotCommonAreas { get; set; }
        public Dictionary<string, double> plotCommonAreasSpecial { get; set; }
        public Dictionary<string, double> plotCommonAreasAll { get; set; }
        public Dictionary<string, double> plotLandAreas { get; set; }
        public Document doc { get; set; }
        public Transaction transaction { get; set; }
        public double areasCount { get; set; }
        public double missingAreasCount { get; set; }
        public string missingAreasData { get; set; }
        public AreaDictionary(Document activeDoc)
        {
            this.doc = activeDoc;
            this.AreasOrganizer = new Dictionary<string, Dictionary<string, List<Area>>>();
            this.propertyCommonAreas = new Dictionary<string, Dictionary<string, double>>();
            this.propertyCommonAreasSpecial = new Dictionary<string, Dictionary<string, double>>();
            this.propertyCommonAreasAll = new Dictionary<string, Dictionary<string, double>>();
            this.propertyIndividualAreas = new Dictionary<string, Dictionary<string, double>>();
            this.plotNames = new List<string>();
            this.plotProperties = new Dictionary<string, List<string>>();
            this.transaction = new Transaction(activeDoc, "Calculate and Update Area Parameters");
            this.plotAreasImp = new Dictionary<string, double>();
            this.plotBuildAreas = new Dictionary<string, double>();
            this.plotTotalBuild = new Dictionary<string, double>();
            this.plotUndergroundAreas = new Dictionary<string, double>();
            this.plotIndividualAreas = new Dictionary<string, double>();
            this.plotCommonAreas = new Dictionary<string, double>();
            this.plotCommonAreasSpecial = new Dictionary<string, double>();
            this.plotCommonAreasAll = new Dictionary<string, double>();
            this.plotLandAreas = new Dictionary<string, double>();
            this.areasCount = 0;
            this.missingAreasCount = 0;

            ProjectInfo projectInfo = activeDoc.ProjectInformation;

            FilteredElementCollector areasCollector = new FilteredElementCollector(activeDoc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();

            // construct main AreaOrganizer Dictionary
            foreach (Element elem in areasCollector)
            {
               Area area = elem as Area;

                string plotName = area.LookupParameter("A Instance Area Plot").AsString();
                string groupName = area.LookupParameter("A Instance Area Group").AsString();

                if (!string.IsNullOrEmpty(plotName) && !string.IsNullOrEmpty(groupName) && area.Area!=0)
                {
                    if (!AreasOrganizer.ContainsKey(plotName))
                    {
                        this.AreasOrganizer.Add(plotName, new Dictionary<string, List<Area>>());
                        this.plotNames.Add(plotName);
                        this.plotProperties.Add(plotName, new List<string>());

                        if (projectInfo.LookupParameter("Plot Number").AsString() == plotName)
                            this.plotAreasImp.Add(plotName, projectInfo.LookupParameter("Plot Area").AsDouble());
                        else if (projectInfo.LookupParameter("Plot Number 1st").AsString() == plotName)
                            this.plotAreasImp.Add(plotName, projectInfo.LookupParameter("Plot Area 1st").AsDouble());
                        else if (projectInfo.LookupParameter("Plot Number 2nd").AsString() == plotName)
                            this.plotAreasImp.Add(plotName, projectInfo.LookupParameter("Plot Area 2nd").AsDouble());
                    }

                    if (!AreasOrganizer[plotName].ContainsKey(groupName))
                    {
                        this.AreasOrganizer[plotName].Add(groupName, new List<Area>());
                        this.plotProperties[plotName].Add(groupName);
                    }

                    this.AreasOrganizer[plotName][groupName].Add(area);
                    areasCount++;
                }

                else
                {
                    // TODO: check whether an Area has a plotName, that is not relative to the ones in the Project Info (is it actually needed?)
                    // TODO
                    // TODO
                    missingAreasCount++;
                    missingAreasData += $"{area.Id} {area.Number} {area.Name} {area.Area}\n";
                    // TODO
                    // TODO
                    // TODO
                }
            }

            // based on AreasOrganizer, construct the plotBuildAreas dictionary
            foreach (string plotName in plotNames)
            {
                plotBuildAreas.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Location").AsString() == "НАЗЕМНА")
                        {
                            plotBuildAreas[plotName] += Math.Round(area.Area / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the plotTotalBuild dictionary
            foreach (string plotName in plotNames)
            {
                plotTotalBuild.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if ((area.LookupParameter("A Instance Area Location").AsString().ToLower() == "надземна" 
                            || area.LookupParameter("A Instance Area Location").AsString().ToLower() == "наземна")
                            && area.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект"
                            && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            plotTotalBuild[plotName] += Math.Round(area.LookupParameter("A Instance Total Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the plotUndergroundAreas dictionary
            foreach (string plotName in plotNames)
            {
                plotUndergroundAreas.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Location").AsString() == "ПОДЗЕМНА")
                        {
                            plotUndergroundAreas[plotName] += Math.Round(area.LookupParameter("A Instance Total Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the plotIndividualAreas dictionary
            foreach (string plotName in plotNames)
            {
                plotIndividualAreas.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект" 
                            && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            plotIndividualAreas[plotName] += Math.Round(area.LookupParameter("A Instance Gross Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the plotCommonAreas dictionary
            foreach (string plotName in plotNames)
            {
                plotCommonAreas.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "обща част")
                        {
                            plotCommonAreas[plotName] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the plotCommonAreasSpecial dictionary
            foreach (string plotName in plotNames)
            {
                plotCommonAreasSpecial.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "обща част"
                            && (area.LookupParameter("A Instance Area Primary").HasValue
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            plotCommonAreasSpecial[plotName] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the plotCommonAreasAll dictionary
            foreach (string plotName in plotNames)
            {
                plotCommonAreasAll.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "обща част")
                        {
                            plotCommonAreasAll[plotName] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the propertyCommonAreas dictionary
            foreach (string plotName in plotNames)
            {
                propertyCommonAreas.Add(plotName, new Dictionary<string, double>());

                foreach (string plotProperty in plotProperties[plotName])
                {
                    propertyCommonAreas[plotName].Add(plotProperty, 0);

                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "обща част"
                            && !(area.LookupParameter("A Instance Area Primary").HasValue
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            propertyCommonAreas[plotName][plotProperty] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }

                // check if there are "A + B + C" type of properties and if there are any, get the total sum of their common areas
                double plusPropertiesCommonSum = 0;
                bool wasFound = false;

                foreach (string plotProperty in plotProperties[plotName])
                {
                    if (plotProperty.Contains("+"))
                    {
                        plusPropertiesCommonSum += propertyCommonAreas[plotName][plotProperty];
                        wasFound = true;
                    }
                }

                // in case such property types were found, redistribute their areas across the rest of the properties
                if (wasFound)
                {
                    double remainingCommonArea = plotCommonAreas[plotName] - plusPropertiesCommonSum;

                    // search for properties of type "A + B"
                    foreach (string plotProperty in plotProperties[plotName])
                    {
                        if (plotProperty.Contains("+"))
                        {
                            // if such is found, redistribute their areas across the rest of the properties
                            string[] splitProperties = plotProperty.Split(new char[] { '+' }, StringSplitOptions.RemoveEmptyEntries).Select(s => s.Trim()).ToArray();

                            foreach (string property in splitProperties)
                            {
                                double ratio = Math.Round(propertyCommonAreas[plotName][property] / remainingCommonArea, 3);

                                double areaToAdd = Math.Round(propertyCommonAreas[plotName][plotProperty] * ratio, 3);
                                propertyCommonAreas[plotName][property] += areaToAdd;
                            }
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the propertyCommonAreasSpecial dictionary
            foreach (string plotName in plotNames)
            {
                propertyCommonAreasSpecial.Add(plotName, new Dictionary<string, double>());

                foreach (string plotProperty in plotProperties[plotName])
                {
                    propertyCommonAreasSpecial[plotName].Add(plotProperty, 0);

                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "обща част"
                            && (area.LookupParameter("A Instance Area Primary").HasValue
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            propertyCommonAreasSpecial[plotName][plotProperty] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on Areas Organizer, construct the propertyCommonAreasAll
            foreach (string plotName in plotNames)
            {
                propertyCommonAreasAll.Add(plotName, new Dictionary<string, double>());

                foreach (string plotProperty in plotProperties[plotName])
                {
                    double sum = propertyCommonAreas[plotName][plotProperty] + propertyCommonAreasSpecial[plotName][plotProperty];

                    propertyCommonAreasAll[plotName].Add(plotProperty, sum);
                }
            }

            // based on AreasOrganizer, construct the propertyIndividualAreas dictionary
            foreach (string plotName in plotNames)
            {
                propertyIndividualAreas.Add(plotName, new Dictionary<string, double>());

                foreach (string plotProperty in plotProperties[plotName])
                {
                    propertyIndividualAreas[plotName].Add(plotProperty, 0);

                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект"
                            && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            propertyIndividualAreas[plotName][plotProperty] += Math.Round(area.LookupParameter("A Instance Gross Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }

            // based on AreasOrganizer, construct the plotLandAreas dictionary
            foreach (string plotName in plotNames)
            {
                plotLandAreas.Add(plotName, 0);

                foreach (string plotProperty in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][plotProperty])
                    {
                        if (area.LookupParameter("A Instance Area Group").AsString().ToLower() == "земя")
                        {
                            plotLandAreas[plotName] += Math.Round(area.LookupParameter("Area").AsDouble() / areaConvert, 2);
                        }
                    }
                }
            }
        }
        private int ExtractLevelNumber(string levelString)
        {
            if (!string.IsNullOrEmpty(levelString))
            {
                var match = Regex.Match(levelString, @"\d+");
                return match.Success ? int.Parse(match.Value) : int.MaxValue;
            }
            else
            {
                return 0;
            }            
        }
        private string ReorderEntrance(string entranceName)
        {
            if (entranceName == "НЕПРИЛОЖИМО")
                return "A";
            else
                return entranceName;
        }
        private void calculateSurplusPercent(Dictionary<string, List<Area>> areaGroup, string parameterName)
        {
            if (areaGroup.Keys.ToList().Count == 0)
            {
                return;
            }

            transaction.Start();

            // calculate building permit surplus
            double buildingPermitTotal = 0;
            foreach (string group in areaGroup.Keys)
            {
                foreach (Area area in areaGroup[group])
                {
                    buildingPermitTotal += Math.Round(area.LookupParameter(parameterName).AsDouble(), 3);
                }
            }

            double surplus = Math.Round(100 - buildingPermitTotal, 3);
            int counter = 0;

            while (Math.Abs(surplus) >= 0.0005)
            {
                if (counter >= areaGroup.Keys.ToList().Count())
                {
                    counter = 0;
                }
                string group = areaGroup.Keys.ToList()[counter];

                if (Math.Round(Math.Abs(surplus) * 1000) >= areaGroup[group].Count())
                {
                    // calculate the deduction total, depending on whether the surplis is positive or negatibe
                    double coefficient = surplus / Math.Abs(surplus);
                    // so far, the result is either 1 or -1
                    double finalDeduction = 0.001 * coefficient;
                    // this is the final deduction calculated value, which would be either -0.001 or 0.001

                    foreach (Area area in areaGroup[group])
                    {
                        // redistribute final deduction towards the Building Permit %
                        double currentPercent = area.LookupParameter(parameterName).AsDouble();
                        area.LookupParameter(parameterName).Set(currentPercent + finalDeduction);
                        surplus -= finalDeduction;
                    }
                }

                counter++;

                if (counter == 10)
                    surplus = 0;
            }

            transaction.Commit();
        }
        private void calculateSurplusPercentandArea(Dictionary<string, List<Area>> areaGroup, string parameterNamePercent, string parameterNameArea, double totalAreaToCalculateFrom)
        {
            if (areaGroup.Keys.ToList().Count == 0)
            {
                return;
            }

            transaction.Start();

            // calculate building permit surplus
            double buildingPermitTotal = 0;
            foreach (string group in areaGroup.Keys)
            {
                foreach (Area area in areaGroup[group])
                {
                    buildingPermitTotal += Math.Round(area.LookupParameter(parameterNamePercent).AsDouble(), 3);
                }
            }

            double surplus = Math.Round(100 - buildingPermitTotal, 3);
            int counter = 0;
            int numbOfCycles = 0;

            while (Math.Abs(surplus) >= 0.0005)
            {
                if (numbOfCycles == 5)
                {
                    string randomGroup = areaGroup.Keys.ToList()[0];
                    Area randomArea = areaGroup[randomGroup][0];
                    string areasInfo = $"plot: {randomArea.LookupParameter("A Instance Area Plot").AsString()} and area group: {randomArea.LookupParameter("A Instance Area Group").AsString()}";

                    TaskDialog.Show("Warning", $"Surplus could not be distributed for parameters {parameterNamePercent} and {parameterNameArea} for {areasInfo}");

                    transaction.Commit();
                    return;
                }
                if (counter >= areaGroup.Keys.ToList().Count())
                {
                    counter = 0;
                    numbOfCycles++;
                }

                string group = areaGroup.Keys.ToList()[counter];

                if (Math.Round(Math.Abs(surplus) * 1000) >= areaGroup[group].Count())
                {
                    // calculate the deduction total, depending on whether the surplis is positive or negatibe
                    double coefficient = surplus / Math.Abs(surplus);
                    // so far, the result is either 1 or -1
                    double finalDeduction = 0.001 * coefficient;
                    // this is the final deduction calculated value, which would be either -0.001 or 0.001

                    foreach (Area area in areaGroup[group])
                    {
                        // redistribute final deduction towards the Building Permit %
                        double calculatedPercent = area.LookupParameter(parameterNamePercent).AsDouble() + finalDeduction;
                        area.LookupParameter(parameterNamePercent).Set(calculatedPercent);
                        surplus -= finalDeduction;

                        // calculate the updated area 
                        double calculatedArea = Math.Round(calculatedPercent * totalAreaToCalculateFrom / 100, 2);

                        // redistribute square meters area accordingly
                        area.LookupParameter(parameterNameArea).Set(calculatedArea * areaConvert);
                    }
                }

                counter++;
            }

            // check if the area was calculated accordingly or if there is a newly formed surplus
            double totalAreaSum = 0;
            foreach (string group in areaGroup.Keys)
            {
                foreach (Area area in areaGroup[group])
                {
                    totalAreaSum += Math.Round(area.LookupParameter(parameterNameArea).AsDouble() / areaConvert, 2);
                }
            }

            double areaSurplus = Math.Round(totalAreaToCalculateFrom - totalAreaSum, 2);
            int areaCounter = 0;
            int numbOfAreaCycles = 0;

            // check if areasurplus is 0, and if yes, redistribute it accordingly as well
            while (Math.Abs(areaSurplus) >= 0.005)
            {
                if (numbOfAreaCycles == 5)
                {
                    string randomGroup = areaGroup.Keys.ToList()[0];
                    Area randomArea = areaGroup[randomGroup][0];
                    string areasInfo = $"plot: {randomArea.LookupParameter("A Instance Area Plot").AsString()} and area group: {randomArea.LookupParameter("A Instance Area Group").AsString()}";

                    TaskDialog.Show("Warning", $"Surplus area could not be distributed for parameters {parameterNameArea} for {areasInfo}. " +
                        $"// Remaining surplus is {areaSurplus} | Number of cycles: {numbOfAreaCycles}");

                    transaction.Commit();
                    return;
                }

                if (areaCounter >= areaGroup.Keys.ToList().Count())
                {
                    areaCounter = 0;
                    numbOfAreaCycles++;
                }

                string group = areaGroup.Keys.ToList()[areaCounter];

                if (Math.Round(Math.Abs(areaSurplus) * 100) >= areaGroup[group].Count())
                {
                    // calculate the deduction total, depending on whether the surplis is positive or negatibe
                    double coefficient = areaSurplus / Math.Abs(areaSurplus);
                    // so far, the result is either 1 or -1
                    double finalDeduction = 0.01 * coefficient;
                    // this is the final deduction calculated value, which would be either -0.01 or 0.01

                    foreach (Area area in areaGroup[group])
                    {
                        // calculate the updated area
                        double calculatedArea = Math.Round(area.LookupParameter(parameterNameArea).AsDouble() / areaConvert + finalDeduction, 2);
                        area.LookupParameter(parameterNameArea).Set(Math.Round(calculatedArea * areaConvert, 2));

                        areaSurplus -= finalDeduction;
                    }
                }

                areaCounter++;
            }
            
            transaction.Commit();
        }
        private void calculateSpecialCommonAreaSurplus(Dictionary<string, List<Area>> areaGroup, string parameterNameArea, double totalAreaToCalculateFrom)
        {
            if (areaGroup.Keys.ToList().Count == 0)
            {
                return;
            }

            transaction.Start();

            double totalCalculatedArea = 0;
            foreach (string group in areaGroup.Keys)
            {
                foreach (Area area in areaGroup[group])
                {
                    totalCalculatedArea += Math.Round(area.LookupParameter(parameterNameArea).AsDouble() / areaConvert, 2);
                }
            }

            double surplus = Math.Round(totalAreaToCalculateFrom - totalCalculatedArea, 2);
            int counter = 0;
            int numbOfCycles = 0;

            while (Math.Abs(surplus) >= 0.0005)
            {
                if (numbOfCycles == 5)
                {
                    string randomGroup = areaGroup.Keys.ToList()[0];
                    Area randomArea = areaGroup[randomGroup][0];
                    string areasInfo = $"plot: {randomArea.LookupParameter("A Instance Area Plot").AsString()} and area group: {randomArea.LookupParameter("A Instance Area Group").AsString()}";

                    TaskDialog.Show("Warning", $"Surplus could not be distributed for parameter {parameterNameArea}, {areasInfo}");

                    transaction.Commit();
                    return;
                }
                if (counter >= areaGroup.Keys.ToList().Count())
                {
                    counter = 0;
                    numbOfCycles++;
                }

                string group = areaGroup.Keys.ToList()[counter];

                // calculate the deduction total, depending on whether the surplis is positive or negatibe
                double coefficient = surplus / Math.Abs(surplus);
                // so far, the result is either 1 or -1
                double finalDeduction = 0.01 * coefficient;
                // this is the final deduction calculated value, which would be either -0.01 or 0.01

                foreach (Area area in areaGroup[group])
                {
                    if (area.LookupParameter("A Instance Common Area Special").HasValue &&
                        Math.Round(area.LookupParameter("A Instance Common Area Special").AsDouble() / areaConvert, 2) != 0 &&
                        area.LookupParameter("A Instance Common Area Special").AsString() != "")
                    {
                        // calculate the updated area
                        double calculatedArea = Math.Round(area.LookupParameter(parameterNameArea).AsDouble() / areaConvert + finalDeduction, 2);
                        area.LookupParameter(parameterNameArea).Set(Math.Round(calculatedArea * areaConvert, 2));

                        surplus -= finalDeduction;
                    }
                }

                counter++;
            }

            transaction.Commit();
        }
        private bool doesHaveRoomsAdjascent(string areaNumber)
        {
            bool hasAdjRooms = false;

            List<Element> rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType().ToList();

            foreach (Element element in rooms)
            {
                Room room = element as Room;

                if (room.LookupParameter("A Instance Area Primary").AsString() == areaNumber)
                {
                    hasAdjRooms = true;
                }
            }

            return hasAdjRooms;
        }
        private List<Room> returnAdjascentRooms(string areaNumber)
        {
            List<Room> rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
                .Where(room => room.LookupParameter("A Instance Area Primary").AsString() == areaNumber).Cast<Room>().ToList();

            rooms.OrderBy(room => room.LookupParameter("Number").AsString());

            return rooms;
        }        
        private Dictionary<List<object>, Room> returnAdjascentRoomsTest(string areaNumber, double areaArea)
        {
            Dictionary<string, List<object>> keyValuePairs = new Dictionary<string, List<object>>();

            List<Room> rooms = new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Rooms).WhereElementIsNotElementType()
            .Where(room => room.LookupParameter("A Instance Area Primary").AsString() == areaNumber).Cast<Room>().ToList();

            // group rooms, based on their Area
            var groupedRooms = rooms.GroupBy(room => Math.Round(room.LookupParameter("Area").AsDouble(), 3));

            // create a dictionary from these groups
            Dictionary<List<double>, List<Room>> percentageDict = new Dictionary<List<double>, List<Room>>();
            double totalpercentage = 0;

            foreach (var group in groupedRooms)
            {
                List<double> listData = new List<double>();
                listData.Add(group.Count());

                double percentage = Math.Round(group.First().LookupParameter("Area").AsDouble()*100/areaArea, 3);
                listData.Add(percentage);

                percentageDict.Add(listData, new List<Room>());

                foreach (Room room in group)
                {
                    percentageDict[listData].Add(room);
                    totalpercentage += percentage;
                }
            }

            // calculate surplus
            double surplus = Math.Round(100 - totalpercentage, 3);
            int counter = 0;

            // redistribute surplus if any
            while (Math.Abs(surplus) >= 0.0005)
            {
                if (counter >= percentageDict.Keys.Count())
                {
                    counter = 0;
                }
                List<double> dictList = percentageDict.Keys.ToList()[counter];

                if (Math.Round(Math.Abs(surplus) * 1000) >= percentageDict[dictList].Count)
                {
                    // calculate the deduction total, depending on whether the surplis is positive or negatibe
                    double coefficient = surplus / Math.Abs(surplus);
                    // so far, the result is either 1 or -1
                    double finalDeduction = 0.001 * coefficient;
                    // this is the final deduction calculated value, which would be either -0.001 or 0.001

                    percentageDict.Keys.ToList()[counter][1] += finalDeduction;

                    foreach (Room room in percentageDict[dictList])
                    {
                        surplus -= finalDeduction;
                    }
                }
            }

            // construct new dictionary
            Dictionary<List<object>, Room> flattenedDict = new Dictionary<List<object>, Room>();

            // tterate through the percentageDict to populate the new dictionary
            foreach (var kvp in percentageDict)
            {
                List<double> keyList = kvp.Key;
                List<Room> roomsList = kvp.Value;

                double percentage = keyList[1];

                foreach (Room room in roomsList)
                {
                    // Create the new key with the room number and percentage
                    List<object> newKey = new List<object>
                    {
                        room.LookupParameter("Number").AsString(), percentage
                    };

                    // Add to the new dictionary
                    flattenedDict[newKey] = room;
                }
            }

            // sort the flattened dictionary based on room number
            var sortedFlattenedDict = flattenedDict.OrderBy(kvp => (string)kvp.Key[0]).ToDictionary(kvp => kvp.Key, kvp => kvp.Value); 

            return sortedFlattenedDict;
        }
        
        public void setGrossArea()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        area.LookupParameter("A Instance Gross Area").Set(area.LookupParameter("Area").AsDouble());
                    }
                }
            }

            transaction.Commit();
        }
        public string calculatePrimaryArea()
        {
            string errorMessage = "";
            List<string> missingNumbers = new List<string>();

            transaction.Start();

            foreach (string plotName in plotNames)
            {
                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area secArea in AreasOrganizer[plotName][property])
                    {
                        if (secArea.LookupParameter("A Instance Area Primary").HasValue && secArea.LookupParameter("A Instance Area Primary").AsString() != "" && secArea.Area != 0)
                        {
                            bool wasFound = false;

                            string[] mainAreaNumbers = secArea.LookupParameter("A Instance Area Primary").AsString().Split(new char[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => s.Trim())
                                .ToArray();

                            foreach (string mainAreaNumber in mainAreaNumbers)
                            {
                                if (secArea.LookupParameter("A Instance Area Group").AsString().ToLower() != "земя")
                                {
                                    foreach (Area mainArea in AreasOrganizer[plotName][property])
                                    {
                                        if (mainArea.Number == mainAreaNumber)
                                        {
                                            wasFound = true;

                                            if (secArea.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект")
                                            {
                                                mainArea.LookupParameter("A Instance Gross Area").Set(
                                                    mainArea.LookupParameter("A Instance Gross Area").AsDouble() + secArea.Area);
                                            }
                                        }
                                    }
                                }
                                
                                else
                                {
                                    foreach (string plotNameMain in plotNames)
                                    {
                                        foreach (string propertyMain in plotProperties[plotNameMain])
                                        {
                                            foreach (Area mainArea in AreasOrganizer[plotNameMain][propertyMain])
                                            {
                                                if (mainArea.Number == mainAreaNumber)
                                                {
                                                    wasFound = true;

                                                    if (secArea.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект")
                                                    {
                                                        mainArea.LookupParameter("A Instance Gross Area").Set(
                                                            mainArea.LookupParameter("A Instance Gross Area").AsDouble() + secArea.Area);
                                                    }
                                                }
                                            }
                                        }
                                    }
                                }
                            }

                            if (!wasFound && !missingNumbers.Contains(secArea.LookupParameter("Number").AsString()))
                            {
                                missingNumbers.Add(secArea.LookupParameter("Number").AsString());
                                errorMessage += $"Грешка: Area {secArea.LookupParameter("Number").AsString()} / id: {secArea.Id} " +
                                    $"/ Посочената Area е зададена като подчинена на такава с несъществуващ номер. Моля, проверете я и стартирайте апликацията отново\n";
                            }
                        }
                    }
                }
            }

            transaction.Commit();

            return errorMessage;
        }
        public void calculateC1C2()
        {
            transaction.Start();

            foreach (Element collectorElement in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
            {
                Area area = collectorElement as Area;

                if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ"
                    && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                {
                    area.LookupParameter("A Instance Price C1/C2").Set(area.LookupParameter("A Instance Gross Area").AsDouble()
                        * area.LookupParameter("A Coefficient Multiplied").AsDouble() / areaConvert);
                }
            }

            transaction.Commit();
        }
        public void calculateCommonAreaPerc()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                foreach (string property in plotProperties[plotName])
                {
                    double totalC1C2 = 0;

                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue 
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double C1C2 = area.LookupParameter("A Instance Price C1/C2").AsDouble();
                            totalC1C2 += C1C2;
                        }
                    }

                    // calculate common area percentage parameter for each area
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue 
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double commonAreaPercent = (area.LookupParameter("A Instance Price C1/C2").AsDouble() / totalC1C2) * 100;
                            area.LookupParameter("A Instance Common Area %").Set(commonAreaPercent);
                        }
                    }
                }
            }

            transaction.Commit();
        }
        public void calculateCommonArea()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                foreach (string property in plotProperties[plotName])
                {
                    double propertyCommonArea = propertyCommonAreas[plotName][property];

                    // calculate common area percentage parameter for each area
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ"
                            && !(area.LookupParameter("A Instance Area Primary").HasValue 
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double commonArea;

                            commonArea = Math.Round(area.LookupParameter("A Instance Common Area %").AsDouble() * propertyCommonArea / 100 * areaConvert, 2);
                            area.LookupParameter("A Instance Common Area").Set(commonArea);
                        }
                    }
                }
            }

            transaction.Commit();
        }
        public void calculateSpecialCommonAreas()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                foreach (string property in plotProperties[plotName])
                {
                    // first set all special common area values to 0
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        area.LookupParameter("A Instance Common Area Special").Set(0);
                    }

                    // check all areas from a given dictionary
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        // check if there is a common area that is set to adjascent to another one
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "обща част" &&
                            area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != "")
                        {
                            // if such is found, find all of the areas, it is set to be adjascent to
                            string[] mainAreaNames = area.LookupParameter("A Instance Area Primary").AsString().Split(new char[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                                .Select(s => s.Trim())
                                .ToArray();

                            double sumC1C2 = 0;
                            List<Area> mainAreaElements = new List<Area>();

                            // find all the areas it is adjascent to and calculate their total C1C2 and add them to a list
                            foreach (string mainAreaName in mainAreaNames)
                            {
                                foreach (Area mainArea in AreasOrganizer[plotName][property])
                                {
                                    if (mainArea.LookupParameter("Number").AsString() == mainAreaName)
                                    {
                                        sumC1C2 += mainArea.LookupParameter("A Instance Price C1/C2").AsDouble();
                                        mainAreaElements.Add(mainArea);
                                    }
                                }
                            }

                            // for each area of the list, calculate its Special Common Area
                            foreach (Area mainArea in mainAreaElements)
                            {
                                double percentage = mainArea.LookupParameter("A Instance Price C1/C2").AsDouble() * 100 / sumC1C2;

                                mainArea.LookupParameter("A Instance Common Area Special")
                                    .Set(mainArea.LookupParameter("A Instance Common Area Special").AsDouble() + (percentage * area.Area / 100));
                            }
                        }
                    }
                }
            }

            transaction.Commit();
        }
        public void calculateTotalArea()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                foreach (string property in plotProperties[plotName])
                {                 
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue 
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            area.LookupParameter("A Instance Total Area").Set(area.LookupParameter("A Instance Gross Area").AsDouble() 
                                + area.LookupParameter("A Instance Common Area").AsDouble());
                        }
                    }
                }
            }

            transaction.Commit();
        }
        public void calculateBuildingPercentPermit()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                double totalPlotC1C2 = 0;

                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" 
                            && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            totalPlotC1C2 += area.LookupParameter("A Instance Price C1/C2").AsDouble();
                        }
                    }
                }

                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" 
                            && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double buildingPercentPermit = (area.LookupParameter("A Instance Price C1/C2").AsDouble() / totalPlotC1C2) * 100;
                            area.LookupParameter("A Instance Building Permit %").Set(buildingPercentPermit);
                        }
                    }
                }
            }

            transaction.Commit();
        }
        public void calculateRlpAreaPercent()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                double plotArea = plotAreasImp[plotName];

                // find all area objects of type 'ЗЕМЯ' and calculate their collective percentage as of the total plot area
                double reductionPercentage = 0;

                // also find the total C1C2 for all individual areas in the property
                double totalC1C2 = 0;

                // calculate the total C1C2 and reduction percentage for the whole plot
                foreach (string property in plotProperties[plotName])
                {                    
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Group").AsString().ToLower() == "земя")
                        {
                            double areaPercentage = 100 * area.LookupParameter("Area").AsDouble() / plotArea;
                            area.LookupParameter("A Instance RLP Area %").Set(areaPercentage);
                            reductionPercentage += areaPercentage;
                        }

                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект"
                            && !(area.LookupParameter("A Instance Area Primary").HasValue 
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            totalC1C2 += area.LookupParameter("A Instance Price C1/C2").AsDouble();
                        }
                    }
                }

                // calculate rlp area percent for each individual area
                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект"
                            && !(area.LookupParameter("A Instance Area Primary").HasValue
                            && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double rlpAreaPercent = area.LookupParameter("A Instance Price C1/C2").AsDouble() * 100 / totalC1C2;

                            rlpAreaPercent = rlpAreaPercent * (100 - reductionPercentage) / 100;

                            area.LookupParameter("A Instance RLP Area %").Set(rlpAreaPercent);
                        }
                    }
                }
            }

            transaction.Commit();
        }
        public void calculateRlpArea()
        {
            transaction.Start();

            foreach (string plotName in plotNames)
            {
                // reduce all the areas of 'ЗЕМЯ' objects from the total plot area
                double remainingPlotArea = plotAreasImp[plotName];

                // calculate the actual RLP area for each Area object
                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        area.LookupParameter("A Instance RLP Area").Set(area.LookupParameter("A Instance RLP Area %").AsDouble() * remainingPlotArea / 100);
                    }
                }
            }

            transaction.Commit();
        }
        public string calculateInstancePropertyCommonAreaPercentage()
        {
            transaction.Start();

            string errorReport = "";

            foreach (string plotName in plotNames)
            {
                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue 
                            && area.LookupParameter("A Instance Area Primary").AsString() != "") && area.Area != 0)
                        {
                            try
                            {
                                double commonAreaImp = area.LookupParameter("A Instance Common Area").AsDouble();
                                double totalAreaImp = area.LookupParameter("A Instance Total Area").AsDouble();

                                double commonAreaPercent = (commonAreaImp * 100) / totalAreaImp;

                                area.LookupParameter("A Instance Property Common Area %").Set(commonAreaPercent);
                            }
                            catch
                            {
                                errorReport += $"{area.Id} {area.Name} A Instance Common Area = {area.LookupParameter("A Instance Common Area").AsDouble()} " +
                                    $"/ A Instance Total Area = {area.LookupParameter("A Instance Total Area").AsDouble()}";
                            }
                        }
                    }
                }
            }

            transaction.Commit();

            return errorReport;
        }
        public void redistributeSurplus()
        {
            foreach (string plotName in plotNames)
            {
                Dictionary<string, List<Area>> areaGroupsAll = new Dictionary<string, List<Area>>();
                Dictionary<string, List<Area>> areaGroupsNoLand = new Dictionary<string, List<Area>>();
                Dictionary<string, Dictionary<string, List<Area>>> areaGroupsSeperateProperties = new Dictionary<string, Dictionary<string, List<Area>>>();

                List<Area> plotAreasNoLand = new List<Area>();
                List<Area> plotAreasAll = new List<Area>();

                // get all the areas in proper lists
                foreach (string property in plotProperties[plotName])
                {
                    if (property.ToLower() != "земя" && property.ToLower() != "траф" && !property.ToLower().Contains('+'))
                    {
                        plotAreasNoLand.AddRange(AreasOrganizer[plotName][property]);
                        plotAreasAll.AddRange(AreasOrganizer[plotName][property]);
                    }
                    else if (property.ToLower() == "земя")
                    {
                        plotAreasAll.AddRange(AreasOrganizer[plotName][property]);
                    }

                    // construct the dictionary for area surplus redistribution
                    List<Area> propertyGroupAreas = new List<Area>();

                    if (property.ToLower() != "земя" && property.ToLower() != "траф" && !property.ToLower().Contains('+'))
                    {
                        foreach (Area area in AreasOrganizer[plotName][property])
                        {
                            if (area.LookupParameter("A Instance Area Category").AsString().ToLower() != "обща част" && 
                                area.LookupParameter("A Instance Area Category").AsString().ToLower() != "изключена от оч" &&
                                !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                            {
                                propertyGroupAreas.Add(area);
                            }
                        }
                    }

                    List<Area> orderedPropertyGroup = propertyGroupAreas
                        .OrderBy(area => area.LookupParameter("A Instance Gross Area").AsDouble())
                        .Reverse()
                        .ToList();
                    var groupedAreasProperty = orderedPropertyGroup.GroupBy(area => area.LookupParameter("A Instance Gross Area").AsDouble());

                    areaGroupsSeperateProperties.Add(property, new Dictionary<string, List<Area>>());

                    int sequenceproperty = 1;

                    foreach (var group in groupedAreasProperty)
                    {
                        int areaCount = group.Count();
                        string key = $"{areaCount}N{sequenceproperty}";
                        areaGroupsSeperateProperties[property].Add(key, group.ToList());
                        sequenceproperty++;
                    }
                }

                // sort lists, based on objects' areas
                List<Area> sortedAreasNoLand = plotAreasNoLand
                    .OrderBy(area => area.LookupParameter("A Instance Gross Area").AsDouble())
                    .Where(area => area.LookupParameter("A Instance Area Category").AsString().ToLower() != "обща част")
                    .Where(area => area.LookupParameter("A Instance Area Category").AsString().ToLower() != "изключена от оч")
                    .Where(area => !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                    .Reverse()
                    .ToList();

                List<Area> sortedAreasAll = plotAreasAll
                    .OrderBy(area => area.LookupParameter("A Instance Gross Area").AsDouble())
                    .Where(area => area.LookupParameter("A Instance Area Category").AsString().ToLower() != "обща част")
                    .Where(area => area.LookupParameter("A Instance Area Category").AsString().ToLower() != "изключена от оч")
                    .Where(area => !(area.LookupParameter("A Instance Area Category").AsString().ToLower() == "самостоятелен обект" && 
                           area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                    .Reverse()
                    .ToList();

                // group areas by their "A Instance Gross Area" value
                var groupedAreasAll = sortedAreasAll.GroupBy(area => area.LookupParameter("A Instance Gross Area").AsDouble());
                var groupedAreasNoLand = sortedAreasNoLand.GroupBy(area => area.LookupParameter("A Instance Gross Area").AsDouble());

                // construct the dictionary for all area groups
                int sequenceAll = 1;

                foreach (var group in groupedAreasAll)
                {
                    int areaCount = group.Count();
                    string key = $"{areaCount}N{sequenceAll}";
                    areaGroupsAll[key] = group.ToList();
                    sequenceAll++;
                }

                // construct the dictionary for all area groups, except for the land
                int sequence = 1;

                foreach (var group in groupedAreasNoLand)
                {
                    int areaCount = group.Count();
                    string key = $"{areaCount}N{sequence}";
                    areaGroupsNoLand[key] = group.ToList();
                    sequence++;
                }
                
                // calculate building permit surplus
                calculateSurplusPercent(areaGroupsNoLand, "A Instance Building Permit %");
                // calculate RLP area percent and RLP Area
                calculateSurplusPercentandArea(areaGroupsAll, "A Instance RLP Area %", "A Instance RLP Area", Math.Round(plotAreasImp[plotName] / areaConvert, 2));
                // calculate common area percent and common area
                foreach (string property in areaGroupsSeperateProperties.Keys)
                {
                    calculateSurplusPercentandArea(areaGroupsSeperateProperties[property], "A Instance Common Area %", "A Instance Common Area", propertyCommonAreas[plotName][property]);
                    calculateSpecialCommonAreaSurplus(areaGroupsSeperateProperties[property], "A Instance Common Area Special", propertyCommonAreasSpecial[plotName][property]);
                }              
            }            
        }
        public void exportToExcel(string filePath, string sheetName)
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = excelApplication.Workbooks.Open(filePath, ReadOnly: false);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[sheetName];

            // set columns' width
            workSheet.Range["A:A"].ColumnWidth = 30;
            workSheet.Range["B:B"].ColumnWidth = 30;
            workSheet.Range["C:C"].ColumnWidth = 15;
            workSheet.Range["D:D"].ColumnWidth = 15;
            workSheet.Range["E:E"].ColumnWidth = 10;
            workSheet.Range["F:N"].ColumnWidth = 5;
            workSheet.Range["O:O"].ColumnWidth = 10;
            workSheet.Range["P:T"].ColumnWidth = 15;
            workSheet.Range["U:W"].ColumnWidth = 15;

            int x = 1;

            // general formatting
            // main title : IPID and project number
            workSheet.Cells[x, 1] = "IPID";
            workSheet.Cells[x, 2] = doc.ProjectInformation.LookupParameter("Project Number").AsString();

            Range mergeRange = workSheet.Range[$"B{x}", $"W{x}"];
            mergeRange.Merge();
            mergeRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            mergeRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            mergeRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
            mergeRange.RowHeight = 35;
            // mergeRange.VerticalAlignment = XlVAlign.xlVAlignTop;

            Range ipIdRange = workSheet.Range[$"A{x}", $"A{x}"];
            ipIdRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            ipIdRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            ipIdRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

            // main title : project name
            x += 2;
            workSheet.Cells[x, 1] = "ОБЕКТ";
            workSheet.Cells[x, 2] = doc.ProjectInformation.LookupParameter("Project Address").AsString();

            Range mergeRangeObject = workSheet.Range[$"B{x}", $"W{x}"];
            mergeRangeObject.Merge();
            mergeRangeObject.Borders.LineStyle = XlLineStyle.xlContinuous;
            mergeRangeObject.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            mergeRangeObject.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);
            mergeRangeObject.RowHeight = 35;
            // mergeRangeObject.VerticalAlignment = XlVAlign.xlVAlignTop;

            Range mergeRangeProjName = workSheet.Range[$"A{x}", $"A{x}"];
            mergeRangeProjName.Borders.LineStyle = XlLineStyle.xlContinuous;
            mergeRangeProjName.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            mergeRangeProjName.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

            foreach (string plotName in plotNames)
            {
                x += 2;
                int rangeStart = x;

                // general plot data
                // plot row
                Range plotRange = workSheet.Range[$"A{x}", $"W{x}"];
                object[] plotStrings = new[] { "УПИ:", $"{Math.Round(plotAreasImp[plotName] / areaConvert, 2)}", "m2", "", 
                    "Самостоятелни обекти и паркоместа:", "", "", "", "", "", "", "Обекти на терен:", "", "", "", "", "", "Забележки:", "", "", "", "", "" };
                plotRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, plotStrings);

                // build up area row
                x++;
                Range baRange = workSheet.Range[$"A{x}", $"W{x}"];
                object[] baStrings = new[] { "ЗП:", $"{Math.Round(plotBuildAreas[plotName], 2)}", "m2", "", "Ателиета:", "", "", "", "0", "бр", "", "Паркоместа:",
                    "", "", "0", "бр", "", "За целите на ценообразуването и площообразуването, от площта на общите части са приспаднати ХХ.ХХкв.м. :", "", "", "", "", "" };
                baRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, baStrings);

                // total build area row
                x++;
                Range tbaRange = workSheet.Range[$"A{x}", $"W{x}"];
                string[] tbaStrings = new[] { "РЗП:", $"{Math.Round(plotTotalBuild[plotName], 2)}", "m2", "",
                    "Апартаменти:", "", "", "", "0", "бр", "", "Дворове:", "", "", "0", "бр", "", "", "", "", "", "", "" };
                tbaRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, tbaStrings);

                // underground row
                x++;
                Range uRange = workSheet.Range[$"A{x}", $"W{x}"];
                string[] uStrings = new[] { "Сутерени:", $"{Math.Round(plotUndergroundAreas[plotName], 2)}", "m2", "", 
                    "Магазини:", "", "", "", "0", "бр", "", "Трафопост:", "", "", "0", "бр", "", "", "", "", "", "", "" };
                uRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, uStrings);

                // underground + tba row
                x++;
                Range utbaRange = workSheet.Range[$"A{x}", $"W{x}"];
                string[] utbaStrings = new[] { "РЗП + Сутерени:", $"{Math.Round(plotUndergroundAreas[plotName], 2) + Math.Round(plotTotalBuild[plotName], 2)}", 
                    "m2", "", "Офиси", "", "", "", "0", "бг", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                utbaRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, utbaStrings);

                // CO row
                x++;
                Range coRange = workSheet.Range[$"A{x}", $"W{x}"];
                string[] coStrings = new[] { "Общо СО", $"{Math.Round(plotIndividualAreas[plotName], 2)}", "m2", "", 
                    "Гаражи", "", "", "", "0", "бр", "", "Данни за обекта:", "", "", "", "", "", "", "", "", "", "", "" };
                coRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, coStrings);

                // CA row
                x++;
                Range caRange = workSheet.Range[$"A{x}", $"W{x}"];
                string[] caStrings = new[] { "Общо ОЧ", $"{Math.Round(plotCommonAreas[plotName], 2)}", "m2", "", 
                    "Складове", "", "", "", "0", "бр", "", "Етажност", "", "", "ет", "", "", "", "", "", "", "", "" };
                caRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, caStrings);

                // land row
                x += 1;
                Range landRange = workSheet.Range[$"A{x}", $"W{x}"];
                string[] landStrings = new[] { "Земя към СО:", $"{Math.Round(plotLandAreas[plotName], 2)}", "m2", "", 
                    "Паркоместа", "", "", "", "0", "бр", "", "Система", "", "монолитна", "", "", "", "", "", "", "", "", "" };
                landRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, landStrings);

                // set borders
                int rangeEnd = x;

                Range cellsOne = workSheet.Range[$"A{rangeStart}", $"C{rangeEnd}"];
                Borders bordersOne = cellsOne.Borders;
                bordersOne[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersOne[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersOne[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersOne[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsOne.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                Range cellsTwo = workSheet.Range[$"D{rangeStart}", $"D{rangeEnd}"];
                Borders bordersTwo = cellsTwo.Borders;
                bordersTwo[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersTwo[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersTwo[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersTwo[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsTwo.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                Range cellsThree = workSheet.Range[$"E{rangeStart}", $"J{rangeEnd}"];
                Borders bordersThree = cellsThree.Borders;
                bordersThree[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersThree[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersThree[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersThree[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsThree.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                Range cellsFour = workSheet.Range[$"K{rangeStart}", $"K{rangeEnd}"];
                Borders bordersFour = cellsFour.Borders;
                bordersFour[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersFour[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersFour[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersFour[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsFour.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                Range cellsFive = workSheet.Range[$"L{rangeStart}", $"P{rangeEnd}"];
                Borders bordersFive = cellsFive.Borders;
                bordersFive[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersFive[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersFive[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersFive[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsFive.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                Range cellsSix = workSheet.Range[$"Q{rangeStart}", $"Q{rangeEnd}"];
                Borders bordersSix = cellsSix.Borders;
                bordersSix[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersSix[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersSix[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersSix[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsSix.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                Range cellsSeven = workSheet.Range[$"R{rangeStart}", $"V{rangeEnd}"];
                Borders bordersSeven = cellsSeven.Borders;
                bordersSeven[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersSeven[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersSeven[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersSeven[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsSeven.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                Range cellsEight = workSheet.Range[$"R{rangeStart}", $"W{rangeEnd}"];
                Borders bordersEight = cellsEight.Borders;
                bordersEight[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                bordersEight[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                bordersEight[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                bordersEight[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;
                cellsEight.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

                // a list, storing data about end lines of each separate property group with respect to column V
                List<string> propertyEndLinesLandSum = new List<string>();
                // a list, storing data about end lines of each separate proeprty group with respect to column W
                List<string> propertyEndLineslandSumArea = new List<string>();

                foreach (string property in plotProperties[plotName])
                {
                    if (!property.Contains("+") && !property.ToLower().Contains("траф"))
                    {
                        x += 2;
                        workSheet.Cells[x, 1] = $"ПЛОЩООБРАЗУВАНЕ САМОСТОЯТЕЛНИ ОБЕКТИ - {property}";
                        Range propertyTitleRange = workSheet.Range[$"A{x}", $"W{x}"];
                        propertyTitleRange.Merge();
                        propertyTitleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                        propertyTitleRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        propertyTitleRange.Font.Bold = true;
                        propertyTitleRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        propertyTitleRange.RowHeight = 35;
                        propertyTitleRange.VerticalAlignment = XlVAlign.xlVAlignTop;

                        x++;
                        Range propertyDataRange = workSheet.Range[$"A{x}", $"W{x}"];
                        string[] propertyData = new[] { $"ПЛОЩ СО: {Math.Round(propertyIndividualAreas[plotName][property], 2)} кв.м", 
                            $"ОБЩО ОЧ: {Math.Round(propertyCommonAreasAll[plotName][property], 2)} кв.м", 
                            $"ОЧ: {Math.Round(propertyCommonAreas[plotName][property], 2)} кв.м", 
                            $"СП.ОЧ: {Math.Round(propertyCommonAreasSpecial[plotName][property], 2)} кв.м",
                            "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                        propertyDataRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, propertyData);
                        propertyDataRange.Font.Bold = true;
                        propertyDataRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        propertyDataRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        propertyDataRange.RowHeight = 35;
                        propertyDataRange.VerticalAlignment = XlVAlign.xlVAlignTop;

                        x++;
                        Range parameterNamesRange = workSheet.Range[$"A{x}", $"W{x}"];
                        string[] parameterNamesData = new[] { "ПЛОЩ СО", "НАИМЕНОВАНИЕ СО", "ПЛОЩ F1(F2)", "ПРИЛЕЖАЩА \nПЛОЩ", "Кпг", "Ки", "Кв", "Км", "Кив", "Кпп", 
                            "Кок", "Ксп", "Кк", "К", "C1(C2)", "И.Ч. \nПАРКОМЕСТА", "ОБЩИ ЧАСТИ - F3", "", "СП.ОБЩИ \nЧАСТИ - F4", "ОБЩО-\nF1(F2)+F3+F4", "ПРАВО \nНА СТРОЕЖ", "ЗЕМЯ", ""};
                        parameterNamesRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, parameterNamesData);
                        parameterNamesRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        parameterNamesRange.Font.Bold = true;
                        parameterNamesRange.Borders.LineStyle = XlLineStyle.xlContinuous;
                        parameterNamesRange.VerticalAlignment = XlVAlign.xlVAlignTop;
                        Range commonAreasMerge = workSheet.Range[$"Q{x}", $"R{x}"];
                        commonAreasMerge.Merge();
                        Range landMerge = workSheet.Range[$"V{x}", $"W{x}"];
                        landMerge.Merge();

                        x++;
                        Range parametersTypeRange = workSheet.Range[$"A{x}", $"W{x}"];
                        string[] parametersTypeData = new[] { "", "", "m2", "m2", "", "", "", "", "", "", "", "", "", "", "", "% и.ч.", "% и.ч.", "m2", "m2", "m2", "% и.ч.", "% и.ч.", "m2" };
                        parametersTypeRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, parametersTypeData);
                        parametersTypeRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        parametersTypeRange.Font.Bold = true;
                        parametersTypeRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                        x++;
                        Range blankLineRange = workSheet.Range[$"A{x}", $"W{x}"];
                        blankLineRange.Merge();
                        blankLineRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                        Borders blankBorders = blankLineRange.Borders;
                        blankBorders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                        blankBorders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                        blankBorders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                        blankBorders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                        x++;

                        int startLine = x;

                        List<Area> sortedAreasNormal = AreasOrganizer[plotName][property]
                                .Where(area => area.LookupParameter("A Instance Area Category").AsString().ToLower().Equals("самостоятелен обект"))
                                //.Where(area => !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                                .Where(area => !area.LookupParameter("A Instance Area Group").AsString().Equals("ЗЕМЯ"))
                                .OrderBy(area => ReorderEntrance(area.LookupParameter("A Instance Area Entrance").AsString()))
                                .ThenBy(area => ExtractLevelNumber(area.LookupParameter("Level").AsValueString()))
                                .ThenBy(area => area.LookupParameter("Number").AsString())
                                .ToList();

                        List<Area> areasToSort = new List<Area>();
                        if (AreasOrganizer[plotName].ContainsKey("ЗЕМЯ"))
                            areasToSort.AddRange(AreasOrganizer[plotName]["ЗЕМЯ"]);
                        if (AreasOrganizer[plotName].ContainsKey("ТРАФ"))
                            areasToSort.AddRange(AreasOrganizer[plotName]["ТРАФ"]);

                        List<Area> sortedAreasGround = areasToSort
                                .Where(area => new List<string> { "земя", "траф" }.Contains(area.LookupParameter("A Instance Area Group").AsString().ToLower()))
                                .OrderBy(area => ReorderEntrance(area.LookupParameter("A Instance Area Entrance").AsString()))
                                .ThenBy(area => ExtractLevelNumber(area.LookupParameter("Level").AsValueString()))
                                .ThenBy(area => area.LookupParameter("Number").AsString())
                                .ToList();

                        List<Area> sortedAreas = new List<Area>();

                        if (property.ToLower() != "земя")
                        {
                            sortedAreas = sortedAreasNormal;
                        }
                        else
                        {
                            sortedAreas = sortedAreasGround;
                        }

                        List<string> levels = new List<string>();
                        List<string> entrances = new List<string>();

                        foreach (Area area in sortedAreas)
                        {
                            if (!area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != "")
                            {
                                if (!entrances.Contains(area.LookupParameter("A Instance Area Entrance").AsString()))
                                {
                                    entrances.Add(area.LookupParameter("A Instance Area Entrance").AsString());
                                    workSheet.Cells[x, 1] = area.LookupParameter("A Instance Area Entrance").AsString();
                                    Range entranceRangeString = workSheet.Range[$"A{x}", $"W{x}"];
                                    entranceRangeString.Merge();
                                    entranceRangeString.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                    x++;
                                    levels.Clear();
                                }

                                if (!levels.Contains(area.LookupParameter("Level").AsValueString()))
                                {
                                    levels.Add(area.LookupParameter("Level").AsValueString());
                                    double levelHeight = Math.Round(doc.GetElement(area.LookupParameter("Level").AsElementId()).LookupParameter("Elevation").AsDouble() * lengthConvert / 100, 3);
                                    string levelHeightStr;

                                    if (levelHeight < 0)
                                    {
                                        string tempString = levelHeight.ToString("F3");
                                        levelHeightStr = $"{tempString}";
                                    }
                                    else if (levelHeight > 0)
                                    {
                                        string tempString = levelHeight.ToString("F3");
                                        levelHeightStr = $"+ {tempString}";
                                    }
                                    else
                                    {
                                        string tempString = levelHeight.ToString("F3");
                                        levelHeightStr = $"± {tempString}";
                                    }

                                    workSheet.Cells[x, 1] = $"{area.LookupParameter("Level").AsValueString()} {levelHeightStr}";
                                    Range levelsRangeString = workSheet.Range[$"A{x}", $"W{x}"];
                                    levelsRangeString.Merge();
                                    levelsRangeString.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                                    x++;
                                }

                                try
                                {
                                    Range cellRangeString = workSheet.Range[$"A{x}", $"B{x}"];
                                    Range cellRangeDouble = workSheet.Range[$"C{x}", $"W{x}"];

                                    string areaNumber = area.LookupParameter("Number").AsString() ?? "SOMETHING'S WRONG";
                                    string areaName = area.LookupParameter("Name")?.AsString() ?? "SOMETHING'S WRONG";
                                    double areaArea = Math.Round(area.LookupParameter("A Instance Gross Area")?.AsDouble() / areaConvert ?? 0.0, 2);
                                    // TODO: rework properly for adjascent area
                                    object areaSubjected = DBNull.Value;
                                    // TODO: rework properly for adjascent area
                                    double ACGA = area.LookupParameter("A Coefficient Garage (Кпг)")?.AsDouble() ?? 0.0;
                                    double ACOR = area.LookupParameter("A Coefficient Orientation (Ки)")?.AsDouble() ?? 0.0;
                                    double ACLE = area.LookupParameter("A Coefficient Level (Кв)")?.AsDouble() ?? 0.0;
                                    double ACLO = area.LookupParameter("A Coefficient Location (Км)")?.AsDouble() ?? 0.0;
                                    double ACHE = area.LookupParameter("A Coefficient Height (Кив)")?.AsDouble() ?? 0.0;
                                    double ACRO = area.LookupParameter("A Coefficient Roof (Кпп)")?.AsDouble() ?? 0.0;
                                    double ACSP = area.LookupParameter("A Coefficient Special (Кок)")?.AsDouble() ?? 0.0;
                                    double ACST = area.LookupParameter("A Coefficient Storage (Ксп)")?.AsDouble() ?? 0.0;
                                    double ACZO = area.LookupParameter("A Coefficient Zones (Кк)")?.AsDouble() ?? 0.0;
                                    double ACCO = area.LookupParameter("A Coefficient Multiplied")?.AsDouble() ?? 0.0;
                                    double C1C2 = Math.Round(area.LookupParameter("A Instance Price C1/C2")?.AsDouble() ?? 0.0, 2);
                                    double areaCommonPercent = Math.Round(area.LookupParameter("A Instance Common Area %")?.AsDouble() ?? 0.0, 3);
                                    double areaCommonArea = Math.Round(area.LookupParameter("A Instance Common Area")?.AsDouble() / areaConvert ?? 0.0, 2);
                                    double areaCommonAreaSpecialPercent = Math.Round(area.LookupParameter("A Instance Common Area Special")?.AsDouble() / areaConvert ?? 0.0, 2);                                    double areaTotalArea = Math.Round((area.LookupParameter("A Instance Total Area")?.AsDouble() / areaConvert ?? 0.0), 2);
                                    double areaPermitPercent = Math.Round(area.LookupParameter("A Instance Building Permit %")?.AsDouble() ?? 0.0, 3);
                                    double areaRLPPercentage = Math.Round(area.LookupParameter("A Instance RLP Area %")?.AsDouble() ?? 0.0, 3);
                                    double areaRLP = Math.Round(area.LookupParameter("A Instance RLP Area")?.AsDouble() / areaConvert ?? 0.0, 2);

                                    string[] areaStringData = new[] { areaNumber, areaName };
                                    object[] areasDoubleData = new object[] { areaArea, areaSubjected, ACGA, ACOR, ACLE, ACLO, ACHE, ACRO, ACSP, ACST, ACZO, ACCO, C1C2, DBNull.Value, 
                                        areaCommonPercent, areaCommonArea, areaCommonAreaSpecialPercent, areaTotalArea, areaPermitPercent, areaRLPPercentage, areaRLP};

                                    cellRangeString.set_Value(XlRangeValueDataType.xlRangeValueDefault, areaStringData);
                                    cellRangeString.Borders.LineStyle = XlLineStyle.xlContinuous;

                                    cellRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, areasDoubleData);
                                    cellRangeDouble.Borders.LineStyle = XlLineStyle.xlContinuous;
                                }
                                catch
                                {
                                    Range cellRangeString = workSheet.Range[$"A{x}", $"B{x}"];
                                    string[] cellsStrings = new[] { "X", "Y" };
                                    cellRangeString.set_Value(XlRangeValueDataType.xlRangeValueDefault, cellsStrings);
                                }

                                x++;

                                if (!doesHaveRoomsAdjascent(area.LookupParameter("Number").AsString()))
                                {
                                    // adjascent areas loop
                                    foreach (Area areaSub in sortedAreas)
                                    {
                                        string primaryArea = areaSub.LookupParameter("A Instance Area Primary").AsString();

                                        if (primaryArea != null && primaryArea.Equals(area.LookupParameter("Number").AsString()))
                                        {
                                            Range areaAdjRangeStr = workSheet.Range[$"A{x}", $"B{x}"];
                                            areaAdjRangeStr.set_Value(XlRangeValueDataType.xlRangeValueDefault, new[] { areaSub.LookupParameter("Number").AsString(),
                                                                                                                        areaSub.LookupParameter("Name").AsString() });

                                            areaAdjRangeStr.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                            areaAdjRangeStr.Borders.LineStyle = XlLineStyle.xlContinuous;

                                            Range areaAdjRangeDouble = workSheet.Range[$"C{x}", $"W{x}"];
                                            areaAdjRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, new object[] {DBNull.Value,
                                                        Math.Round(areaSub.LookupParameter("Area").AsDouble() / areaConvert, 2), DBNull.Value, DBNull.Value,
                                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                        Math.Round(areaSub.LookupParameter("A Instance RLP Area %")?.AsDouble() ?? 0.0, 3),
                                                        Math.Round(areaSub.LookupParameter("A Instance RLP Area")?.AsDouble() / areaConvert ?? 0.0, 2) });

                                            Borders areaAdjRangeBorders = areaAdjRangeDouble.Borders;
                                            areaAdjRangeBorders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                                            areaAdjRangeBorders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                                            areaAdjRangeBorders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                                            areaAdjRangeBorders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                                            x++;
                                        }
                                    }

                                    // also search for adjascent areas from within ground property group
                                    if (AreasOrganizer[plotName].ContainsKey("ЗЕМЯ"))
                                    {
                                        foreach (Area areaGround in AreasOrganizer[plotName]["ЗЕМЯ"])
                                        {
                                            string primaryArea = areaGround.LookupParameter("A Instance Area Primary").AsString();

                                            if (primaryArea != null && primaryArea.Equals(area.LookupParameter("Number").AsString()))
                                            {
                                                Range areaAdjRangeStr = workSheet.Range[$"A{x}", $"B{x}"];
                                                areaAdjRangeStr.set_Value(XlRangeValueDataType.xlRangeValueDefault, new[] { areaGround.LookupParameter("Number").AsString(), areaGround.LookupParameter("Name").AsString() });

                                                areaAdjRangeStr.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                                areaAdjRangeStr.Borders.LineStyle = XlLineStyle.xlContinuous;

                                                Range areaAdjRangeDouble = workSheet.Range[$"C{x}", $"W{x}"];
                                                areaAdjRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, new object[] {DBNull.Value,
                                                        Math.Round(areaGround.LookupParameter("Area").AsDouble() / areaConvert, 2), DBNull.Value, DBNull.Value,
                                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                        Math.Round(areaGround.LookupParameter("A Instance RLP Area %")?.AsDouble() ?? 0.0, 3),
                                                        Math.Round(areaGround.LookupParameter("A Instance RLP Area")?.AsDouble() / areaConvert ?? 0.0, 2) });

                                                Borders areaAdjRangeBorders = areaAdjRangeDouble.Borders;
                                                areaAdjRangeBorders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                                                areaAdjRangeBorders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                                                areaAdjRangeBorders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                                                areaAdjRangeBorders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                                                x++;
                                            }
                                        }
                                    }

                                    // adjascent areas loop for special common areas
                                    foreach (Area areaSub in AreasOrganizer[plotName][property])
                                    {
                                        if (areaSub.LookupParameter("A Instance Area Category").AsString().ToLower() == "обща част")
                                        {
                                            string primaryArea = areaSub.LookupParameter("A Instance Area Primary").AsString();

                                            if (primaryArea != null)
                                            {
                                                string[] result = primaryArea.Split(new char[] { '+' }, StringSplitOptions.RemoveEmptyEntries)
                                                    .Select(s => s.Trim())
                                                    .ToArray();

                                                if (result.Contains(area.LookupParameter("Number").AsString()))
                                                {
                                                    Range areaAdjRangeStr = workSheet.Range[$"A{x}", $"B{x}"];
                                                    areaAdjRangeStr.set_Value(XlRangeValueDataType.xlRangeValueDefault, new[] { areaSub.LookupParameter("Number").AsString(),
                                                                                                                                areaSub.LookupParameter("Name").AsString() });

                                                    areaAdjRangeStr.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                                    areaAdjRangeStr.Borders.LineStyle = XlLineStyle.xlContinuous;

                                                    Range areaAdjRangeDouble = workSheet.Range[$"C{x}", $"W{x}"];
                                                    areaAdjRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, new object[] {DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                                DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                                DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value });

                                                    Borders areaAdjRangeBorders = areaAdjRangeDouble.Borders;
                                                    areaAdjRangeBorders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                                                    areaAdjRangeBorders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                                                    areaAdjRangeBorders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                                                    areaAdjRangeBorders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                                                    x++;
                                                }
                                            }
                                        }
                                    }
                                }
                                else
                                {
                                    double totalArea = Math.Round(area.LookupParameter("A Instance Gross Area").AsDouble(), 3);

                                    Dictionary<List<object>, Room> adjascentRooms = returnAdjascentRoomsTest(area.LookupParameter("Number").AsString(), totalArea);

                                    foreach (List<object> key in adjascentRooms.Keys)
                                    {
                                        Room room = adjascentRooms[key];

                                        Range areaAdjRangeStr = workSheet.Range[$"A{x}", $"B{x}"];
                                        areaAdjRangeStr.set_Value(XlRangeValueDataType.xlRangeValueDefault,
                                            new[] { key[0], room.LookupParameter("Name").AsString() });

                                        areaAdjRangeStr.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                        areaAdjRangeStr.Borders.LineStyle = XlLineStyle.xlContinuous;

                                        Range areaAdjRangeDouble = workSheet.Range[$"C{x}", $"W{x}"];
                                        areaAdjRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, new object[] {DBNull.Value,
                                                        Math.Round(room.LookupParameter("Area").AsDouble() / areaConvert, 2), DBNull.Value, DBNull.Value,
                                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                        DBNull.Value, DBNull.Value, key[1], DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                                        DBNull.Value, DBNull.Value, DBNull.Value});

                                        Borders areaAdjRangeBorders = areaAdjRangeDouble.Borders;
                                        areaAdjRangeBorders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                                        areaAdjRangeBorders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                                        areaAdjRangeBorders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                                        areaAdjRangeBorders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                                        x++;
                                    }
                                }
                            }
                        }

                        int endLine = x - 1;
                        propertyEndLinesLandSum.Add($"V{x}");
                        propertyEndLineslandSumArea.Add($"W{x}");

                        Range colorRange = workSheet.Range[$"C{startLine}", $"W{endLine}"];
                        colorRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.AliceBlue);

                        // set a formula for the total area sum of F1/F2
                        Range sumF1F2 = workSheet.Range[$"C{x}", $"C{x}"];
                        sumF1F2.Formula = $"=SUM(C{startLine + 2}:C{endLine})";

                        // set a formula for the total sum of adjascent areas
                        Range sumAdjascent = workSheet.Range[$"D{x}", $"D{x}"];
                        sumAdjascent.Formula = $"=SUM(D{startLine + 2}:D{endLine})";

                        // set a formula for the total sum of C1/C2
                        Range sumC1C2 = workSheet.Range[$"O{x}", $"O{x}"];
                        sumC1C2.Formula = $"=SUM(O{startLine + 2}:O{endLine})";

                        // set a formula for the total sum of Common Areas 
                        Range sumCommonAreasPercentage = workSheet.Range[$"Q{x}", $"Q{x}"];
                        sumCommonAreasPercentage.Formula = $"=SUM(Q{startLine + 2}:Q{endLine})";

                        // set a formula for the total sum of Special Common Areas Percentage
                        Range sumCommonArea = workSheet.Range[$"R{x}", $"R{x}"];
                        sumCommonArea.Formula = $"=SUM(R{startLine + 2}:R{endLine})";

                        // set a formula for the total sum of Special Common Areas Percentage
                        Range sumSpecialCommonArea = workSheet.Range[$"S{x}", $"S{x}"];
                        sumSpecialCommonArea.Formula = $"=SUM(S{startLine + 2}:S{endLine})";

                        // set a formula for the total sum of Total Area
                        Range totalAreaF1F2F3F4 = workSheet.Range[$"T{x}", $"T{x}"];
                        totalAreaF1F2F3F4.Formula = $"=SUM(T{startLine + 2}:T{endLine})";

                        // set a formula for the total sum of Building Right Percentage
                        Range sumBuildingRightPercentage = workSheet.Range[$"U{x}", $"U{x}"];
                        sumBuildingRightPercentage.Formula = $"=SUM(U{startLine + 2}:U{endLine})";

                        // set a formula for the total sum of Land Area Percentage
                        Range landAreapercentage = workSheet.Range[$"V{x}", $"V{x}"];
                        landAreapercentage.Formula = $"=SUM(V{startLine + 2}:V{endLine})";

                        // set a formula for the total sum of Land Area
                        Range landArea = workSheet.Range[$"W{x}", $"W{x}"];
                        landArea.Formula = $"=SUM(W{startLine + 2}:W{endLine})";

                        // set coloring for the summed up rows
                        Range colorRangePropertySum = workSheet.Range[$"A{endLine + 1}", $"W{endLine + 1}"];
                        colorRangePropertySum.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                        x++;
                    }
                }

                string propertyLandSumFormula = $"=SUM({string.Join(",", propertyEndLinesLandSum)})";
                workSheet.Range[$"V{x}"].Formula = propertyLandSumFormula;
                workSheet.Range[$"V{x}"].Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.AntiqueWhite);

                string propertyLandSumAreaFormula = $"=SUM({string.Join(",", propertyEndLineslandSumArea)})";
                workSheet.Range[$"W{x}"].Formula = propertyLandSumAreaFormula;
                workSheet.Range[$"W{x}"].Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.AntiqueWhite);
            }

            workBook.Save();
            workBook.Close();
        }
    }
}
