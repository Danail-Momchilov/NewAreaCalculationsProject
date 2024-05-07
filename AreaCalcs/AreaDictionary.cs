﻿using Autodesk.Revit.DB;
using Autodesk.Revit.UI;
using System;
using System.Collections.Generic;
using System.Linq;
using System.Text;
using System.Threading.Tasks;
using Microsoft.Office.Interop.Excel;
using Document = Autodesk.Revit.DB.Document;
using System.Runtime.InteropServices;

namespace AreaCalculations
{
    internal class AreaDictionary
    {
        public Dictionary<string, Dictionary<string, List<Area>>> AreasOrganizer { get; set; }
        public List<string> plotNames { get; set; }
        public Dictionary<string, double> plotAreasImp { get; set; }
        public Dictionary<string, List<string>> plotProperties { get; set; }
        public Document doc { get; set; }
        public Transaction transaction { get; set; }
        public double areasCount { get; set; }
        public AreaDictionary(Document activeDoc)
        {
            this.doc = activeDoc;
            this.AreasOrganizer = new Dictionary<string, Dictionary<string, List<Area>>>();
            this.plotNames = new List<string>();
            this.plotProperties = new Dictionary<string, List<string>>();
            this.transaction = new Transaction(activeDoc, "Calculate and Update Area Parameters");
            this.plotAreasImp = new Dictionary<string, double>();
            this.areasCount = 0;

            ProjectInfo projectInfo = activeDoc.ProjectInformation;

            FilteredElementCollector areasCollector = new FilteredElementCollector(activeDoc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();

            foreach (Area area in areasCollector)
            {
                string plotName = area.LookupParameter("A Instance Area Plot").AsString();
                string groupName = area.LookupParameter("A Instance Area Group").AsString();

                if (!string.IsNullOrEmpty(plotName) && !string.IsNullOrEmpty(groupName))
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
                    // TODO: to be continued
                    // TODO: check whether an Area has a plotName, that is not relative to the ones in the Project Info (is it actually needed?)
                }
            }
        }

        private double areaConvert = 10.763914692;
        public string calculatePrimaryArea()
        {
            string errorMessage = "";
            List<string> missingNumbers = new List<string>();

            transaction.Start();
            
            foreach (Area mainArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
            {
                mainArea.LookupParameter("A Instance Gross Area").Set(mainArea.LookupParameter("Area").AsDouble());

                foreach (Area secArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
                {
                    if (secArea.LookupParameter("A Instance Area Primary").AsString() == mainArea.LookupParameter("Number").AsString() && secArea.LookupParameter("A Instance Area Primary").HasValue && secArea.Area != 0)
                    {
                        double sum = mainArea.LookupParameter("A Instance Gross Area").AsDouble() + secArea.LookupParameter("Area").AsDouble();
                        mainArea.LookupParameter("A Instance Gross Area").Set(sum);
                    }
                }
            }

            foreach (Area secArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
            {
                if (secArea.LookupParameter("A Instance Area Primary").HasValue && secArea.LookupParameter("A Instance Area Primary").AsString() != "" && secArea.Area != 0)
                {
                    bool wasFound = false;

                    foreach (Area mainArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
                        if (secArea.LookupParameter("A Instance Area Primary").AsString() == mainArea.LookupParameter("Number").AsString())
                            wasFound = true;

                    if (!wasFound && !missingNumbers.Contains(secArea.LookupParameter("Number").AsString()))
                    {
                        missingNumbers.Add(secArea.LookupParameter("Number").AsString());
                        errorMessage += $"Грешка: Area {secArea.LookupParameter("Number").AsString()} / id: {secArea.Id} / Посочената Area е зададена като подчинена на такава с несъществуващ номер. Моля, проверете го и стартирайте апликацията отново\n";
                    }
                }
            }
            /*
            
            foreach (Area mainArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
            {
                mainArea.LookupParameter("A Instance Gross Area").Set(mainArea.LookupParameter("Area").AsDouble());
            }

            foreach (Area secondaryArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
            {
                bool wasfound = false;
                
                if (secondaryArea.LookupParameter("A Instance Area Primary").HasValue && secondaryArea.LookupParameter("A Instance Area Primary").AsString() != "" && secondaryArea.Area != 0)
                {
                    foreach (Area mainArea in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
                    {
                        if (secondaryArea.LookupParameter("A Instance Area Primary").AsString() == mainArea.LookupParameter("Number").AsString())
                        {
                            double sum = mainArea.LookupParameter("A Instance Gross Area").AsDouble() + secondaryArea.LookupParameter("Area").AsDouble();

                            bool wasSet = mainArea.LookupParameter("A Instance Gross Area").Set(sum);

                            wasfound = true;

                            // test

                            errorMessage += $"An area of: {secondaryArea.LookupParameter("Area").AsDouble()/areaConvert} is added to the one of Area Number: {mainArea.Number}, " +
                                $"whioch is at the moment: {mainArea.Area/areaConvert}! The total sum is {secondaryArea.LookupParameter("Area").AsDouble() / areaConvert + mainArea.Area / areaConvert}\n" +
                                $"calculated sum = {sum/areaConvert}\ncalculated sum in imperial = {sum}\nwasSet = {wasSet}\n" +
                                $"value for A Instance Gross Area after being set: = {mainArea.LookupParameter("A Instance Gross Area").AsDouble()}\n" +
                                $"the secondary Area object, added to the main one, has a number {secondaryArea.Number} and an Id: {secondaryArea.Id}\n\n";

                            // test
                        }
                    }

                    if (!wasfound && !missingNumbers.Contains(secondaryArea.LookupParameter("Number").AsString()))
                    {
                        missingNumbers.Add(secondaryArea.LookupParameter("Number").AsString());
                        errorMessage += $"Грешка: Area {secondaryArea.LookupParameter("Number").AsString()} / id: {secondaryArea.Id} / Посочената Area е зададена като подчинена на такава с несъществуващ номер. Моля, проверете го и стартирайте апликацията отново\n";
                    }
                }
            }
            */
            transaction.Commit();

            return errorMessage;
        }
        public void calculateC1C2()
        {
            transaction.Start();

            foreach (Area area in new FilteredElementCollector(doc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType().ToList())
            {
                if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                {
                    area.LookupParameter("A Instance Price C1/C2").Set((area.LookupParameter("A Instance Gross Area").AsDouble() * area.LookupParameter("A Coefficient Multiplied").AsDouble()) / areaConvert);
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
                    double totaC1C2 = 0;

                    // calculate total summed up C1C2 for the given property
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double C1C2 = area.LookupParameter("A Instance Price C1/C2").AsDouble();
                            totaC1C2 += C1C2;
                        }
                    }

                    // calculate common area percentage parameter for each area
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double commonAreaPercent = (area.LookupParameter("A Instance Price C1/C2").AsDouble() / totaC1C2) * 100;
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
                    double commonAreas = 0;

                    // calculate total summed up area of all common areas                    
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "ОБЩА ЧАСТ")
                        {
                            commonAreas += area.LookupParameter("Area").AsDouble();
                        }
                    }

                    // calculate common area percentage parameter for each area
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double commonArea = 0;

                            commonArea = (area.LookupParameter("A Instance Common Area %").AsDouble() * commonAreas) / 100;
                            area.LookupParameter("A Instance Common Area").Set(commonArea);
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
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            area.LookupParameter("A Instance Total Area").Set(area.LookupParameter("A Instance Gross Area").AsDouble() + area.LookupParameter("A Instance Common Area").AsDouble());
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
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            totalPlotC1C2 += area.LookupParameter("A Instance Price C1/C2").AsDouble();
                        }
                    }
                }

                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
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
                double totalPlotC1C2 = 0;

                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            totalPlotC1C2 += area.LookupParameter("A Instance Price C1/C2").AsDouble();
                        }
                    }
                }

                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double rlpAreaPercentage = (area.LookupParameter("A Instance Total Area").AsDouble() / totalPlotC1C2) * 100;
                            area.LookupParameter("A Instance RLP Area %").Set(rlpAreaPercentage / areaConvert); 
                            // TODO: check this calculation. Why does it work properly, while the calculateBuildingPercentPermit method works without the need to apply the areaConvert?
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
                double plotAreaImp = plotAreasImp[plotName];

                foreach (string property in plotProperties[plotName])
                {
                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != ""))
                        {
                            double rlpAreaImp = (plotAreaImp * area.LookupParameter("A Instance RLP Area %").AsDouble()) / 100;
                            area.LookupParameter("A Instance RLP Area").Set(rlpAreaImp);
                        }
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
                        if (area.LookupParameter("A Instance Area Category").AsString() == "САМОСТОЯТЕЛЕН ОБЕКТ" && !(area.LookupParameter("A Instance Area Primary").HasValue && area.LookupParameter("A Instance Area Primary").AsString() != "") && area.Area != 0)
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
                                errorReport += $"{area.Id} {area.Name} A Instance Common Area = {area.LookupParameter("A Instance Common Area").AsDouble()} / A Instance Total Area = {area.LookupParameter("A Instance Total Area").AsDouble()}";
                            }
                        }
                    }
                }
            }

            transaction.Commit();

            return errorReport;
        }
        public void exportToExcel(string filePath, string sheetName)
        {
            Microsoft.Office.Interop.Excel.Application excelApplication = new Microsoft.Office.Interop.Excel.Application();
            Workbook workBook = excelApplication.Workbooks.Open(filePath, ReadOnly : false);
            Worksheet workSheet = (Worksheet)workBook.Worksheets[sheetName];

            int x = 1;

            foreach (string plotName in plotNames)
            {
                x += 2;

                foreach (string property in plotProperties[plotName])
                {
                    x += 1;

                    foreach (Area area in AreasOrganizer[plotName][property])
                    {
                        try
                        {
                            Range cellRangeString = workSheet.Range[$"A{x}", $"B{x}"];
                            Range cellRangeDouble = workSheet.Range[$"C{x}", $"X{x}"];

                            // todo: check them all once again in compliance with the chart structure
                            string areaNumber = area.LookupParameter("Number")?.AsString() ?? "WRONG";
                            string areaName = area.LookupParameter("Name")?.AsString() ?? "WRONG";
                            double areaArea = area.LookupParameter("A Instance Total Area")?.AsDouble() * areaConvert ?? 0.0;
                            // todo: rework properly for subjectivated area
                            double areaSubjected = area.LookupParameter("Area")?.AsDouble() * areaConvert ?? 0.0;
                            // todo: rework properly for subjectivated area
                            double ACGA = area.LookupParameter("A Coefficient Garage (Кпг)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACOR = area.LookupParameter("A Coefficient Orientation (Ки)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACLE = area.LookupParameter("A Coefficient Level (Кв)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACLO = area.LookupParameter("A Coefficient Location (Км)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACHE = area.LookupParameter("A Coefficient Height (Кив)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACRO = area.LookupParameter("A Coefficient Roof (Кпп)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACSP = area.LookupParameter("A Coefficient Special (Кок)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACST = area.LookupParameter("A Coefficient Storage (Ксп)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACZO = area.LookupParameter("A Coefficient Zones (Кк)")?.AsDouble() * areaConvert ?? 0.0;
                            double ACCO = area.LookupParameter("A Coefficient Multiplied")?.AsDouble() * areaConvert ?? 0.0;
                            double C1C2 = area.LookupParameter("A Instance Price C1/C2")?.AsDouble() * areaConvert ?? 0.0;
                            double areaCommonPercent = area.LookupParameter("A Instance Common Area %")?.AsDouble() * areaConvert ?? 0.0;
                            double areaCommonArea = area.LookupParameter("A Instance Common Area")?.AsDouble() * areaConvert ?? 0.0;
                            double areaTotalArea = (area.LookupParameter("A Instance Total Area")?.AsDouble() * areaConvert ?? 0.0) + (area.LookupParameter("A Instance Common Area")?.AsDouble() * areaConvert ?? 0.0);
                            double areaPermitPercent = area.LookupParameter("A Instance Building Permit %")?.AsDouble() * areaConvert ?? 0.0;
                            double areaRLPPercentage = area.LookupParameter("A Instance RLP Area &")?.AsDouble() * areaConvert ?? 0.0;
                            double areaRLP = area.LookupParameter("A Instance RLP Area")?.AsDouble() * areaConvert ?? 0.0;
                            // todo: check them all once again in compliance with the chart structure

                            string[] areaStringData = new[] { areaNumber, areaName };
                            double[] areasDoubleData = new[] { areaArea, areaSubjected, ACGA, ACOR, ACLE, ACLO, ACHE, ACRO, ACSP, ACST, ACZO, ACCO, C1C2, areaCommonPercent, areaCommonArea, areaTotalArea, areaPermitPercent, areaRLPPercentage, areaRLP };

                            cellRangeString.set_Value(XlRangeValueDataType.xlRangeValueDefault, areaStringData);
                            cellRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, areasDoubleData);                            
                        }
                        catch
                        {
                            Range cellRangeString = workSheet.Range[$"A{x}", $"B{x}"];
                            string[] cellsStrings = new[] { "X", "Y" };
                            cellRangeString.set_Value(XlRangeValueDataType.xlRangeValueDefault, cellsStrings);
                        }

                        x += 1;
                    }
                }
            }

            workBook.SaveAs(filePath, XlFileFormat.xlWorkbookDefault, Type.Missing, false, false, XlSaveAsAccessMode.xlNoChange);
            workBook.Close();
            /*
            excelApplication.Quit();

            Marshal.ReleaseComObject(workSheet);
            Marshal.ReleaseComObject(workBook);
            Marshal.ReleaseComObject(excelApplication);*/
        }
    }
}
