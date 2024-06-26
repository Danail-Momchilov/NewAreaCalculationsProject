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
using System.Drawing;
using static System.Net.Mime.MediaTypeNames;

namespace AreaCalculations
{
    internal class AreaDictionary
    {
        public Dictionary<string, Dictionary<string, List<Area>>> AreasOrganizer { get; set; }
        public List<string> plotNames { get; set; }
        public Dictionary<string, double> plotAreasImp { get; set; } // why is this thing storing data in imperial ?!?!
        public Dictionary<string, List<string>> plotProperties { get; set; }
        public Dictionary<string, double> plotBuildAreas { get; set; }
        public Dictionary<string, double> plotTotalBuild { get; set; }
        public Document doc { get; set; }
        public Transaction transaction { get; set; }
        public double areasCount { get; set; }
        public double missingAreasCount { get; set; }
        public string missingAreasData { get; set; }
        public AreaDictionary(Document activeDoc)
        {
            this.doc = activeDoc;
            this.AreasOrganizer = new Dictionary<string, Dictionary<string, List<Area>>>();
            this.plotNames = new List<string>();
            this.plotProperties = new Dictionary<string, List<string>>();
            this.transaction = new Transaction(activeDoc, "Calculate and Update Area Parameters");
            this.plotAreasImp = new Dictionary<string, double>();
            this.plotBuildAreas = new Dictionary<string, double>();
            this.plotTotalBuild = new Dictionary<string, double>();
            this.areasCount = 0;
            this.missingAreasCount = 0;

            ProjectInfo projectInfo = activeDoc.ProjectInformation;

            FilteredElementCollector areasCollector = new FilteredElementCollector(activeDoc).OfCategory(BuiltInCategory.OST_Areas).WhereElementIsNotElementType();

            // construct main AreaOrganizer Dictionary

            foreach (Area area in areasCollector)
            {
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
                            plotBuildAreas[plotName] += area.Area / areaConvert;
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
                        if (area.LookupParameter("A Instance Area Location").AsString() != "ПОДЗЕМНА" && area.LookupParameter("A Instance Area Primary").AsString() != "")
                        {
                            plotTotalBuild[plotName] += area.LookupParameter("A Instance Total Area").AsDouble() / areaConvert;
                        }
                    }
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

            // set columns' width
            workSheet.Range["A:A"].ColumnWidth = 20;
            workSheet.Range["B:B"].ColumnWidth = 55;
            workSheet.Range["C:C"].ColumnWidth = 20;
            workSheet.Range["D:D"].ColumnWidth = 20;
            workSheet.Range["E:E"].ColumnWidth = 10;
            workSheet.Range["F:N"].ColumnWidth = 5;
            workSheet.Range["O:S"].ColumnWidth = 20;
            workSheet.Range["T:V"].ColumnWidth = 10;

            int x = 1;

            // general formatting
            // main title : IPID and project number
            workSheet.Cells[x, 1] = "IPID";
            workSheet.Cells[x, 2] = doc.ProjectInformation.LookupParameter("Project Number").AsString();

            Range mergeRange = workSheet.Range[$"B{x}", $"V{x}"];
            mergeRange.Merge();
            mergeRange.Borders.LineStyle = XlLineStyle.xlContinuous;
            mergeRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            mergeRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

            Range ipIdRange = workSheet.Range[$"A{x}", $"A{x}"];
            ipIdRange.Borders.LineStyle= XlLineStyle.xlContinuous;
            ipIdRange.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            ipIdRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

            // main title : project name
            x += 2;
            workSheet.Cells[x, 1] = "ОБЕКТ";
            workSheet.Cells[x, 2] = doc.ProjectInformation.LookupParameter("Project Address").AsString();

            Range mergeRangeObject = workSheet.Range[$"B{x}", $"V{x}"];
            mergeRangeObject.Merge();
            mergeRangeObject.Borders.LineStyle = XlLineStyle.xlContinuous;
            mergeRangeObject.HorizontalAlignment = XlHAlign.xlHAlignLeft;
            mergeRangeObject.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.White);

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
                Range plotRange = workSheet.Range[$"A{x}", $"V{x}"];
                object[] plotStrings = new[] { "УПИ:", $"{Math.Round(plotAreasImp[plotName] / areaConvert, 3)}", "m2", "", "Самостоятелни обекти и паркоместа:", "", "", "", "", "", "", "Обекти на терен:", "", "", "", "", "", "Забележки:", "", "", "", ""};
                plotRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, plotStrings);

                // build up area row
                x++;
                Range baRange = workSheet.Range[$"A{x}", $"V{x}"];
                object[] baStrings = new[] { "ЗП:", $"{Math.Round(plotBuildAreas[plotName], 3)}", "m2", "", "Ателиета:", "", "", "", "0", "бр", "", "Паркоместа:", "", "", "0", "бр", "", "За целите на ценообразуването и площообразуването, от площта на общите части са приспаднати ХХ.ХХкв.м. :", "", "", "", ""};
                baRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, baStrings);

                // total build area row
                x++;
                Range tbaRange = workSheet.Range[$"A{x}", $"V{x}"];
                string[] tbaStrings = new[] { "РЗП:", $"{Math.Round(plotTotalBuild[plotName], 3)}", "m2", "", "Апартаменти:", "", "", "", "0", "бр", "", "Дворове:", "", "", "0", "бр", "", "", "", "", "", "" };
                tbaRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, tbaStrings);

                // underground row
                x++;
                Range uRange = workSheet.Range[$"A{x}", $"V{x}"];
                string[] uStrings = new[] { "Сутерени:", "X", "m2", "", "Магазини:", "", "", "", "0", "бр", "", "Трафопост:", "", "", "0", "бр", "", "", "", "", "", "" };
                uRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, uStrings);

                // underground + tba row
                x++;
                Range utbaRange = workSheet.Range[$"A{x}", $"V{x}"];
                string[] utbaStrings = new[] { "РЗП + Сутерени:", "X", "m2", "", "Офиси", "", "", "", "0", "бг", "", "", "", "", "", "", "", "", "", "", "", "" };
                utbaRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, utbaStrings);

                // CO row
                x++;
                Range coRange = workSheet.Range[$"A{x}", $"V{x}"];
                string[] coStrings = new[] { "Общо СО", "X", "m2", "", "Гаражи", "", "", "", "0", "бр", "", "Данни за обекта:", "", "", "", "", "", "", "", "", "", "" };
                coRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, coStrings);

                // CA row
                x++;
                Range caRange = workSheet.Range[$"A{x}", $"V{x}"];
                string[] caStrings = new[] { "Общо ОЧ", "X", "m2", "", "Складове", "", "", "", "0", "бр", "", "Етажност", "", "", "ет", "", "", "", "", "", "", "" };
                caRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, caStrings);

                // land row
                x += 1;
                Range landRange = workSheet.Range[$"A{x}", $"V{x}"];
                string[] landStrings = new[] { "Земя към СО:", "X", "m2", "", "Паркоместа", "", "", "", "0", "бр", "", "Система", "", "монолитна", "", "", "", "", "", "", "", "" };
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

                foreach (string property in plotProperties[plotName])
                {
                    x += 2;
                    workSheet.Cells[x, 1] = $"ПЛОЩООБРАЗУВАНЕ САМОСТОЯТЕЛНИ ОБЕКТИ - {property}";
                    Range propertyTitleRange = workSheet.Range[$"A{x}", $"V{x}"];
                    propertyTitleRange.Merge();
                    propertyTitleRange.HorizontalAlignment = XlHAlign.xlHAlignCenter;
                    propertyTitleRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    propertyTitleRange.Font.Bold = true;
                    propertyTitleRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                    x++;
                    Range propertyDataRange = workSheet.Range[$"A{x}", $"V{x}"];
                    string[] propertyData = new[] { "ПЛОЩ СО:", "", "ПЛОЩ ОЧ:", "", "ОБЩО:", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "", "" };
                    propertyDataRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, propertyData);
                    propertyDataRange.Font.Bold = true;
                    propertyDataRange.Borders.LineStyle= XlLineStyle.xlContinuous;
                    propertyDataRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    x++;
                    Range parameterNamesRange = workSheet.Range[$"A{x}", $"V{x}"];
                    string[] parameterNamesData = new[] { "ПЛОЩ СО", "НАИМЕНОВАНИЕ СО", "ПЛОЩ F1(F2)", "ПРИЛЕЖАЩА ПЛОЩ", "Кпг", "Ки", "Кв", "Км", "Кив", "Кпп", "Кок", "Ксп", "Кк", "К", "C1(C2)", "ОБЩИ ЧАСТИ - F3", "", "ОБЩО-F1(F2)+F3", "ПРАВО НА СТРОЕЖ", "ЗЕМЯ", "", "Id" };
                    parameterNamesRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, parameterNamesData);
                    parameterNamesRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    parameterNamesRange.Font.Bold = true;
                    parameterNamesRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                    x++;
                    Range parametersTypeRange = workSheet.Range[$"A{x}", $"V{x}"];
                    string[] parametersTypeData = new[] { "", "", "m2", "m2", "", "", "", "", "", "", "", "", "", "", "", "% и.ч.", "m2", "m2", "% и.ч.", "% и.ч.", "m2"};
                    parametersTypeRange.set_Value(XlRangeValueDataType.xlRangeValueDefault, parametersTypeData);
                    parametersTypeRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    parametersTypeRange.Font.Bold = true;
                    parametersTypeRange.Borders.LineStyle = XlLineStyle.xlContinuous;

                    x++;
                    Range blankLineRange = workSheet.Range[$"A{x}", $"V{x}"];
                    blankLineRange.Merge();
                    blankLineRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                    Borders blankBorders = blankLineRange.Borders;
                    blankBorders[XlBordersIndex.xlEdgeLeft].LineStyle = XlLineStyle.xlContinuous;
                    blankBorders[XlBordersIndex.xlEdgeTop].LineStyle = XlLineStyle.xlContinuous;
                    blankBorders[XlBordersIndex.xlEdgeRight].LineStyle = XlLineStyle.xlContinuous;
                    blankBorders[XlBordersIndex.xlEdgeBottom].LineStyle = XlLineStyle.xlContinuous;

                    x += 2;

                    int startLine = x;

                    // TODO: Solve Areas level sorting (AR-FP-01 going in the end) (and sort properly, it is currently temporary solution)

                    List<Area> sortedAreas = AreasOrganizer[plotName][property]
                        .Where(area => !area.LookupParameter("Number").AsString().Contains("ОЧ"))
                        .Where(area => !(area.LookupParameter("A Instance Area Group").AsString().Equals("ЗЕМЯ") && area.LookupParameter("A Instance Area Primary").HasValue))
                        .OrderBy(area => area.LookupParameter("A Instance Area Entrance").AsString())
                        .ThenBy(area => area.LookupParameter("Level").AsString())
                        .ThenBy(area => area.LookupParameter("Number").AsString()).ToList();

                    // TODO: Solve Areas level sorting (AR-FP-01 going in the end) (and sort properly, it is currently temporary solution)

                    List<string> levels = new List<string>();
                    List<string> entrances = new List<string>();
                                        

                    foreach (Area area in sortedAreas)
                    {
                        // entrance and level mark

                        if (!entrances.Contains(area.LookupParameter("A Instance Area Entrance").AsString()))
                        {
                            entrances.Add(area.LookupParameter("A Instance Area Entrance").AsString());
                            workSheet.Cells[x, 1] = area.LookupParameter("A Instance Area Entrance").AsString();
                            Range entranceRangeString = workSheet.Range[$"A{x}", $"V{x}"];
                            entranceRangeString.Merge();
                            entranceRangeString.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            x++;
                        }

                        if (!levels.Contains(area.LookupParameter("Level").AsValueString()))
                        {
                            levels.Add(area.LookupParameter("Level").AsValueString());
                            workSheet.Cells[x, 1] = area.LookupParameter("Level").AsValueString();
                            Range levelsRangeString = workSheet.Range[$"A{x}", $"V{x}"];
                            levelsRangeString.Merge();
                            levelsRangeString.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);
                            x++;
                        }

                        try
                        {
                            Range cellRangeString = workSheet.Range[$"A{x}", $"B{x}"];
                            Range cellRangeDouble = workSheet.Range[$"C{x}", $"V{x}"];

                            // TODO: check them all once again in compliance with the chart structure
                            string areaNumber = area.LookupParameter("Number").AsString() ?? "SOMETHING'S WRONG";
                            string areaName = area.LookupParameter("Name")?.AsString() ?? "SOMETHING'S WRONG";
                            double areaArea = Math.Round(area.LookupParameter("A Instance Total Area")?.AsDouble() / areaConvert ?? 0.0, 3);
                            // TODO: rework properly for subjectivated area
                            object areaSubjected = DBNull.Value;
                            // TODO: rework properly for subjectivated area
                            double ACGA = Math.Round(area.LookupParameter("A Coefficient Garage (Кпг)")?.AsDouble() ?? 0.0, 3);
                            double ACOR = Math.Round(area.LookupParameter("A Coefficient Orientation (Ки)")?.AsDouble() ?? 0.0, 3);
                            double ACLE = Math.Round(area.LookupParameter("A Coefficient Level (Кв)")?.AsDouble() ?? 0.0, 3);
                            double ACLO = Math.Round(area.LookupParameter("A Coefficient Location (Км)")?.AsDouble() ?? 0.0, 3);
                            double ACHE = Math.Round(area.LookupParameter("A Coefficient Height (Кив)")?.AsDouble() ?? 0.0, 3);
                            double ACRO = Math.Round(area.LookupParameter("A Coefficient Roof (Кпп)")?.AsDouble() ?? 0.0, 3);
                            double ACSP = Math.Round(area.LookupParameter("A Coefficient Special (Кок)")?.AsDouble() ?? 0.0, 3);
                            double ACST = Math.Round(area.LookupParameter("A Coefficient Storage (Ксп)")?.AsDouble() ?? 0.0, 3);
                            double ACZO = Math.Round(area.LookupParameter("A Coefficient Zones (Кк)")?.AsDouble() ?? 0.0, 3);
                            double ACCO = Math.Round(area.LookupParameter("A Coefficient Multiplied")?.AsDouble() ?? 0.0, 3);
                            double C1C2 = Math.Round(area.LookupParameter("A Instance Price C1/C2")?.AsDouble() ?? 0.0, 3);
                            double areaCommonPercent = Math.Round(area.LookupParameter("A Instance Common Area %")?.AsDouble() ?? 0.0, 3);
                            double areaCommonArea = Math.Round(area.LookupParameter("A Instance Common Area")?.AsDouble() / areaConvert ?? 0.0, 3);
                            double areaTotalArea = Math.Round((area.LookupParameter("A Instance Total Area")?.AsDouble() / areaConvert ?? 0.0) + (area.LookupParameter("A Instance Common Area")?.AsDouble() / areaConvert ?? 0.0), 3);
                            double areaPermitPercent = Math.Round(area.LookupParameter("A Instance Building Permit %")?.AsDouble() ?? 0.0, 3);
                            double areaRLPPercentage = Math.Round(area.LookupParameter("A Instance RLP Area %")?.AsDouble() ?? 0.0, 3);
                            double areaRLP = Math.Round(area.LookupParameter("A Instance RLP Area")?.AsDouble() / areaConvert ?? 0.0, 3);
                            int integerValue = area.Id.IntegerValue;
                            double areaID = integerValue;
                            // TODO: check them all once again in compliance with the chart structure

                            string[] areaStringData = new[] { areaNumber, areaName };
                            object[] areasDoubleData = new object[] { areaArea, areaSubjected, ACGA, ACOR, ACLE, ACLO, ACHE, ACRO, ACSP, ACST, ACZO, ACCO, C1C2, areaCommonPercent, areaCommonArea, areaTotalArea, areaPermitPercent, areaRLPPercentage, areaRLP, areaID };

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

                        x ++;

                        // adjascent areas loop

                        foreach (Area areaSub in sortedAreas)
                        {
                            string primaryArea = areaSub.LookupParameter("A Instance Area Primary").AsString();

                            if (primaryArea != null && primaryArea.Equals(area.LookupParameter("Number").AsString()))
                            {
                                Range areaAdjRangeStr = workSheet.Range[$"A{x}", $"B{x}"];
                                areaAdjRangeStr.set_Value(XlRangeValueDataType.xlRangeValueDefault, new[] { areaSub.LookupParameter("Number").AsString(), areaSub.LookupParameter("Name").AsString() });

                                areaAdjRangeStr.HorizontalAlignment = XlHAlign.xlHAlignRight;
                                areaAdjRangeStr.Borders.LineStyle= XlLineStyle.xlContinuous;
                                
                                Range areaAdjRangeDouble = workSheet.Range[$"C{x}", $"V{x}"];
                                areaAdjRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, new object[] {DBNull.Value, Math.Round(areaSub.LookupParameter("Area").AsDouble() / areaConvert, 3), DBNull.Value, DBNull.Value,
                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value });

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

                                    Range areaAdjRangeDouble = workSheet.Range[$"C{x}", $"V{x}"];
                                    areaAdjRangeDouble.set_Value(XlRangeValueDataType.xlRangeValueDefault, new object[] {DBNull.Value, Math.Round(areaGround.LookupParameter("Area").AsDouble() / areaConvert, 3), DBNull.Value, DBNull.Value,
                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value,
                                        DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value, DBNull.Value });

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

                    int endLine = x-1;

                    Range colorRange = workSheet.Range[$"C{startLine}", $"U{endLine}"];
                    colorRange.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.DarkSeaGreen);

                    // set a formula for the total area sum of F1/F2
                    Range sumF1F2 = workSheet.Range[$"C{x}", $"C{x}"];
                    sumF1F2.Formula = $"=SUM(C{startLine+2}:C{endLine})";

                    // set a formula for the total sum of adjascent areas
                    Range sumAdjascent = workSheet.Range[$"D{x}", $"D{x}"];
                    sumAdjascent.Formula = $"=SUM(D{startLine+2}:D{endLine})";

                    // set a formula for the total sum of C1/C2
                    Range sumC1C2 = workSheet.Range[$"O{x}", $"O{x}"];
                    sumC1C2.Formula = $"=SUM(O{startLine + 2}:O{endLine})";

                    // set a formula for the total sum of Common Areas Percentage
                    Range sumCommonPercent = workSheet.Range[$"P{x}", $"P{x}"];
                    sumCommonPercent.Formula = $"=SUM(P{startLine + 2}:P{endLine})";

                    // set a formula for the total sum of Common Areas Percentage
                    Range sumIdealPercent = workSheet.Range[$"P{x}", $"P{x}"];
                    sumIdealPercent.Formula = $"=SUM(P{startLine + 2}:P{endLine})";

                    // set a formula for the total sum of Common Areas Percentage
                    Range sumIdealArea = workSheet.Range[$"Q{x}", $"Q{x}"];
                    sumIdealArea.Formula = $"=SUM(Q{startLine + 2}:Q{endLine})";

                    // set a formula for the total sum of Common Areas Percentage
                    Range sumF1F2F3 = workSheet.Range[$"R{x}", $"R{x}"];
                    sumF1F2F3.Formula = $"=SUM(R{startLine + 2}:R{endLine})";

                    // set a formula for the total sum of Common Areas Percentage
                    Range buildingRights = workSheet.Range[$"S{x}", $"S{x}"];
                    buildingRights.Formula = $"=SUM(S{startLine + 2}:S{endLine})";

                    // set a formula for the total sum of Common Areas Percentage
                    Range landPercent = workSheet.Range[$"T{x}", $"T{x}"];
                    landPercent.Formula = $"=SUM(T{startLine + 2}:T{endLine})";

                    // set a formula for the total sum of Common Areas Percentage
                    Range landArea = workSheet.Range[$"U{x}", $"U{x}"];
                    landArea.Formula = $"=SUM(U{startLine + 2}:U{endLine})";

                    // set coloring for the summed up rows
                    Range colorRangePropertySum = workSheet.Range[$"A{endLine + 1}", $"V{endLine + 1}"];
                    colorRangePropertySum.Interior.Color = ColorTranslator.ToOle(System.Drawing.Color.LightGray);

                    x++;
                }
            }

            workBook.Save();
            workBook.Close();
        }
    }
}
