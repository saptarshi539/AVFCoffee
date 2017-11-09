using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeInfrastructure.Flexcel
{
    public class Flexcelsum : IFlexcelsum
    {
        public ChartDataDTO getOutputFromExcel(double earlyHectares, double peakHectares, double oldHectares, bool conventional, bool organic, bool transition, double workerSalarySoles, double productionQuintales, double transportCostSoles, double costPriceSolesPerQuintal)
        {

            XlsFile xls = new XlsFile(true);
            CreateFileForSheet1(xls, earlyHectares);
            CreateFileForSheet2(xls, peakHectares, oldHectares);
            ChartDataDTO calculatedValue = CreateFileForSheet3(xls);
            

            //Save the file as XLS
            //xls.Save(openFileDialog1.FileName);
            return calculatedValue;
        }

        private void CreateFileForSheet2(XlsFile xls, Double peakHectares, Double oldHectares)
        {
            //xls.NewFile(3, TExcelFileFormat.v2016);    //Create a new Excel file with 3 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Sheet1";
            xls.ActiveSheet = 2;
            xls.SheetName = "Sheet2";
            xls.ActiveSheet = 3;
            xls.SheetName = "Sheet3";

            xls.ActiveSheet = 2;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsCheckCompatibility = false;

            //Sheet Options
            xls.SheetName = "Sheet2";

            //Printer Settings
            xls.PrintOptions = TPrintOptions.Orientation | TPrintOptions.NoPls;

            //Set the cell values
            xls.SetCellValue(9, 8, peakHectares);
            xls.SetCellValue(9, 9, oldHectares);

            //Cell selection and scroll position.
            xls.SelectCell(6, 8, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "SAPTARSHI MALLICK");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2017-11-09T15:42:06Z");

            xls.DocumentProperties.SetStandardProperty(TPropertyId.Company, "Cornell University");

            xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test1.xlsx"));
        }

        private void CreateFileForSheet1(XlsFile xls, Double earlyHectares)
        {
            xls.NewFile(3, TExcelFileFormat.v2016);    //Create a new Excel file with 3 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Sheet1";
            xls.ActiveSheet = 2;
            xls.SheetName = "Sheet2";
            xls.ActiveSheet = 3;
            xls.SheetName = "Sheet3";

            xls.ActiveSheet = 1;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsCheckCompatibility = false;

            //Printer Settings
            xls.PrintOptions = TPrintOptions.Orientation | TPrintOptions.NoPls;

            //Set the cell values
            xls.SetCellValue(9, 8, earlyHectares);

            //Cell selection and scroll position.
            xls.SelectCell(9, 8, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "SAPTARSHI MALLICK");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2017-11-09T15:42:06Z");

            xls.DocumentProperties.SetStandardProperty(TPropertyId.Company, "Cornell University");

            xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test1.xlsx"));
        }

        private ChartDataDTO CreateFileForSheet3(XlsFile xls)
        {
            //xls.NewFile(3, TExcelFileFormat.v2016);    //Create a new Excel file with 3 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Sheet1";
            xls.ActiveSheet = 2;
            xls.SheetName = "Sheet2";
            xls.ActiveSheet = 3;
            xls.SheetName = "Sheet3";

            xls.ActiveSheet = 3;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsCheckCompatibility = false;

            //Sheet Options
            xls.SheetName = "Sheet3";

            //Set the cell values
            //xls.SetCellValue(9, 9, oldHectares);
            xls.SetCellValue(9, 8, new TFormula("=Sheet1!H9 + Sheet2!H9"));
            xls.SetCellValue(9, 10, new TFormula("=Sheet1!H9 + Sheet2!H9 +Sheet2!I9"));
            xls.SetCellValue(9, 11, new TFormula("=Sheet2!H9 +Sheet2!I9"));
            xls.SetCellValue(9, 12, new TFormula("=Sheet2!I9 - Sheet2!H9"));

            //Cell selection and scroll position.
            xls.SelectCell(9, 8, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "SAPTARSHI MALLICK");


            xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test1.xlsx"));
            Double coopval1 = Convert.ToDouble(xls.GetCellValue(9, 11)) + 0.03;
            Double coopval2 = Convert.ToDouble(xls.GetCellValue(9, 12)) - 0.02;
            Double coopval3 = Convert.ToDouble(xls.GetCellValue(9, 10)) + 0.05;
            ChartDataDTO cd = new ChartDataDTO();
            List<Double> prod = new List<double>();
            List<Double> coop = new List<double>();
            coop.Add(Math.Round(coopval1, 2));
            coop.Add(Math.Round(coopval2, 2));
            coop.Add(Math.Round(coopval3, 2));
            cd.cooperative = coop;
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 11)),2));
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 12)),2));
            prod.Add(Math.Round(Convert.ToDouble(xls.GetCellValue(9, 10)),2));
            cd.producer = prod;
            

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2017-11-09T15:42:06Z");

            xls.DocumentProperties.SetStandardProperty(TPropertyId.Company, "Cornell University");

           

            return cd;//Convert.ToDouble(xls.GetCellValue(9, 12));
        }

        public String sumcells()
        {
            XlsFile xls = new XlsFile(1, TExcelFileFormat.v2016, true);

            //Enters a string into A1.

            xls.SetCellValue(1, 1, "Hello from FlexCel!");

            //Enters a number into A2.
            //Note that xls.SetCellValue(2, 1, "7") would enter a string.
            xls.SetCellValue(2, 1, 7);

            //Enter another floating point number.
            //All numbers in Excel are floating point,
            //so even if you enter an integer, it will be stored as double.
            xls.SetCellValue(3, 1, 11.3);

            //Enters a formula into A4.
            xls.SetCellValue(4, 1, new TFormula("=Sum(A2:A3)"));

            //Saves the file to the "Documents" folder.
            xls.Save(System.IO.Path.Combine(Environment.GetFolderPath(Environment.SpecialFolder.Personal), "test.xlsx"));

            return Convert.ToString(xls.GetCellValue(4, 1));
        }
    }
}
