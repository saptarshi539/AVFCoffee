using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class Inputs_1_Ref
    {
        public void inputs1Ref(ExcelFile xls)
        {
            xls.NewFile(31, TExcelFileFormat.v2010);    //Create a new Excel file with 31 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Metrics";
            xls.ActiveSheet = 2;
            xls.SheetName = "Inputs 1.0";
            xls.ActiveSheet = 3;
            xls.SheetName = "Inputs advance 2.0 (eng)";
            xls.ActiveSheet = 4;
            xls.SheetName = "Outcome 1.0";
            xls.ActiveSheet = 5;
            xls.SheetName = "Additional 2.0";
            xls.ActiveSheet = 6;
            xls.SheetName = "Fixed 2.0";
            xls.ActiveSheet = 7;
            xls.SheetName = "Variable 2.0";
            xls.ActiveSheet = 8;
            xls.SheetName = "General Budget 2.0";
            xls.ActiveSheet = 9;
            xls.SheetName = "DATABASE_Schema";
            xls.ActiveSheet = 10;
            xls.SheetName = "Inputs 2.0 Conv. default values";
            xls.ActiveSheet = 11;
            xls.SheetName = "Inputs 2.0 Conv. new inputs";
            xls.ActiveSheet = 12;
            xls.SheetName = "Inputs advanced 2.0 (esp_eng)";
            xls.ActiveSheet = 13;
            xls.SheetName = "Inputs TOT advanced";
            xls.ActiveSheet = 14;
            xls.SheetName = "Gral Conf. Summary";
            xls.ActiveSheet = 15;
            xls.SheetName = "Inputs 1.0 default values";
            xls.ActiveSheet = 16;
            xls.SheetName = "Inputs 1.0 Conv. new values";
            xls.ActiveSheet = 17;
            xls.SheetName = "Outcome TOTAL_Adj";
            xls.ActiveSheet = 18;
            xls.SheetName = "Outcome_Y_Adjustment";
            xls.ActiveSheet = 19;
            xls.SheetName = "Outcome_L Adjustment";
            xls.ActiveSheet = 20;
            xls.SheetName = "Proportions";
            xls.ActiveSheet = 21;
            xls.SheetName = "Budget_Supuestos";
            xls.ActiveSheet = 22;
            xls.SheetName = "Budget_Equipo";
            xls.ActiveSheet = 23;
            xls.SheetName = "Budget_M Obra";
            xls.ActiveSheet = 24;
            xls.SheetName = "Budget_Presupuesto";
            xls.ActiveSheet = 25;
            xls.SheetName = "Budget_Valor de M Obra";
            xls.ActiveSheet = 26;
            xls.SheetName = "Budget_Establecimiento";
            xls.ActiveSheet = 27;
            xls.SheetName = "Budget_Sostenemiento";
            xls.ActiveSheet = 28;
            xls.SheetName = "Outcome 1.0 pre_metric_currency";
            xls.ActiveSheet = 29;
            xls.SheetName = "Conversiones";
            xls.ActiveSheet = 30;
            xls.SheetName = "Proporción de productividad";
            xls.ActiveSheet = 31;
            xls.SheetName = "Inputs 1.0 (Ref)";

            xls.ActiveSheet = 31;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Inputs 1.0 (Ref)";

            //Styles.
            TFlxFormat StyleFmt;
            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Normal, 0));
            StyleFmt.Font.Size20 = 240;
            xls.SetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Normal, 0), StyleFmt);

            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Comma, 0));
            StyleFmt.Font.Size20 = 240;
            StyleFmt.Format = "_-* #,##0.00_-;\\-* #,##0.00_-;_-* \"-\"??_-;_-@_-";
            xls.SetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Comma, 0), StyleFmt);

            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0));
            StyleFmt.Font.Size20 = 240;
            xls.SetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), StyleFmt);

            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Normal, 0));
            StyleFmt.Format = "_-* #,##0.00\\ \"€\"_-;\\-* #,##0.00\\ \"€\"_-;_-* \"-\"??\\ \"€\"_-;_-@_-";
            xls.SetStyle("Currency 2", StyleFmt);

            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Followed_Hyperlink, 0));
            StyleFmt.Font.Size20 = 240;
            StyleFmt.VAlignment = TVFlxAlignment.bottom;
            StyleFmt.Locked = true;
            xls.SetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Followed_Hyperlink, 0), StyleFmt);

            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Hyperlink, 0));
            StyleFmt.Font.Size20 = 240;
            StyleFmt.VAlignment = TVFlxAlignment.bottom;
            StyleFmt.Locked = true;
            xls.SetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Hyperlink, 0), StyleFmt);

            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0));
            StyleFmt.Font.Size20 = 240;
            xls.SetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), StyleFmt);

            StyleFmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Normal, 0));
            StyleFmt.Format = "0%";
            xls.SetStyle("Percent 2", StyleFmt);

            //Named Ranges
            TXlsNamedRange Range;
            string RangeName;
            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 26, 32, "=Budget_Establecimiento!$A$3:$C$53");
            //You could also use: Range = new TXlsNamedRange(RangeName, 26, 26, 3, 1, 53, 3, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 23, 32, "='Budget_M Obra'!$A$1:$K$86");
            //You could also use: Range = new TXlsNamedRange(RangeName, 23, 23, 1, 1, 86, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 24, 32, "=Budget_Presupuesto!$A$34:$J$46");
            //You could also use: Range = new TXlsNamedRange(RangeName, 24, 24, 34, 1, 46, 10, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 27, 32, "=Budget_Sostenemiento!$A$1:$K$44");
            //You could also use: Range = new TXlsNamedRange(RangeName, 27, 27, 1, 1, 44, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 21, 32, "=Budget_Supuestos!$A$276:$G$297");
            //You could also use: Range = new TXlsNamedRange(RangeName, 21, 21, 276, 1, 297, 7, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 25, 32, "='Budget_Valor de M Obra'!$A$2:$J$85");
            //You could also use: Range = new TXlsNamedRange(RangeName, 25, 25, 2, 1, 85, 10, 32);
            xls.SetNamedRange(Range);


            //Printer Settings
            xls.PrintXResolution = 600;
            xls.PrintYResolution = 600;
            xls.PrintOptions = TPrintOptions.Orientation;
            xls.PrintPaperSize = TPaperSize.Letter;

            //Set up rows and columns
            xls.DefaultColWidth = 2261;

            xls.SetColWidth(1, 2, 2261);    //(8.08 + 0.75) * 256

            xls.SetColWidth(3, 3, 8064);    //(30.75 + 0.75) * 256

            xls.SetColWidth(4, 5, 2261);    //(8.08 + 0.75) * 256

            xls.SetColWidth(6, 6, 5888);    //(22.25 + 0.75) * 256

            xls.SetColWidth(7, 7, 2176);    //(7.75 + 0.75) * 256

            xls.SetColWidth(8, 8, 3242);    //(11.91 + 0.75) * 256

            xls.SetColWidth(9, 9, 1578);    //(5.41 + 0.75) * 256

            xls.SetColWidth(10, 10, 3328);    //(12.25 + 0.75) * 256

            xls.SetColWidth(11, 16384, 2261);    //(8.08 + 0.75) * 256

            xls.SetRowHeight(15, 900);    //45.00 * 20
            xls.SetRowHeight(17, 900);    //45.00 * 20
            xls.SetRowHeight(19, 600);    //30.00 * 20
            xls.SetRowHeight(27, 320);    //16.00 * 20
            xls.SetRowHeight(28, 320);    //16.00 * 20
            xls.SetRowHeight(29, 320);    //16.00 * 20
            xls.SetRowHeight(30, 320);    //16.00 * 20
            xls.SetRowHeight(31, 320);    //16.00 * 20

            //Set the cell values
            xls.SetCellValue(6, 3, "Hectares young trees");
            xls.SetCellValue(6, 4, 1.03);
            xls.SetCellValue(7, 3, "Hectares mature trees");
            xls.SetCellValue(7, 4, 1.94);
            xls.SetCellValue(8, 3, "Hectares old trees");
            xls.SetCellValue(8, 4, 1.97);
            xls.SetCellValue(10, 3, "Chemical");
            xls.SetCellValue(10, 4, 1);
            xls.SetCellValue(11, 3, "Organic ");
            xls.SetCellValue(11, 4, 0);
            xls.SetCellValue(12, 3, "Transition");
            xls.SetCellValue(12, 4, 0);
            xls.SetCellValue(14, 3, "Salary");
            xls.SetCellValue(14, 4, 93.1);
            xls.SetCellValue(14, 8, "Pounds/ha");

            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));
            xls.SetCellValue(15, 3, "How many quintales of coffee do you produce on average in one year per hectare?");
            xls.SetCellValue(15, 4, 14);
            xls.SetCellValue(15, 8, new TFormula("=D15*Conversiones!C14"));

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));
            xls.SetCellValue(17, 3, "How much do you pay in pesos to transport your coffee  from the farm to the collection"
            + " center in one year? ");
            xls.SetCellValue(17, 4, 1355.5);

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));
            xls.SetCellValue(19, 3, "What price did you received per quintal of coffee?");

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));
            xls.SetCellValue(19, 4, 3207);
            xls.SetCellValue(23, 7, "Mexico");
            xls.SetCellValue(23, 8, "Colombia");
            xls.SetCellValue(23, 9, "Peru");
            xls.SetCellValue(23, 10, "Honduras");
            xls.SetCellValue(23, 11, "Colombia");
            xls.SetCellValue(24, 7, "Cesmach");
            xls.SetCellValue(24, 8, "Andes");
            xls.SetCellValue(24, 9, "ADISA");
            xls.SetCellValue(24, 10, "COMSA-Parch.");
            xls.SetCellValue(24, 11, "FCC");
            xls.SetCellValue(25, 6, "Productivity (Pounds/ht)");
            xls.SetCellValue(25, 7, 1168);
            xls.SetCellValue(25, 8, 5107.46);

            fmt = xls.GetCellVisibleFormatDef(25, 9);
            fmt.Font.Size20 = 220;
            xls.SetCellFormat(25, 9, xls.AddFormat(fmt));
            xls.SetCellValue(25, 9, 3565);

            fmt = xls.GetCellVisibleFormatDef(25, 10);
            fmt.Format = "0";
            xls.SetCellFormat(25, 10, xls.AddFormat(fmt));
            xls.SetCellValue(25, 10, 4365.72478643215);

            fmt = xls.GetCellVisibleFormatDef(25, 11);
            fmt.Format = "0.0";
            xls.SetCellFormat(25, 11, xls.AddFormat(fmt));
            xls.SetCellValue(25, 11, 2588.49979975634);
            xls.SetCellValue(27, 6, "Cost (US/ht)");

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));
            xls.SetCellValue(28, 6, "Variable ");

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));
            xls.SetCellValue(28, 7, 1127);

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));
            xls.SetCellValue(28, 8, 6752);

            fmt = xls.GetCellVisibleFormatDef(28, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xC5, 0xE0, 0xB3);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 9, xls.AddFormat(fmt));
            xls.SetCellValue(28, 9, 2988);

            fmt = xls.GetCellVisibleFormatDef(28, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xC5, 0xE0, 0xB3);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 10, xls.AddFormat(fmt));
            xls.SetCellValue(28, 10, 3898);

            fmt = xls.GetCellVisibleFormatDef(28, 11);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 11, xls.AddFormat(fmt));
            xls.SetCellValue(28, 11, 1817);
            xls.SetCellValue(29, 6, "Fixed");

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));
            xls.SetCellValue(29, 7, 1146);

            fmt = xls.GetCellVisibleFormatDef(29, 8);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 8, xls.AddFormat(fmt));
            xls.SetCellValue(29, 8, 7160);

            fmt = xls.GetCellVisibleFormatDef(29, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xE2, 0xEF, 0xD9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 9, xls.AddFormat(fmt));
            xls.SetCellValue(29, 9, 3212);

            fmt = xls.GetCellVisibleFormatDef(29, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xE2, 0xEF, 0xD9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 10, xls.AddFormat(fmt));
            xls.SetCellValue(29, 10, 4086);

            fmt = xls.GetCellVisibleFormatDef(29, 11);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 11, xls.AddFormat(fmt));
            xls.SetCellValue(29, 11, 1960);
            xls.SetCellValue(30, 6, "Depreciation ");

            fmt = xls.GetCellVisibleFormatDef(30, 7);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(30, 7, xls.AddFormat(fmt));
            xls.SetCellValue(30, 7, 1765);

            fmt = xls.GetCellVisibleFormatDef(30, 8);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(30, 8, xls.AddFormat(fmt));
            xls.SetCellValue(30, 8, 7511);

            fmt = xls.GetCellVisibleFormatDef(30, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xC5, 0xE0, 0xB3);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(30, 9, xls.AddFormat(fmt));
            xls.SetCellValue(30, 9, 4044);

            fmt = xls.GetCellVisibleFormatDef(30, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xC5, 0xE0, 0xB3);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(30, 10, xls.AddFormat(fmt));
            xls.SetCellValue(30, 10, 5283);

            fmt = xls.GetCellVisibleFormatDef(30, 11);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(30, 11, xls.AddFormat(fmt));
            xls.SetCellValue(30, 11, 2276);
            xls.SetCellValue(31, 6, "Total");

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));
            xls.SetCellValue(31, 7, 2263);

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));
            xls.SetCellValue(31, 8, 9283);

            fmt = xls.GetCellVisibleFormatDef(31, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xE2, 0xEF, 0xD9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(31, 9, xls.AddFormat(fmt));
            xls.SetCellValue(31, 9, 4658);

            fmt = xls.GetCellVisibleFormatDef(31, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xE2, 0xEF, 0xD9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(31, 10, xls.AddFormat(fmt));
            xls.SetCellValue(31, 10, 6027);

            fmt = xls.GetCellVisibleFormatDef(31, 11);
            fmt.Font.Size20 = 220;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(31, 11, xls.AddFormat(fmt));
            xls.SetCellValue(31, 11, 2846);
            xls.SetCellValue(35, 7, "Mexico");
            xls.SetCellValue(35, 8, "Colombia");
            xls.SetCellValue(35, 9, "Peru");
            xls.SetCellValue(35, 10, "Honduras");
            xls.SetCellValue(35, 11, "Colombia");
            xls.SetCellValue(36, 7, "Cesmach");
            xls.SetCellValue(36, 8, "Andes");
            xls.SetCellValue(36, 9, "ADISA");
            xls.SetCellValue(36, 10, "COMSA-Parch.");
            xls.SetCellValue(36, 11, "FCC");
            xls.SetCellValue(37, 6, "Productivity (Quintales/ht)");
            xls.SetCellValue(37, 7, new TFormula("=G25/Conversiones!$C$14"));
            xls.SetCellValue(37, 8, new TFormula("=H25/Conversiones!$C$14"));
            xls.SetCellValue(37, 9, new TFormula("=I25/Conversiones!$C$14"));
            xls.SetCellValue(37, 10, new TFormula("=J25/Conversiones!$C$14"));
            xls.SetCellValue(37, 11, new TFormula("=K25/Conversiones!$C$14"));
            xls.SetCellValue(39, 6, "Cost (US/ht)");

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));
            xls.SetCellValue(40, 6, "Variable ");
            xls.SetCellValue(40, 7, new TFormula("=G28*Conversiones!$F$24"));
            xls.SetCellValue(40, 8, new TFormula("=H28*Conversiones!$F$24"));
            xls.SetCellValue(40, 9, new TFormula("=I28*Conversiones!$F$24"));
            xls.SetCellValue(40, 10, new TFormula("=J28*Conversiones!$F$24"));
            xls.SetCellValue(40, 11, new TFormula("=K28*Conversiones!$F$24"));
            xls.SetCellValue(41, 6, "Fixed");
            xls.SetCellValue(41, 7, new TFormula("=G29*Conversiones!$F$24"));
            xls.SetCellValue(41, 8, new TFormula("=H29*Conversiones!$F$24"));
            xls.SetCellValue(41, 9, new TFormula("=I29*Conversiones!$F$24"));
            xls.SetCellValue(41, 10, new TFormula("=J29*Conversiones!$F$24"));
            xls.SetCellValue(41, 11, new TFormula("=K29*Conversiones!$F$24"));
            xls.SetCellValue(42, 6, "Depreciation ");
            xls.SetCellValue(42, 7, new TFormula("=G30*Conversiones!$F$24"));
            xls.SetCellValue(42, 8, new TFormula("=H30*Conversiones!$F$24"));
            xls.SetCellValue(42, 9, new TFormula("=I30*Conversiones!$F$24"));
            xls.SetCellValue(42, 10, new TFormula("=J30*Conversiones!$F$24"));
            xls.SetCellValue(42, 11, new TFormula("=K30*Conversiones!$F$24"));
            xls.SetCellValue(43, 6, "Total");
            xls.SetCellValue(43, 7, new TFormula("=G31*Conversiones!$F$24"));
            xls.SetCellValue(43, 8, new TFormula("=H31*Conversiones!$F$24"));
            xls.SetCellValue(43, 9, new TFormula("=I31*Conversiones!$F$24"));
            xls.SetCellValue(43, 10, new TFormula("=J31*Conversiones!$F$24"));
            xls.SetCellValue(43, 11, new TFormula("=K31*Conversiones!$F$24"));

            //Cell selection and scroll position.
            xls.SelectCell(6, 3, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "Juan Hernandez");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-08T03:31:31Z");

        }

    }
}
