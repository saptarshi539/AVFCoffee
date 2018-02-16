using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class Inputs_1_metric_currency
    {
        public void Inputs1MetricCurrency(ExcelFile xls)
        {
            //xls.NewFile(20, TExcelFileFormat.v2016);    //Create a new Excel file with 20 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Inputs 1.0";
            xls.ActiveSheet = 2;
            xls.SheetName = "Outcome 1.0";
            xls.ActiveSheet = 3;
            xls.SheetName = "DATABASE_Schema";
            xls.ActiveSheet = 4;
            xls.SheetName = "Outcome TOTAL_Adj";
            xls.ActiveSheet = 5;
            xls.SheetName = "Outcome_Y_Adjustment";
            xls.ActiveSheet = 6;
            xls.SheetName = "Outcome_L Adjustment";
            xls.ActiveSheet = 7;
            xls.SheetName = "Proportions";
            xls.ActiveSheet = 8;
            xls.SheetName = "Inputs advanced";
            xls.ActiveSheet = 9;
            xls.SheetName = "Budget_Supuestos";
            xls.ActiveSheet = 10;
            xls.SheetName = "Budget_Equipo";
            xls.ActiveSheet = 11;
            xls.SheetName = "Budget_M Obra";
            xls.ActiveSheet = 12;
            xls.SheetName = "Budget_Presupuesto";
            xls.ActiveSheet = 13;
            xls.SheetName = "Budget_Valor de M Obra";
            xls.ActiveSheet = 14;
            xls.SheetName = "Budget_Establecimiento";
            xls.ActiveSheet = 15;
            xls.SheetName = "Budget_Sostenemiento";
            xls.ActiveSheet = 16;
            xls.SheetName = "Inputs 1.0_metric_currency";
            xls.ActiveSheet = 17;
            xls.SheetName = "Outcome 1.0 pre_metric_currency";
            xls.ActiveSheet = 18;
            xls.SheetName = "Conversiones";
            xls.ActiveSheet = 19;
            xls.SheetName = "Proporción de productividad";
            xls.ActiveSheet = 20;
            xls.SheetName = "Inputs 1.0 (Ref)";

            xls.ActiveSheet = 16;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Inputs 1.0_metric_currency";

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
            Range = new TXlsNamedRange(RangeName, 14, 32, "=Budget_Establecimiento!$A$3:$C$53");
            //You could also use: Range = new TXlsNamedRange(RangeName, 14, 14, 3, 1, 53, 3, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 11, 32, "='Budget_M Obra'!$A$1:$K$86");
            //You could also use: Range = new TXlsNamedRange(RangeName, 11, 11, 1, 1, 86, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 12, 32, "=Budget_Presupuesto!$A$34:$J$46");
            //You could also use: Range = new TXlsNamedRange(RangeName, 12, 12, 34, 1, 46, 10, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 15, 32, "=Budget_Sostenemiento!$A$1:$K$44");
            //You could also use: Range = new TXlsNamedRange(RangeName, 15, 15, 1, 1, 44, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 9, 32, "=Budget_Supuestos!$A$276:$G$297");
            //You could also use: Range = new TXlsNamedRange(RangeName, 9, 9, 276, 1, 297, 7, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 13, 32, "='Budget_Valor de M Obra'!$A$2:$J$85");
            //You could also use: Range = new TXlsNamedRange(RangeName, 13, 13, 2, 1, 85, 10, 32);
            xls.SetNamedRange(Range);


            //Printer Settings
            xls.PrintXResolution = 600;
            xls.PrintYResolution = 600;
            xls.PrintOptions = TPrintOptions.Orientation;
            xls.PrintPaperSize = TPaperSize.Letter;

            //Set up rows and columns
            xls.DefaultColWidth = 2261;

            xls.SetColWidth(3, 3, 8064);    //(30.75 + 0.75) * 256

            xls.SetColWidth(4, 4, 2261);    //(8.08 + 0.75) * 256

            xls.SetColWidth(8, 8, 4309);    //(16.08 + 0.75) * 256

            xls.SetColWidth(9, 9, 2858);    //(10.41 + 0.75) * 256

            xls.SetColWidth(10, 10, 3925);    //(14.58 + 0.75) * 256

            xls.SetRowHeight(14, 620);    //31.00 * 20
            xls.SetRowHeight(15, 930);    //46.50 * 20
            xls.SetRowHeight(17, 930);    //46.50 * 20
            xls.SetRowHeight(19, 620);    //31.00 * 20
            xls.SetRowHeight(24, 620);    //31.00 * 20

            //Set the cell values
            xls.SetCellValue(5, 8, "From inputs (Soles)");
            xls.SetCellValue(5, 9, "Dollars");
            xls.SetCellValue(5, 10, "Mexican Pesos");
            xls.SetCellValue(6, 3, "Hectares of tree early production");

            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));
            xls.SetCellValue(6, 4, new TFormula("='Inputs 1.0'!E6"));
            xls.SetCellValue(7, 3, "Hectares of trees on peak production");

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));
            xls.SetCellValue(7, 4, new TFormula("='Inputs 1.0'!E7"));
            xls.SetCellValue(8, 3, "Hectares old trees");

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));
            xls.SetCellValue(8, 4, new TFormula("='Inputs 1.0'!E8"));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));
            xls.SetCellValue(10, 3, "Conventional");

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));
            xls.SetCellValue(10, 4, new TFormula("='Inputs 1.0'!E10"));
            xls.SetCellValue(11, 3, "Organic ");

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));
            xls.SetCellValue(11, 4, new TFormula("='Inputs 1.0'!E11"));
            xls.SetCellValue(12, 3, "Transition");

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));
            xls.SetCellValue(12, 4, new TFormula("='Inputs 1.0'!E12"));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));
            xls.SetCellValue(14, 3, "How much do you pay per day to your workers on average?");

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, new TFormula("=J14"));

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));
            xls.SetCellValue(14, 8, new TFormula("='Inputs 1.0'!E14"));
            xls.SetCellValue(14, 9, new TFormula("=H14/Conversiones!$D$24"));
            xls.SetCellValue(14, 10, new TFormula("=I14*Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));
            xls.SetCellValue(15, 3, "How many quintales of coffee do you produce on average in one year per hectare?");

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));
            xls.SetCellValue(15, 4, new TFormula("='Inputs 1.0'!$E$15"));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));
            xls.SetCellValue(17, 3, "How much do you pay in pesos to transport your coffee  from the farm to the collection"
            + " center in one year? ");

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));
            xls.SetCellValue(17, 4, new TFormula("=J17"));

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));
            xls.SetCellValue(17, 8, new TFormula("='Inputs 1.0'!E17"));
            xls.SetCellValue(17, 9, new TFormula("=H17/Conversiones!$D$24"));
            xls.SetCellValue(17, 10, new TFormula("=I17*Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));
            xls.SetCellValue(19, 3, "What price did you received per quintal of coffee?");

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));
            xls.SetCellValue(19, 4, new TFormula("=J19"));

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));
            xls.SetCellValue(19, 8, new TFormula("='Inputs 1.0'!E19"));
            xls.SetCellValue(19, 9, new TFormula("=H19/Conversiones!$D$24"));
            xls.SetCellValue(19, 10, new TFormula("=I19*Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));
            xls.SetCellValue(24, 3, "Chemical fertilizers (soles per hectare)");

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));
            xls.SetCellValue(24, 4, new TFormula("=J24"));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));
            xls.SetCellValue(24, 8, new TFormula("='Inputs 1.0'!E24"));
            xls.SetCellValue(24, 9, new TFormula("=H24/Conversiones!$D$24"));
            xls.SetCellValue(24, 10, new TFormula("=I24*Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));
            xls.SetCellValue(25, 3, "Organic fertilizers (soles per hectare)");

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.WrapText = true;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));
            xls.SetCellValue(25, 4, new TFormula("=J25"));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));
            xls.SetCellValue(25, 8, new TFormula("='Inputs 1.0'!E25"));
            xls.SetCellValue(25, 9, new TFormula("=H25/Conversiones!$D$24"));
            xls.SetCellValue(25, 10, new TFormula("=I25*Conversiones!$F$24"));

            //Cell selection and scroll position.
            xls.SelectCell(25, 4, false);
            xls.ScrollWindow(14, 1);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "Juan Hernandez");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-07T22:31:31Z");

        }


    }
}
