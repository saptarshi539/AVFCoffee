using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class Proportions
    {
        public void proportions(ExcelFile xls)
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

            xls.ActiveSheet = 7;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Proportions";

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

            xls.SetColWidth(3, 3, 4906);    //(18.41 + 0.75) * 256

            xls.SetColWidth(4, 4, 2901);    //(10.58 + 0.75) * 256

            xls.SetColWidth(5, 5, 2858);    //(10.41 + 0.75) * 256

            xls.SetColWidth(6, 7, 2346);    //(8.41 + 0.75) * 256

            xls.SetColWidth(8, 8, 2986);    //(10.91 + 0.75) * 256

            xls.SetColWidth(9, 9, 3584);    //(13.25 + 0.75) * 256

            xls.SetColWidth(10, 10, 2901);    //(10.58 + 0.75) * 256

            xls.SetColWidth(11, 11, 2901);    //(10.58 + 0.75) * 256

            xls.SetColWidth(12, 12, 2858);    //(10.41 + 0.75) * 256

            xls.SetRowHeight(4, 620);    //31.00 * 20

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));
            xls.SetCellValue(4, 4, "Land by age");

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));
            xls.SetCellValue(4, 5, "Percentage of land");

            fmt = xls.GetCellVisibleFormatDef(4, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 8, xls.AddFormat(fmt));
            xls.SetCellValue(4, 8, "Reported productivity");

            fmt = xls.GetCellVisibleFormatDef(4, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 9, xls.AddFormat(fmt));
            xls.SetCellValue(4, 9, "years");

            fmt = xls.GetCellVisibleFormatDef(4, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 10, xls.AddFormat(fmt));
            xls.SetCellValue(4, 10, "Production average");

            fmt = xls.GetCellVisibleFormatDef(4, 11);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 12);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 13, xls.AddFormat(fmt));
            xls.SetCellValue(5, 3, "Hectares young trees");

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));
            xls.SetCellValue(5, 4, new TFormula("='Inputs 1.0_metric_currency'!D6"));
            xls.SetCellValue(5, 5, new TFormula("=D5/$D$8"));

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));
            xls.SetCellValue(5, 8, new TFormula("='Inputs 1.0_metric_currency'!$D$15"));
            xls.SetCellValue(5, 9, " Yr 2,3");
            xls.SetCellValue(5, 10, new TFormula("=H5/5.1"));
            xls.SetCellValue(6, 3, "Hectares mature trees");

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));
            xls.SetCellValue(6, 4, new TFormula("='Inputs 1.0_metric_currency'!D7"));
            xls.SetCellValue(6, 5, new TFormula("=D6/$D$8"));
            xls.SetCellValue(6, 9, "Yr 4,5,6");
            xls.SetCellValue(6, 10, new TFormula("=H5"));
            xls.SetCellValue(7, 3, "Hectares old trees");

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));
            xls.SetCellValue(7, 4, new TFormula("='Inputs 1.0_metric_currency'!D8"));
            xls.SetCellValue(7, 5, new TFormula("=D7/$D$8"));
            xls.SetCellValue(7, 9, "YR 7,8");
            xls.SetCellValue(7, 10, new TFormula("=H5*1.1"));

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));
            xls.SetCellValue(8, 3, "Total hectares");

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));
            xls.SetCellValue(8, 4, new TFormula("=SUM(D5:D7)"));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));
            xls.SetCellValue(8, 5, new TFormula("=SUM(E5:E7)"));

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));
            xls.SetCellValue(9, 10, new TFormula("=AVERAGE(J5:J7)"));
            xls.SetCellValue(14, 10, "Quintales");
            xls.SetCellValue(14, 11, "Kilogramos");
            xls.SetCellValue(14, 12, "Libras");
            xls.SetCellValue(15, 10, new TFormula("=J5"));
            xls.SetCellValue(15, 11, new TFormula("=J15*Conversiones!$D$14"));
            xls.SetCellValue(15, 12, new TFormula("=K15*Conversiones!$C$11"));
            xls.SetCellValue(16, 10, new TFormula("=J6"));
            xls.SetCellValue(16, 11, new TFormula("=J16*Conversiones!$D$14"));
            xls.SetCellValue(16, 12, new TFormula("=K16*Conversiones!$C$11"));
            xls.SetCellValue(17, 10, new TFormula("=J7"));
            xls.SetCellValue(17, 11, new TFormula("=J17*Conversiones!$D$14"));
            xls.SetCellValue(17, 12, new TFormula("=K17*Conversiones!$C$11"));
            xls.SetCellValue(19, 10, new TFormula("=AVERAGE(J15:J17)"));
            xls.SetCellValue(19, 11, new TFormula("=AVERAGE(K15:K17)"));
            xls.SetCellValue(19, 12, new TFormula("=AVERAGE(L15:L17)"));

            //Cell selection and scroll position.
            xls.SelectCell(5, 10, false);

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
