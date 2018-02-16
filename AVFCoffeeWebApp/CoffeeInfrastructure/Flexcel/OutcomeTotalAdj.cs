using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class OutcomeTotalAdj
    {
        public void Outcome_TOTAL_Adj(ExcelFile xls)
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

            xls.ActiveSheet = 4;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Outcome TOTAL_Adj";

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

            xls.SetColWidth(2, 4, 2261);    //(8.08 + 0.75) * 256

            xls.SetColWidth(5, 5, 3072);    //(11.25 + 0.75) * 256

            xls.SetColWidth(6, 7, 2261);    //(8.08 + 0.75) * 256

            xls.SetColWidth(8, 12, 1493);    //(5.08 + 0.75) * 256

            xls.SetColWidth(14, 14, 2304);    //(8.25 + 0.75) * 256

            xls.SetColWidth(15, 15, 6144);    //(23.25 + 0.75) * 256

            xls.SetColWidth(16, 16, 3285);    //(12.08 + 0.75) * 256

            xls.SetColWidth(17, 17, 2858);    //(10.41 + 0.75) * 256

            xls.SetColWidth(18, 18, 5888);    //(22.25 + 0.75) * 256

            xls.SetColWidth(20, 20, 4181);    //(15.58 + 0.75) * 256

            xls.SetRowHeight(4, 1050);    //52.50 * 20
            xls.SetRowHeight(5, 780);    //39.00 * 20
            xls.SetRowHeight(6, 1050);    //52.50 * 20
            xls.SetRowHeight(7, 790);    //39.50 * 20
            xls.SetRowHeight(8, 1820);    //91.00 * 20
            xls.SetRowHeight(12, 1060);    //53.00 * 20
            xls.SetRowHeight(13, 790);    //39.50 * 20
            xls.SetRowHeight(14, 1060);    //53.00 * 20
            xls.SetRowHeight(15, 800);    //40.00 * 20
            xls.SetRowHeight(16, 1820);    //91.00 * 20
            xls.SetRowHeight(18, 530);    //26.50 * 20

            //Merged Cells
            xls.MergeCells(3, 14, 3, 18);
            xls.MergeCells(12, 14, 12, 15);

            //Set the cell values
            xls.SetCellValue(2, 21, "Nota: Los costos en general se han calculado basado en una productividad promedio"
            + " sobre la cual se le pregunto al productor cuanto invertia en X o Y insumo");
            xls.SetCellValue(2, 38, "La productividad promedio en CESMACH fue relativamente baja 1,168 lbs/ht  Executive"
            + " report");
            xls.SetCellValue(2, 47, "Quintales");

            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(3, 14);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 14, xls.AddFormat(fmt));
            xls.SetCellValue(3, 14, "Table 6. Conventional breakeven return at different levels of enterprise costs assuming"
            + " average cost and productivity  (years 2 to 8)");

            fmt = xls.GetCellVisibleFormatDef(3, 15);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 16);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 17);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 18);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 18, xls.AddFormat(fmt));
            xls.SetCellValue(3, 21, "Si se quieren ver los costos a productividades distintas aplicamos una regla de tres"
            + " dada la productividad que se asume como referencia en este archivo, basado en CESMACH");
            xls.SetCellValue(3, 38, "En este archivo la referencias fueron 14 quintales, que se aproxima a 1400 lbs por"
            + " hectarea");

            fmt = xls.GetCellVisibleFormatDef(3, 47);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 47, xls.AddFormat(fmt));
            xls.SetCellValue(3, 47, 14);

            fmt = xls.GetCellVisibleFormatDef(4, 14);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 14, xls.AddFormat(fmt));
            xls.SetCellValue(4, 14, 1);

            fmt = xls.GetCellVisibleFormatDef(4, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 15, xls.AddFormat(fmt));
            xls.SetCellValue(4, 15, 3);

            fmt = xls.GetCellVisibleFormatDef(4, 16);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 16, xls.AddFormat(fmt));
            xls.SetCellValue(4, 16, "Costo producción cereza (Pesos/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(4, 17);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 17, xls.AddFormat(fmt));
            xls.SetCellValue(4, 17, "Breakeven -  Retorno (Pesos/quintal)");

            fmt = xls.GetCellVisibleFormatDef(4, 18);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 18, xls.AddFormat(fmt));
            xls.SetCellValue(4, 18, "Breakeven Implications");

            fmt = xls.GetCellVisibleFormatDef(5, 14);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(5, 14, xls.AddFormat(fmt));
            xls.SetCellValue(5, 14, 1);

            fmt = xls.GetCellVisibleFormatDef(5, 15);
            fmt.Font.Size20 = 200;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 15, xls.AddFormat(fmt));
            xls.SetCellValue(5, 15, "Total Variable Costs");

            fmt = xls.GetCellVisibleFormatDef(5, 16);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(5, 16, xls.AddFormat(fmt));
            xls.SetCellValue(5, 16, new TFormula("=Outcome_Y_Adjustment!I7+'Outcome_L Adjustment'!J3"));

            fmt = xls.GetCellVisibleFormatDef(5, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(5, 17, xls.AddFormat(fmt));
            xls.SetCellValue(5, 17, new TFormula("=(P5/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(5, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 18, xls.AddFormat(fmt));
            xls.SetCellValue(5, 18, "If the return is below this level, coffee is uneconomical to produce.");

            fmt = xls.GetCellVisibleFormatDef(6, 14);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(6, 14, xls.AddFormat(fmt));
            xls.SetCellValue(6, 14, 2);

            fmt = xls.GetCellVisibleFormatDef(6, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(6, 15, xls.AddFormat(fmt));
            xls.SetCellValue(6, 15, "Total Cash Costs = Total Variable Costs + Membership & Certification Costs + Taxes"
            + " on Land + Miscellaneous Supplies");

            fmt = xls.GetCellVisibleFormatDef(6, 16);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(6, 16, xls.AddFormat(fmt));
            xls.SetCellValue(6, 16, new TFormula("=Outcome_Y_Adjustment!I8+'Outcome_L Adjustment'!J4"));

            fmt = xls.GetCellVisibleFormatDef(6, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(6, 17, xls.AddFormat(fmt));
            xls.SetCellValue(6, 17, new TFormula("=(P6/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(6, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(6, 18, xls.AddFormat(fmt));
            xls.SetCellValue(6, 18, "The second breakeven return allows the producer to stay in business in the short run.");

            fmt = xls.GetCellVisibleFormatDef(7, 14);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(7, 14, xls.AddFormat(fmt));
            xls.SetCellValue(7, 14, 3);

            fmt = xls.GetCellVisibleFormatDef(7, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(7, 15, xls.AddFormat(fmt));
            xls.SetCellValue(7, 15, "Out Of Pocket Costs = Total Cash Costs + Depreciation Costs");

            fmt = xls.GetCellVisibleFormatDef(7, 16);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(7, 16, xls.AddFormat(fmt));
            xls.SetCellValue(7, 16, new TFormula("=Outcome_Y_Adjustment!I9+'Outcome_L Adjustment'!J5"));

            fmt = xls.GetCellVisibleFormatDef(7, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(7, 17, xls.AddFormat(fmt));
            xls.SetCellValue(7, 17, new TFormula("=(P7/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(7, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(7, 18, xls.AddFormat(fmt));
            xls.SetCellValue(7, 18, "The third breakeven allows the producer to stay in business in the long run.");

            fmt = xls.GetCellVisibleFormatDef(8, 14);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(8, 14, xls.AddFormat(fmt));
            xls.SetCellValue(8, 14, 4);

            fmt = xls.GetCellVisibleFormatDef(8, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(8, 15, xls.AddFormat(fmt));
            xls.SetCellValue(8, 15, " Total Costs = Out of Pocket Costs + Amortized Establishment Costs + Management Costs"
            + " + Opportunity Costs");

            fmt = xls.GetCellVisibleFormatDef(8, 16);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(8, 16, xls.AddFormat(fmt));
            xls.SetCellValue(8, 16, new TFormula("=Outcome_Y_Adjustment!I10+'Outcome_L Adjustment'!J6"));

            fmt = xls.GetCellVisibleFormatDef(8, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(8, 17, xls.AddFormat(fmt));
            xls.SetCellValue(8, 17, new TFormula("=(P8/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(8, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(8, 18, xls.AddFormat(fmt));
            xls.SetCellValue(8, 18, "The fourth breakeven return is the total cost breakeven return. Only when this breakeven"
            + " return is received can the grower recover all out-of-pocket expenses plus opportunity"
            + " costs.");

            fmt = xls.GetCellVisibleFormatDef(9, 14);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(9, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 15);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(9, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 16);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 17);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 18);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(9, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 14);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(10, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 15);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 15, xls.AddFormat(fmt));
            xls.SetCellValue(10, 15, "Precio actual en pesos Quintal:");

            fmt = xls.GetCellVisibleFormatDef(10, 16);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 16, xls.AddFormat(fmt));
            xls.SetCellValue(10, 16, new TFormula("='Outcome_L Adjustment'!$C$8"));

            fmt = xls.GetCellVisibleFormatDef(10, 17);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(10, 17, xls.AddFormat(fmt));
            xls.SetCellValue(10, 17, 1454.6724605601);

            fmt = xls.GetCellVisibleFormatDef(10, 18);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 18, xls.AddFormat(fmt));
            xls.SetCellValue(10, 18, 0.526102155717937);

            fmt = xls.GetCellVisibleFormatDef(11, 14);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(11, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 15);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(11, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 16);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 17);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 18);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(11, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 14);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(12, 14, xls.AddFormat(fmt));
            xls.SetCellValue(12, 14, "Cost definition");

            fmt = xls.GetCellVisibleFormatDef(12, 15);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(12, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 16);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 16, xls.AddFormat(fmt));
            xls.SetCellValue(12, 16, "Costo producción pergamino (US/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(12, 17);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 17, xls.AddFormat(fmt));
            xls.SetCellValue(12, 17, "Breakeven Retorno (us/pound pregamino)");

            fmt = xls.GetCellVisibleFormatDef(12, 18);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 18, xls.AddFormat(fmt));
            xls.SetCellValue(12, 18, "Breakeven Implications");

            fmt = xls.GetCellVisibleFormatDef(12, 20);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 20, xls.AddFormat(fmt));
            xls.SetCellValue(12, 20, "Costo producción pergamino (US/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
            xls.SetCellValue(13, 3, "Variable Costs");

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, "Fixed costs");

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));
            xls.SetCellValue(13, 5, "Total costs and depreciation");

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));
            xls.SetCellValue(13, 6, "Total");

            fmt = xls.GetCellVisibleFormatDef(13, 14);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(13, 14, xls.AddFormat(fmt));
            xls.SetCellValue(13, 14, 1);

            fmt = xls.GetCellVisibleFormatDef(13, 15);
            fmt.Font.Size20 = 200;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 15, xls.AddFormat(fmt));
            xls.SetCellValue(13, 15, "Total Variable Costs");

            fmt = xls.GetCellVisibleFormatDef(13, 16);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 16, xls.AddFormat(fmt));
            xls.SetCellValue(13, 16, new TFormula("=P5/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(13, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 17, xls.AddFormat(fmt));
            xls.SetCellValue(13, 17, new TFormula("=(Q5/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(13, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 18, xls.AddFormat(fmt));
            xls.SetCellValue(13, 18, "If the return is below this level, coffee is uneconomical to produce.");
            xls.SetCellValue(13, 20, new TFormula("=P13*Conversiones!$D$24"));

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));
            xls.SetCellValue(14, 2, "Productor 1");

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));
            xls.SetCellValue(14, 3, new TFormula("=+Q13"));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, new TFormula("=Q14-Q13"));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));
            xls.SetCellValue(14, 5, new TFormula("=Q16-Q14"));

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Size20 = 200;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));
            xls.SetCellValue(14, 6, new TFormula("=SUM(C14:E14)"));

            fmt = xls.GetCellVisibleFormatDef(14, 14);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(14, 14, xls.AddFormat(fmt));
            xls.SetCellValue(14, 14, 2);

            fmt = xls.GetCellVisibleFormatDef(14, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(14, 15, xls.AddFormat(fmt));
            xls.SetCellValue(14, 15, "Total Cash Costs = Total Variable Costs + Membership & Certification Costs + Taxes"
            + " on Land + Miscellaneous Supplies");

            fmt = xls.GetCellVisibleFormatDef(14, 16);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 16, xls.AddFormat(fmt));
            xls.SetCellValue(14, 16, new TFormula("=P6/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(14, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 17, xls.AddFormat(fmt));
            xls.SetCellValue(14, 17, new TFormula("=(Q6/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(14, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 18, xls.AddFormat(fmt));
            xls.SetCellValue(14, 18, "The second breakeven return allows the producer to stay in business in the short run.");
            xls.SetCellValue(14, 20, new TFormula("=P14*Conversiones!$D$24"));

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));
            xls.SetCellValue(15, 2, "Cooperative ");

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));
            xls.SetCellValue(15, 3, 1.05);

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));
            xls.SetCellValue(15, 4, 0.06);

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));
            xls.SetCellValue(15, 5, 0.8);

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Size20 = 200;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));
            xls.SetCellValue(15, 6, new TFormula("=SUM(C15:E15)"));

            fmt = xls.GetCellVisibleFormatDef(15, 14);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(15, 14, xls.AddFormat(fmt));
            xls.SetCellValue(15, 14, 3);

            fmt = xls.GetCellVisibleFormatDef(15, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(15, 15, xls.AddFormat(fmt));
            xls.SetCellValue(15, 15, "Out Of Pocket Costs = Total Cash Costs + Depreciation Costs");

            fmt = xls.GetCellVisibleFormatDef(15, 16);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 16, xls.AddFormat(fmt));
            xls.SetCellValue(15, 16, new TFormula("=P7/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(15, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 17, xls.AddFormat(fmt));
            xls.SetCellValue(15, 17, new TFormula("=(Q7/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(15, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(15, 18, xls.AddFormat(fmt));
            xls.SetCellValue(15, 18, "The third breakeven allows the producer to stay in business in the long run.");
            xls.SetCellValue(15, 20, new TFormula("=P15*Conversiones!$D$24"));

            fmt = xls.GetCellVisibleFormatDef(16, 14);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(16, 14, xls.AddFormat(fmt));
            xls.SetCellValue(16, 14, 4);

            fmt = xls.GetCellVisibleFormatDef(16, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(16, 15, xls.AddFormat(fmt));
            xls.SetCellValue(16, 15, " Total Costs = Out of Pocket Costs + Amortized Establishment Costs + Management Costs"
            + " + Opportunity Costs");

            fmt = xls.GetCellVisibleFormatDef(16, 16);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(16, 16, xls.AddFormat(fmt));
            xls.SetCellValue(16, 16, new TFormula("=P8/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(16, 17);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(16, 17, xls.AddFormat(fmt));
            xls.SetCellValue(16, 17, new TFormula("=(Q8/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(16, 18);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(16, 18, xls.AddFormat(fmt));
            xls.SetCellValue(16, 18, "The fourth breakeven return is the total cost breakeven return. Only when this breakeven"
            + " return is received can the grower recover all out-of-pocket expenses plus opportunity"
            + " costs.");
            xls.SetCellValue(16, 20, new TFormula("=P16*Conversiones!$D$24"));

            fmt = xls.GetCellVisibleFormatDef(17, 14);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(17, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 15);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(17, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 16);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(17, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 17);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(17, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 18);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(17, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 14);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(18, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 15);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(18, 15, xls.AddFormat(fmt));
            xls.SetCellValue(18, 15, "Precio actual en dolares por libra:");

            fmt = xls.GetCellVisibleFormatDef(18, 16);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 16, xls.AddFormat(fmt));
            xls.SetCellValue(18, 16, new TFormula("='Outcome_L Adjustment'!$C$16"));

            fmt = xls.GetCellVisibleFormatDef(18, 17);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(18, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 18);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(18, 18, xls.AddFormat(fmt));

            //Comments

            TRTFRun[] Runs;
            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
            TFlxFont fnt;
            fnt = xls.GetDefaultFont;
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 15;
            fnt = xls.GetDefaultFont;
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetComment(15, 2, new TRichString("Juan Hernandez:\nExogenous at this time. From previous studies", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(15, 2, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 15, 0, 3, 203, 16, 13, 5, 981);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nExogenous at this time. From previous studies", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(15, 2, CommentProps);

            //Cell selection and scroll position.
            xls.SelectCell(2, 7, false);

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
