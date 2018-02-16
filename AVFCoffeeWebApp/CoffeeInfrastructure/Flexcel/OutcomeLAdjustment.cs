using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class OutcomeLAdjustment
    {
        public void Outcome_L_Adjustment(ExcelFile xls)
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

            xls.ActiveSheet = 6;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Outcome_L Adjustment";

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
            xls.PrintXResolution = 300;
            xls.PrintYResolution = 300;
            xls.PrintOptions = TPrintOptions.Orientation;
            xls.PrintPaperSize = TPaperSize.Letter;

            //Set up rows and columns
            xls.DefaultColWidth = 2261;

            xls.SetColWidth(1, 1, 1109);    //(3.58 + 0.75) * 256

            TFlxFormat ColFmt;
            ColFmt = xls.GetFormat(xls.GetColFormat(1));
            ColFmt.Font.Size20 = 200;
            ColFmt.HAlignment = THFlxAlignment.left;
            xls.SetColFormat(1, 1, xls.AddFormat(ColFmt));

            xls.SetColWidth(2, 2, 6186);    //(23.41 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(2));
            ColFmt.Font.Size20 = 200;
            ColFmt.WrapText = true;
            xls.SetColFormat(2, 2, xls.AddFormat(ColFmt));

            xls.SetColWidth(3, 3, 4053);    //(15.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(3));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(3, 3, xls.AddFormat(ColFmt));

            xls.SetColWidth(4, 4, 4010);    //(14.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(4));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(4, 4, xls.AddFormat(ColFmt));

            xls.SetColWidth(5, 5, 8405);    //(32.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(5));
            ColFmt.Font.Size20 = 200;
            ColFmt.WrapText = true;
            xls.SetColFormat(5, 5, xls.AddFormat(ColFmt));

            xls.SetColWidth(6, 8, 2261);    //(8.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(6));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(6, 8, xls.AddFormat(ColFmt));

            xls.SetColWidth(9, 9, 4778);    //(17.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(9));
            ColFmt.Font.Size20 = 200;
            ColFmt.ParentStyle = xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0);
            ColFmt.Format = "_(\"$\"* #,##0.00_);_(\"$\"* \\(#,##0.00\\);_(\"$\"* \"-\"??_);_(@_)";
            xls.SetColFormat(9, 9, xls.AddFormat(ColFmt));

            xls.SetColWidth(10, 10, 7082);    //(26.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(10));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(10, 10, xls.AddFormat(ColFmt));

            xls.SetColWidth(11, 11, 6528);    //(24.75 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(11));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(11, 11, xls.AddFormat(ColFmt));

            xls.SetColWidth(12, 12, 5802);    //(21.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(12));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(12, 12, xls.AddFormat(ColFmt));

            xls.SetColWidth(13, 15, 2261);    //(8.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(13));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(13, 15, xls.AddFormat(ColFmt));

            xls.SetColWidth(16, 16, 3498);    //(12.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(16));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(16, 16, xls.AddFormat(ColFmt));

            xls.SetColWidth(17, 16384, 2261);    //(8.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(17));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(17, 16384, xls.AddFormat(ColFmt));
            xls.DefaultRowHeight = 280;

            xls.SetRowHeight(1, 660);    //33.00 * 20
            xls.SetRowHeight(2, 1240);    //62.00 * 20

            TFlxFormat RowFmt;
            RowFmt = xls.GetFormat(xls.GetRowFormat(2));
            RowFmt.Font.Size20 = 200;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(2, xls.AddFormat(RowFmt));
            xls.SetRowHeight(3, 520);    //26.00 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(3));
            RowFmt.Font.Size20 = 200;
            RowFmt.VAlignment = TVFlxAlignment.top;
            xls.SetRowFormat(3, xls.AddFormat(RowFmt));
            xls.SetRowHeight(4, 1040);    //52.00 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(4));
            RowFmt.Font.Size20 = 200;
            RowFmt.VAlignment = TVFlxAlignment.top;
            xls.SetRowFormat(4, xls.AddFormat(RowFmt));
            xls.SetRowHeight(5, 520);    //26.00 * 20
            xls.SetRowHeight(6, 1300);    //65.00 * 20
            xls.SetRowHeight(10, 780);    //39.00 * 20
            xls.SetRowHeight(11, 520);    //26.00 * 20
            xls.SetRowHeight(12, 1040);    //52.00 * 20
            xls.SetRowHeight(13, 520);    //26.00 * 20
            xls.SetRowHeight(14, 1300);    //65.00 * 20
            xls.SetRowHeight(16, 520);    //26.00 * 20

            //Merged Cells
            xls.MergeCells(1, 1, 1, 5);
            xls.MergeCells(10, 1, 10, 2);
            xls.MergeCells(1, 11, 2, 11);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 1);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(1, 1, xls.AddFormat(fmt));
            xls.SetCellValue(1, 1, "Table 6. Conventional breakeven return at different levels of enterprise costs assuming"
            + " average cost and productivity  (years 2 to 8)");

            fmt = xls.GetCellVisibleFormatDef(1, 2);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(1, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 3);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(1, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 4);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(1, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 5);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(1, 5, xls.AddFormat(fmt));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 9, xls.AddFormat(fmt));
            xls.SetCellValue(1, 9, "Referencia");

            fmt = xls.GetCellVisibleFormatDef(1, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(1, 10, xls.AddFormat(fmt));
            xls.SetCellValue(1, 10, "Diferencia asumiendo Y dado (1419.6 pounds/ht)");

            fmt = xls.GetCellVisibleFormatDef(1, 11);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(1, 11, xls.AddFormat(fmt));
            xls.SetCellValue(1, 11, "Assumptions reference");

            fmt = xls.GetCellVisibleFormatDef(1, 12);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(1, 12, xls.AddFormat(fmt));
            xls.SetCellValue(1, 12, "Salary");
            xls.SetCellValue(1, 13, 93.1);
            xls.SetCellValue(1, 16, "US");
            xls.SetCellValue(1, 17, new TFormula("=M1/Conversiones!F24"));

            fmt = xls.GetCellVisibleFormatDef(2, 1);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 1, xls.AddFormat(fmt));
            xls.SetCellValue(2, 1, 1);

            fmt = xls.GetCellVisibleFormatDef(2, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(2, 2, xls.AddFormat(fmt));
            xls.SetCellValue(2, 2, 3);

            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));
            xls.SetCellValue(2, 3, "Costo producción cereza (Pesos/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(2, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(2, 4, xls.AddFormat(fmt));
            xls.SetCellValue(2, 4, "Breakeven -  Retorno (Pesos/quintal)");

            fmt = xls.GetCellVisibleFormatDef(2, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(2, 5, xls.AddFormat(fmt));
            xls.SetCellValue(2, 5, "Breakeven Implications");

            fmt = xls.GetCellVisibleFormatDef(2, 9);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(2, 9, xls.AddFormat(fmt));
            xls.SetCellValue(2, 9, "Costo producción cereza (Pesos/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(2, 10);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(2, 10, xls.AddFormat(fmt));
            xls.SetCellValue(2, 10, "DIFERENCIA Costo producción cereza (Pesos/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(2, 11);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(2, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 12);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(2, 12, xls.AddFormat(fmt));
            xls.SetCellValue(2, 12, "How many quintales of coffee do you produce on average in one year per hectare?");

            fmt = xls.GetCellVisibleFormatDef(2, 13);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 13, xls.AddFormat(fmt));
            xls.SetCellValue(2, 13, 14);

            fmt = xls.GetCellVisibleFormatDef(2, 16);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 16, xls.AddFormat(fmt));
            xls.SetCellValue(2, 16, "POUNDS/HT");

            fmt = xls.GetCellVisibleFormatDef(2, 17);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 17, xls.AddFormat(fmt));
            xls.SetCellValue(2, 17, new TFormula("=M2*Conversiones!C14"));

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));
            xls.SetCellValue(3, 1, 1);

            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Font.Size20 = 200;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));
            xls.SetCellValue(3, 2, "Total Variable Costs");

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));
            xls.SetCellValue(3, 3, new TFormula("=(Budget_Presupuesto!K46*Budget_Supuestos!B6)/Budget_Supuestos!B6"));

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));
            xls.SetCellValue(3, 4, new TFormula("=(C3/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));
            xls.SetCellValue(3, 5, "If the return is below this level, coffee is uneconomical to produce.");

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));
            xls.SetCellValue(3, 8, 1);

            fmt = xls.GetCellVisibleFormatDef(3, 9);
            fmt.Font.Size20 = 200;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(3, 9, xls.AddFormat(fmt));
            xls.SetCellValue(3, 9, 19895.212680941);

            fmt = xls.GetCellVisibleFormatDef(3, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(3, 10, xls.AddFormat(fmt));
            xls.SetCellValue(3, 10, new TFormula("=C3-I3"));

            fmt = xls.GetCellVisibleFormatDef(4, 1);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(4, 1, xls.AddFormat(fmt));
            xls.SetCellValue(4, 1, 2);

            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));
            xls.SetCellValue(4, 2, "Total Cash Costs = Total Variable Costs + Membership & Certification Costs + Taxes"
            + " on Land + Miscellaneous Supplies");

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));
            xls.SetCellValue(4, 3, new TFormula("=C3+((Budget_Presupuesto!K58-Budget_Presupuesto!K29)+Budget_Presupuesto!K72+ (Budget_Presupuesto!K73*Budget_Supuestos!B6))/Budget_Supuestos!B6"));

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));
            xls.SetCellValue(4, 4, new TFormula("=(C4/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));
            xls.SetCellValue(4, 5, "The second breakeven return allows the producer to stay in business in the short run.");

            fmt = xls.GetCellVisibleFormatDef(4, 8);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(4, 8, xls.AddFormat(fmt));
            xls.SetCellValue(4, 8, 2);

            fmt = xls.GetCellVisibleFormatDef(4, 9);
            fmt.Font.Size20 = 200;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(4, 9, xls.AddFormat(fmt));
            xls.SetCellValue(4, 9, 20205.4130854457);

            fmt = xls.GetCellVisibleFormatDef(4, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(4, 10, xls.AddFormat(fmt));
            xls.SetCellValue(4, 10, new TFormula("=C4-I4"));

            fmt = xls.GetCellVisibleFormatDef(5, 1);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(5, 1, xls.AddFormat(fmt));
            xls.SetCellValue(5, 1, 3);

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
            xls.SetCellValue(5, 2, "Out Of Pocket Costs = Total Cash Costs + Depreciation Costs");

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));
            xls.SetCellValue(5, 3, new TFormula("=C4+(Budget_Presupuesto!K61*Budget_Supuestos!B6+Budget_Presupuesto!K62+Budget_Presupuesto!K63*Budget_Supuestos!B6)/Budget_Supuestos!B6"));

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));
            xls.SetCellValue(5, 4, new TFormula("=(C5/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));
            xls.SetCellValue(5, 5, "The third breakeven allows the producer to stay in business in the long run.");

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));
            xls.SetCellValue(5, 8, 3);

            fmt = xls.GetCellVisibleFormatDef(5, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(5, 9, xls.AddFormat(fmt));
            xls.SetCellValue(5, 9, 31525.9748975092);

            fmt = xls.GetCellVisibleFormatDef(5, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(5, 10, xls.AddFormat(fmt));
            xls.SetCellValue(5, 10, new TFormula("=C5-I5"));

            fmt = xls.GetCellVisibleFormatDef(6, 1);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(6, 1, xls.AddFormat(fmt));
            xls.SetCellValue(6, 1, 4);

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, " Total Costs = Out of Pocket Costs + Amortized Establishment Costs + Management Costs"
            + " + Opportunity Costs");

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));
            xls.SetCellValue(6, 3, new TFormula("=C5+(Budget_Presupuesto!K67*Budget_Supuestos!B6+ Budget_Presupuesto!K68*Budget_Supuestos!B6+Budget_Presupuesto!K74)/Budget_Supuestos!B6"));

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));
            xls.SetCellValue(6, 4, new TFormula("=(C6/Budget_Supuestos!$L$155)"));

            fmt = xls.GetCellVisibleFormatDef(6, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(6, 5, xls.AddFormat(fmt));
            xls.SetCellValue(6, 5, "The fourth breakeven return is the total cost breakeven return. Only when this breakeven"
            + " return is received can the grower recover all out-of-pocket expenses plus opportunity"
            + " costs.");

            fmt = xls.GetCellVisibleFormatDef(6, 8);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(6, 8, xls.AddFormat(fmt));
            xls.SetCellValue(6, 8, 4);

            fmt = xls.GetCellVisibleFormatDef(6, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(6, 9, xls.AddFormat(fmt));
            xls.SetCellValue(6, 9, 40189.7533185618);

            fmt = xls.GetCellVisibleFormatDef(6, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(6, 10, xls.AddFormat(fmt));
            xls.SetCellValue(6, 10, new TFormula("=C6-I6"));

            fmt = xls.GetCellVisibleFormatDef(7, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(7, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));
            xls.SetCellValue(8, 2, "Precio actual en pesos Quintal:");

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));
            xls.SetCellValue(8, 3, new TFormula("=Budget_Supuestos!B48"));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));
            xls.SetCellValue(8, 4, new TFormula("=C8/Conversiones!C11"));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));
            xls.SetCellValue(8, 5, new TFormula("=D8/Conversiones!E24"));

            fmt = xls.GetCellVisibleFormatDef(8, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 1);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(10, 1, xls.AddFormat(fmt));
            xls.SetCellValue(10, 1, "Cost definition");

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));
            xls.SetCellValue(10, 3, "Costo producción pergamino (US/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));
            xls.SetCellValue(10, 4, "Breakeven Retorno (us/pound pregamino)");

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
            xls.SetCellValue(10, 5, "Breakeven Implications");

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));
            xls.SetCellValue(10, 9, "Costo producción pergamino (US/Hectarea)");

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));
            xls.SetCellValue(11, 1, 1);

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Font.Size20 = 200;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));
            xls.SetCellValue(11, 2, "Total Variable Costs");

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));
            xls.SetCellValue(11, 3, new TFormula("=C3/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));
            xls.SetCellValue(11, 4, new TFormula("=(D3/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));
            xls.SetCellValue(11, 5, "If the return is below this level, coffee is uneconomical to produce.");

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.Font.Size20 = 200;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));
            xls.SetCellValue(11, 8, 1);

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));
            xls.SetCellValue(11, 9, new TFormula("=I3/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(11, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(11, 10, xls.AddFormat(fmt));
            xls.SetCellValue(11, 10, new TFormula("=C11-I11"));

            fmt = xls.GetCellVisibleFormatDef(12, 1);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(12, 1, xls.AddFormat(fmt));
            xls.SetCellValue(12, 1, 2);

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));
            xls.SetCellValue(12, 2, "Total Cash Costs = Total Variable Costs + Membership & Certification Costs + Taxes"
            + " on Land + Miscellaneous Supplies");

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));
            xls.SetCellValue(12, 3, new TFormula("=C4/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));
            xls.SetCellValue(12, 4, new TFormula("=(D4/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));
            xls.SetCellValue(12, 5, "The second breakeven return allows the producer to stay in business in the short run.");

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));
            xls.SetCellValue(12, 8, 2);

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));
            xls.SetCellValue(12, 9, new TFormula("=I4/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(12, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(12, 10, xls.AddFormat(fmt));
            xls.SetCellValue(12, 10, new TFormula("=C12-I12"));

            fmt = xls.GetCellVisibleFormatDef(13, 1);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(13, 1, xls.AddFormat(fmt));
            xls.SetCellValue(13, 1, 3);

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));
            xls.SetCellValue(13, 2, "Out Of Pocket Costs = Total Cash Costs + Depreciation Costs");

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
            xls.SetCellValue(13, 3, new TFormula("=C5/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, new TFormula("=(D5/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));
            xls.SetCellValue(13, 5, "The third breakeven allows the producer to stay in business in the long run.");

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));
            xls.SetCellValue(13, 8, 3);

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));
            xls.SetCellValue(13, 9, new TFormula("=I5/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(13, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(13, 10, xls.AddFormat(fmt));
            xls.SetCellValue(13, 10, new TFormula("=C13-I13"));

            fmt = xls.GetCellVisibleFormatDef(14, 1);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(14, 1, xls.AddFormat(fmt));
            xls.SetCellValue(14, 1, 4);

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            fmt.Lotus123Prefix = true;
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));
            xls.SetCellValue(14, 2, " Total Costs = Out of Pocket Costs + Amortized Establishment Costs + Management Costs"
            + " + Opportunity Costs");

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));
            xls.SetCellValue(14, 3, new TFormula("=C6/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, new TFormula("=(D6/Conversiones!$C$14)/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));
            xls.SetCellValue(14, 5, "The fourth breakeven return is the total cost breakeven return. Only when this breakeven"
            + " return is received can the grower recover all out-of-pocket expenses plus opportunity"
            + " costs.");

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Font.Size20 = 200;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));
            xls.SetCellValue(14, 8, 4);

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.Font.Size20 = 200;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));
            xls.SetCellValue(14, 9, new TFormula("=I6/Conversiones!$F$24"));

            fmt = xls.GetCellVisibleFormatDef(14, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.VAlignment = TVFlxAlignment.top;
            xls.SetCellFormat(14, 10, xls.AddFormat(fmt));
            xls.SetCellValue(14, 10, new TFormula("=C14-I14"));

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Size20 = 200;
            fmt.WrapText = true;
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));
            xls.SetCellValue(16, 2, "Precio actual en dolares por libra:");

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));
            xls.SetCellValue(16, 3, new TFormula("=(C8/Conversiones!C14)/Conversiones!F24"));

            //Cell selection and scroll position.
            xls.SelectCell(3, 3, false);

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
