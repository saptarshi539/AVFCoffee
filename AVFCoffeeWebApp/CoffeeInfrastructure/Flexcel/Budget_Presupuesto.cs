using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;
namespace CoffeeInfrastructure.Flexcel
{
    public class Budget_Presupuesto
    {

        public void BudgetPresupuesto(ExcelFile xls)
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

            xls.ActiveSheet = 24;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Budget_Presupuesto";
            xls.SheetZoom = 82;
            xls.SheetView = new TSheetView(TSheetViewType.Normal, true, true, 82, 82, 0);

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

            //You can set the margins in 2 ways, the one commented here or the one below:
            //    TXlsMargins PrintMargins = xls.GetPrintMargins();
            //    PrintMargins.Left = 0.75;
            //    PrintMargins.Top = 1;
            //    PrintMargins.Right = 0.75;
            //    PrintMargins.Bottom = 1;
            //    PrintMargins.Header = 0.5;
            //    PrintMargins.Footer = 0.5;
            //    xls.SetPrintMargins(PrintMargins);
            xls.SetPrintMargins(new TXlsMargins(0.75, 1, 0.75, 1, 0.5, 0.5));
            xls.PrintToFit = true;
            xls.PrintScale = 67;
            xls.PrintXResolution = 600;
            xls.PrintYResolution = 600;
            xls.PrintOptions = TPrintOptions.None;
            xls.PrintPaperSize = TPaperSize.Letter;

            //Set up rows and columns
            xls.DefaultColWidth = 2773;

            xls.SetColWidth(1, 1, 13056);    //(50.25 + 0.75) * 256

            xls.SetColWidth(2, 2, 3754);    //(13.91 + 0.75) * 256

            TFlxFormat ColFmt;
            ColFmt = xls.GetFormat(xls.GetColFormat(2));
            ColFmt.Format = "0";
            xls.SetColFormat(2, 2, xls.AddFormat(ColFmt));

            xls.SetColWidth(3, 3, 4138);    //(15.41 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(3));
            ColFmt.Format = "0";
            xls.SetColFormat(3, 3, xls.AddFormat(ColFmt));

            xls.SetColWidth(4, 4, 3370);    //(12.41 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(4));
            ColFmt.Format = "0";
            xls.SetColFormat(4, 4, xls.AddFormat(ColFmt));

            xls.SetColWidth(5, 10, 3200);    //(11.75 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(5));
            ColFmt.Format = "0";
            xls.SetColFormat(5, 10, xls.AddFormat(ColFmt));

            xls.SetColWidth(11, 11, 2944);    //(10.75 + 0.75) * 256

            xls.SetColWidth(12, 12, 3200);    //(11.75 + 0.75) * 256

            xls.SetColWidth(13, 16384, 2773);    //(10.08 + 0.75) * 256

            xls.SetRowHeight(1, 600);    //30.00 * 20
            xls.SetRowHeight(4, 740);    //37.00 * 20
            xls.SetRowHidden(6, true);
            xls.SetRowHidden(7, true);
            xls.SetRowHidden(8, true);
            xls.SetRowHidden(9, true);
            xls.SetRowHidden(10, true);
            xls.SetRowHidden(11, true);
            xls.SetRowHidden(12, true);
            xls.SetRowHidden(13, true);
            xls.SetRowHidden(14, true);
            xls.SetRowHidden(15, true);
            xls.SetRowHidden(16, true);
            xls.SetRowHidden(17, true);
            xls.SetRowHidden(18, true);
            xls.SetRowHidden(19, true);

            TFlxFormat RowFmt;
            RowFmt = xls.GetFormat(xls.GetRowFormat(20));
            RowFmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetRowFormat(20, xls.AddFormat(RowFmt));
            xls.SetRowHeight(23, 600);    //30.00 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(29));
            RowFmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            RowFmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetRowFormat(29, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(31));
            RowFmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            RowFmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetRowFormat(31, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(46));
            RowFmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetRowFormat(46, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(48));
            RowFmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetRowFormat(48, xls.AddFormat(RowFmt));

            //Merged Cells
            xls.MergeCells(1, 12, 1, 20);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(1, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 2);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 2, xls.AddFormat(fmt));
            xls.SetCellValue(1, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(1, 3);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 3, xls.AddFormat(fmt));
            xls.SetCellValue(1, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(1, 4);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 4, xls.AddFormat(fmt));
            xls.SetCellValue(1, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(1, 5);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 5, xls.AddFormat(fmt));
            xls.SetCellValue(1, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(1, 6);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 6, xls.AddFormat(fmt));
            xls.SetCellValue(1, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(1, 7);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 7, xls.AddFormat(fmt));
            xls.SetCellValue(1, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(1, 8);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 8, xls.AddFormat(fmt));
            xls.SetCellValue(1, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(1, 9);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 9, xls.AddFormat(fmt));
            xls.SetCellValue(1, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(1, 10);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 10, xls.AddFormat(fmt));
            xls.SetCellValue(1, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(1, 11);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(1, 11, xls.AddFormat(fmt));
            xls.SetCellValue(1, 11, "Promedio Año 2 - 8");

            fmt = xls.GetCellVisibleFormatDef(1, 12);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 12, xls.AddFormat(fmt));
            xls.SetCellValue(1, 12, "NOTA: El presupuesto se hará por hectarea, tomando kilos como unidad de medida para"
            + " los calculos");

            fmt = xls.GetCellVisibleFormatDef(1, 13);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 14);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 15);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 16);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 17);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 18);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 19);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 20);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(2, 1, xls.AddFormat(fmt));
            xls.SetCellValue(2, 1, "Ingresos");

            fmt = xls.GetCellVisibleFormatDef(2, 2);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 4);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 5);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 6);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 7);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 8);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 9);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 10);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(2, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(2, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Format = "#,##0";
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));
            xls.SetCellValue(3, 1, new TFormula("=Budget_Supuestos!J171"));

            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 6);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 7);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 9);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 10);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(3, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.WrapText = true;
            xls.SetCellFormat(4, 1, xls.AddFormat(fmt));

            TRTFRun[] Runs;
            Runs = new TRTFRun[6];
            Runs[0].FirstChar = 36;
            TFlxFont fnt;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 44;
            fnt = xls.GetDefaultFont;
            Runs[1].FontIndex = xls.AddFont(fnt);
            Runs[2].FirstChar = 88;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            Runs[2].FontIndex = xls.AddFont(fnt);
            Runs[3].FirstChar = 103;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            Runs[3].FontIndex = xls.AddFont(fnt);
            Runs[4].FirstChar = 112;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            Runs[4].FontIndex = xls.AddFont(fnt);
            Runs[5].FirstChar = 119;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            Runs[5].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(4, 1, new TRichString("     Baseline (Precio pergamino seco * Nº Kg)                                    "
            + "       ESTO ES Precio QUINTALES * No. QUINTALES", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(4, 1, "&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;Baseline (Precio pergamino seco<font color = 'blue'>&nbsp;*"
            //    + " N&ordm; Kg</font>) &nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nb"
            //    + "sp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;&nbsp;<font"
            //    + " color = 'red'>ESTO ES Precio&nbsp;</font><font color = 'black'><b>QUINTALES</b></font><font"
            //    + " color = 'red'>&nbsp;* No.&nbsp;</font><font color = 'black'><b>QUINTALES</b></font>")


            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));
            xls.SetCellValue(4, 4, new TFormula("=Budget_Supuestos!$B$49*Budget_Supuestos!L147"));

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));
            xls.SetCellValue(4, 5, new TFormula("=Budget_Supuestos!$B$49*Budget_Supuestos!L148"));

            fmt = xls.GetCellVisibleFormatDef(4, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 6, xls.AddFormat(fmt));
            xls.SetCellValue(4, 6, new TFormula("=Budget_Supuestos!$B$49*Budget_Supuestos!L149"));

            fmt = xls.GetCellVisibleFormatDef(4, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 7, xls.AddFormat(fmt));
            xls.SetCellValue(4, 7, new TFormula("=Budget_Supuestos!$B$49*Budget_Supuestos!L150"));

            fmt = xls.GetCellVisibleFormatDef(4, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 8, xls.AddFormat(fmt));
            xls.SetCellValue(4, 8, new TFormula("=Budget_Supuestos!$B$49*Budget_Supuestos!L151"));

            fmt = xls.GetCellVisibleFormatDef(4, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 9, xls.AddFormat(fmt));
            xls.SetCellValue(4, 9, new TFormula("=Budget_Supuestos!$B$49*Budget_Supuestos!L152"));

            fmt = xls.GetCellVisibleFormatDef(4, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 10, xls.AddFormat(fmt));
            xls.SetCellValue(4, 10, new TFormula("=Budget_Supuestos!$B$49*Budget_Supuestos!L153"));

            fmt = xls.GetCellVisibleFormatDef(4, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 11, xls.AddFormat(fmt));
            xls.SetCellValue(4, 11, new TFormula("=AVERAGE(D4:J4)"));

            fmt = xls.GetCellVisibleFormatDef(5, 1);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 1, xls.AddFormat(fmt));
            xls.SetCellValue(5, 1, "Venta \"Cerezo\"");

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));
            xls.SetCellValue(5, 4, new TFormula("=E5*(D4/E4)"));

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));
            xls.SetCellValue(5, 5, new TFormula("=F5*(E4/F4)"));

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));
            xls.SetCellValue(5, 6, new TFormula("=G5*(F4/G4)"));

            fmt = xls.GetCellVisibleFormatDef(5, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 7, xls.AddFormat(fmt));
            xls.SetCellValue(5, 7, new TFormula("=H5*(G4/H4)"));

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));
            xls.SetCellValue(5, 8, new TFormula("=I5*(H4/I4)"));

            fmt = xls.GetCellVisibleFormatDef(5, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 9, xls.AddFormat(fmt));
            xls.SetCellValue(5, 9, new TFormula("=J5*(I4/J4)"));

            fmt = xls.GetCellVisibleFormatDef(5, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 10, xls.AddFormat(fmt));
            xls.SetCellValue(5, 10, new TFormula("=Budget_Supuestos!$B$52"));

            fmt = xls.GetCellVisibleFormatDef(5, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 11, xls.AddFormat(fmt));
            xls.SetCellValue(5, 11, new TFormula("=AVERAGE(D5:J5)"));

            fmt = xls.GetCellVisibleFormatDef(6, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 1, xls.AddFormat(fmt));
            xls.SetCellValue(7, 1, new TFormula("=Budget_Supuestos!C144"));

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 1, xls.AddFormat(fmt));
            xls.SetCellValue(8, 1, "     Baseline (pergamino)");

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));
            xls.SetCellValue(8, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));
            xls.SetCellValue(8, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));
            xls.SetCellValue(8, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));
            xls.SetCellValue(8, 7, 0);

            fmt = xls.GetCellVisibleFormatDef(8, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 8, xls.AddFormat(fmt));
            xls.SetCellValue(8, 8, 0);

            fmt = xls.GetCellVisibleFormatDef(8, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 9, xls.AddFormat(fmt));
            xls.SetCellValue(8, 9, 0);

            fmt = xls.GetCellVisibleFormatDef(8, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 10, xls.AddFormat(fmt));
            xls.SetCellValue(8, 10, 0);

            fmt = xls.GetCellVisibleFormatDef(8, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 11, xls.AddFormat(fmt));
            xls.SetCellValue(8, 11, new TFormula("=AVERAGE(D8:J8)"));

            fmt = xls.GetCellVisibleFormatDef(9, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 1, xls.AddFormat(fmt));
            xls.SetCellValue(9, 1, "     FT Premium (oro)");

            fmt = xls.GetCellVisibleFormatDef(9, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));
            xls.SetCellValue(9, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));
            xls.SetCellValue(9, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));
            xls.SetCellValue(9, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));
            xls.SetCellValue(9, 7, 0);

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));
            xls.SetCellValue(9, 8, 0);

            fmt = xls.GetCellVisibleFormatDef(9, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 9, xls.AddFormat(fmt));
            xls.SetCellValue(9, 9, 0);

            fmt = xls.GetCellVisibleFormatDef(9, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 10, xls.AddFormat(fmt));
            xls.SetCellValue(9, 10, 0);

            fmt = xls.GetCellVisibleFormatDef(9, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 11, xls.AddFormat(fmt));
            xls.SetCellValue(9, 11, new TFormula("=AVERAGE(D9:J9)"));

            fmt = xls.GetCellVisibleFormatDef(10, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 1, xls.AddFormat(fmt));
            xls.SetCellValue(10, 1, "     Organic Premium (oro)");

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));
            xls.SetCellValue(10, 4, new TFormula("=IF(Budget_Supuestos!$B$58=1,Budget_Supuestos!$B$64*(Budget_Supuestos!$B$207*(Budget_Supuestos!I147*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
            xls.SetCellValue(10, 5, new TFormula("=IF(Budget_Supuestos!$B$58=1,Budget_Supuestos!$B$64*(Budget_Supuestos!$B$207*(Budget_Supuestos!I148*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));
            xls.SetCellValue(10, 6, new TFormula("=IF(Budget_Supuestos!$B$58=1,Budget_Supuestos!$B$64*(Budget_Supuestos!$B$207*(Budget_Supuestos!I149*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));
            xls.SetCellValue(10, 7, new TFormula("=IF(Budget_Supuestos!$B$58=1,Budget_Supuestos!$B$64*(Budget_Supuestos!$B$207*(Budget_Supuestos!I150*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));
            xls.SetCellValue(10, 8, new TFormula("=IF(Budget_Supuestos!$B$58=1,Budget_Supuestos!$B$64*(Budget_Supuestos!$B$207*(Budget_Supuestos!I151*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));
            xls.SetCellValue(10, 9, new TFormula("=IF(Budget_Supuestos!$B$58=1,Budget_Supuestos!$B$64*(Budget_Supuestos!$B$207*(Budget_Supuestos!I152*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(10, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 10, xls.AddFormat(fmt));
            xls.SetCellValue(10, 10, new TFormula("=IF(Budget_Supuestos!$B$58=1,Budget_Supuestos!$B$64*(Budget_Supuestos!$B$207*(Budget_Supuestos!I153*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(10, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 11, xls.AddFormat(fmt));
            xls.SetCellValue(10, 11, new TFormula("=AVERAGE(D10:J10)"));

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));
            xls.SetCellValue(11, 1, "     Prima \"Cooperativa\" (x volumen café oro)");

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));
            xls.SetCellValue(11, 4, new TFormula("=IF(Budget_Supuestos!$B$59=1,Budget_Supuestos!$B$65*(Budget_Supuestos!$B$207*(Budget_Supuestos!I147*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));
            xls.SetCellValue(11, 5, new TFormula("=IF(Budget_Supuestos!$B$59=1,Budget_Supuestos!$B$65*(Budget_Supuestos!$B$207*(Budget_Supuestos!I148*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));
            xls.SetCellValue(11, 6, new TFormula("=IF(Budget_Supuestos!$B$59=1,Budget_Supuestos!$B$65*(Budget_Supuestos!$B$207*(Budget_Supuestos!I149*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));
            xls.SetCellValue(11, 7, new TFormula("=IF(Budget_Supuestos!$B$59=1,Budget_Supuestos!$B$65*(Budget_Supuestos!$B$207*(Budget_Supuestos!I150*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));
            xls.SetCellValue(11, 8, new TFormula("=IF(Budget_Supuestos!$B$59=1,Budget_Supuestos!$B$65*(Budget_Supuestos!$B$207*(Budget_Supuestos!I151*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));
            xls.SetCellValue(11, 9, new TFormula("=IF(Budget_Supuestos!$B$59=1,Budget_Supuestos!$B$65*(Budget_Supuestos!$B$207*(Budget_Supuestos!I152*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(11, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 10, xls.AddFormat(fmt));
            xls.SetCellValue(11, 10, new TFormula("=IF(Budget_Supuestos!$B$59=1,Budget_Supuestos!$B$65*(Budget_Supuestos!$B$207*(Budget_Supuestos!I153*Conversiones!$C$11))*Budget_Supuestos!$B$67,0)"));

            fmt = xls.GetCellVisibleFormatDef(11, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 11, xls.AddFormat(fmt));
            xls.SetCellValue(11, 11, new TFormula("=AVERAGE(D11:J11)"));

            fmt = xls.GetCellVisibleFormatDef(12, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 1);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 1, xls.AddFormat(fmt));
            xls.SetCellValue(13, 1, "Otra variedad 1:");

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 1, xls.AddFormat(fmt));
            xls.SetCellValue(14, 1, "     Baseline");

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 1, xls.AddFormat(fmt));
            xls.SetCellValue(15, 1, "     Flo Premium");

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 1, xls.AddFormat(fmt));
            xls.SetCellValue(16, 1, "     Organic Premium");

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 1, xls.AddFormat(fmt));
            xls.SetCellValue(17, 1, "     Prima \"Cooperativa\" (x volumen café oro)");

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 1, xls.AddFormat(fmt));
            xls.SetCellValue(20, 1, "Sub Total Returns ($/hectare)");

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));
            xls.SetCellValue(20, 4, new TFormula("=SUM(D4:D18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));
            xls.SetCellValue(20, 5, new TFormula("=SUM(E4:E18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));
            xls.SetCellValue(20, 6, new TFormula("=SUM(F4:F18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));
            xls.SetCellValue(20, 7, new TFormula("=SUM(G4:G18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 8, xls.AddFormat(fmt));
            xls.SetCellValue(20, 8, new TFormula("=SUM(H4:H18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 9, xls.AddFormat(fmt));
            xls.SetCellValue(20, 9, new TFormula("=SUM(I4:I18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 10, xls.AddFormat(fmt));
            xls.SetCellValue(20, 10, new TFormula("=SUM(J4:J18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 11, xls.AddFormat(fmt));
            xls.SetCellValue(20, 11, new TFormula("=SUM(K4:K18)"));

            fmt = xls.GetCellVisibleFormatDef(20, 12);
            fmt.Format = "0";
            xls.SetCellFormat(20, 12, xls.AddFormat(fmt));
            xls.SetCellValue(20, 12, new TFormula("=AVERAGE(D20:J20)"));

            fmt = xls.GetCellVisibleFormatDef(21, 1);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 1, xls.AddFormat(fmt));
            xls.SetCellValue(22, 1, "Otros ingresos indirectos");

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            fmt.WrapText = true;
            xls.SetCellFormat(23, 1, xls.AddFormat(fmt));
            xls.SetCellValue(23, 1, "Transferencias de la cooperativa en dinero o bienes (fertilizantes, abonos)");

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));
            xls.SetCellValue(23, 2, new TFormula("=Budget_Supuestos!D80"));

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));
            xls.SetCellValue(23, 3, new TFormula("=Budget_Supuestos!D81"));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));
            xls.SetCellValue(23, 4, new TFormula("=Budget_Supuestos!D82"));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));
            xls.SetCellValue(23, 5, new TFormula("=Budget_Supuestos!D83"));

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));
            xls.SetCellValue(23, 6, new TFormula("=Budget_Supuestos!D84"));

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));
            xls.SetCellValue(23, 7, new TFormula("=Budget_Supuestos!D85"));

            fmt = xls.GetCellVisibleFormatDef(23, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 8, xls.AddFormat(fmt));
            xls.SetCellValue(23, 8, new TFormula("=Budget_Supuestos!D86"));

            fmt = xls.GetCellVisibleFormatDef(23, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 9, xls.AddFormat(fmt));
            xls.SetCellValue(23, 9, new TFormula("=Budget_Supuestos!D87"));

            fmt = xls.GetCellVisibleFormatDef(23, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 10, xls.AddFormat(fmt));
            xls.SetCellValue(23, 10, new TFormula("=Budget_Supuestos!D88"));

            fmt = xls.GetCellVisibleFormatDef(23, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 11, xls.AddFormat(fmt));
            xls.SetCellValue(23, 11, new TFormula("=AVERAGE(D23:J23)"));

            fmt = xls.GetCellVisibleFormatDef(24, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 1, xls.AddFormat(fmt));
            xls.SetCellValue(24, 1, "Capacitaciones");

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));
            xls.SetCellValue(24, 2, new TFormula("=Budget_Supuestos!D96"));

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));
            xls.SetCellValue(24, 3, new TFormula("=Budget_Supuestos!D97"));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));
            xls.SetCellValue(24, 4, new TFormula("=Budget_Supuestos!D98"));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));
            xls.SetCellValue(24, 5, new TFormula("=Budget_Supuestos!D99"));

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));
            xls.SetCellValue(24, 6, new TFormula("=Budget_Supuestos!D100"));

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));
            xls.SetCellValue(24, 7, new TFormula("=Budget_Supuestos!D101"));

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));
            xls.SetCellValue(24, 8, new TFormula("=Budget_Supuestos!D102"));

            fmt = xls.GetCellVisibleFormatDef(24, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 9, xls.AddFormat(fmt));
            xls.SetCellValue(24, 9, new TFormula("=Budget_Supuestos!D103"));

            fmt = xls.GetCellVisibleFormatDef(24, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 10, xls.AddFormat(fmt));
            xls.SetCellValue(24, 10, new TFormula("=Budget_Supuestos!D104"));

            fmt = xls.GetCellVisibleFormatDef(24, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 11, xls.AddFormat(fmt));
            xls.SetCellValue(24, 11, new TFormula("=AVERAGE(D24:J24)"));

            fmt = xls.GetCellVisibleFormatDef(25, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 1, xls.AddFormat(fmt));
            xls.SetCellValue(25, 1, "Prestamos");

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 1, xls.AddFormat(fmt));
            xls.SetCellValue(26, 1, "Prestamos de la cooperativa");

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));
            xls.SetCellValue(26, 2, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));
            xls.SetCellValue(26, 3, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));
            xls.SetCellValue(26, 4, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));
            xls.SetCellValue(26, 5, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));
            xls.SetCellValue(26, 6, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));
            xls.SetCellValue(26, 7, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));
            xls.SetCellValue(26, 8, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 9, xls.AddFormat(fmt));
            xls.SetCellValue(26, 9, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 10, xls.AddFormat(fmt));
            xls.SetCellValue(26, 10, new TFormula("=Budget_Supuestos!$F$110/9"));

            fmt = xls.GetCellVisibleFormatDef(26, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 11, xls.AddFormat(fmt));
            xls.SetCellValue(26, 11, new TFormula("=AVERAGE(D26:J26)"));

            fmt = xls.GetCellVisibleFormatDef(27, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 1, xls.AddFormat(fmt));
            xls.SetCellValue(27, 1, "Prestamos otros bancos o prestamistas");

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));
            xls.SetCellValue(27, 2, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));
            xls.SetCellValue(27, 3, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));
            xls.SetCellValue(27, 4, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));
            xls.SetCellValue(27, 5, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));
            xls.SetCellValue(27, 6, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));
            xls.SetCellValue(27, 7, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 8, xls.AddFormat(fmt));
            xls.SetCellValue(27, 8, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 9, xls.AddFormat(fmt));
            xls.SetCellValue(27, 9, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 10, xls.AddFormat(fmt));
            xls.SetCellValue(27, 10, new TFormula("=Budget_Supuestos!$F$126/9"));

            fmt = xls.GetCellVisibleFormatDef(27, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 11, xls.AddFormat(fmt));
            xls.SetCellValue(27, 11, new TFormula("=AVERAGE(D27:J27)"));

            fmt = xls.GetCellVisibleFormatDef(28, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 1, xls.AddFormat(fmt));
            xls.SetCellValue(29, 1, "Sub Total Otros Ingresos");

            fmt = xls.GetCellVisibleFormatDef(29, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, new TFormula("=SUM(B23:B27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));
            xls.SetCellValue(29, 3, new TFormula("=SUM(C23:C27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));
            xls.SetCellValue(29, 4, new TFormula("=SUM(D23:D27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));
            xls.SetCellValue(29, 5, new TFormula("=SUM(E23:E27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));
            xls.SetCellValue(29, 6, new TFormula("=SUM(F23:F27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));
            xls.SetCellValue(29, 7, new TFormula("=SUM(G23:G27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 8, xls.AddFormat(fmt));
            xls.SetCellValue(29, 8, new TFormula("=SUM(H23:H27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 9, xls.AddFormat(fmt));
            xls.SetCellValue(29, 9, new TFormula("=SUM(I23:I27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 10, xls.AddFormat(fmt));
            xls.SetCellValue(29, 10, new TFormula("=SUM(J23:J27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 11, xls.AddFormat(fmt));
            xls.SetCellValue(29, 11, new TFormula("=SUM(K23:K27)"));

            fmt = xls.GetCellVisibleFormatDef(29, 12);
            fmt.Format = "0";
            xls.SetCellFormat(29, 12, xls.AddFormat(fmt));
            xls.SetCellValue(29, 12, new TFormula("=AVERAGE(D29:J29)"));

            fmt = xls.GetCellVisibleFormatDef(29, 15);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(29, 15, xls.AddFormat(fmt));
            xls.SetCellValue(29, 15, new TFormula("=(K56+K57)-(K26+K27)"));

            fmt = xls.GetCellVisibleFormatDef(30, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 1, xls.AddFormat(fmt));
            xls.SetCellValue(31, 1, "Total Ingresos");

            fmt = xls.GetCellVisibleFormatDef(31, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));
            xls.SetCellValue(31, 2, new TFormula("=B20+B29"));

            fmt = xls.GetCellVisibleFormatDef(31, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));
            xls.SetCellValue(31, 3, new TFormula("=C20+C29"));

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));
            xls.SetCellValue(31, 4, new TFormula("=D20+D29"));

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));
            xls.SetCellValue(31, 5, new TFormula("=E20+E29"));

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));
            xls.SetCellValue(31, 6, new TFormula("=F20+F29"));

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));
            xls.SetCellValue(31, 7, new TFormula("=G20+G29"));

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));
            xls.SetCellValue(31, 8, new TFormula("=H20+H29"));

            fmt = xls.GetCellVisibleFormatDef(31, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 9, xls.AddFormat(fmt));
            xls.SetCellValue(31, 9, new TFormula("=I20+I29"));

            fmt = xls.GetCellVisibleFormatDef(31, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 10, xls.AddFormat(fmt));
            xls.SetCellValue(31, 10, new TFormula("=J20+J29"));

            fmt = xls.GetCellVisibleFormatDef(31, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 11, xls.AddFormat(fmt));
            xls.SetCellValue(31, 11, new TFormula("=K20+K29"));

            fmt = xls.GetCellVisibleFormatDef(31, 12);
            fmt.Format = "0";
            xls.SetCellFormat(31, 12, xls.AddFormat(fmt));
            xls.SetCellValue(31, 12, new TFormula("=AVERAGE(D31:J31)"));

            fmt = xls.GetCellVisibleFormatDef(32, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(32, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(32, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 5);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Format = "0";
            xls.SetCellFormat(33, 5, xls.AddFormat(fmt));
            xls.SetCellValue(33, 5, new TFormula("=E31/Conversiones!F24"));

            fmt = xls.GetCellVisibleFormatDef(33, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(33, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 17;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0xB0, 0x50);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[0].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(34, 1, new TRichString("Costos variables (pesos MX /hectarea)", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(34, 1, "Costos variables&nbsp;<font color = '#00b050'>(pesos MX /hectarea)</font>")


            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 11, xls.AddFormat(fmt));
            xls.SetCellValue(35, 1, "Germinador");

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));
            xls.SetCellValue(35, 2, new TFormula("=Budget_Establecimiento!$B$15"));

            fmt = xls.GetCellVisibleFormatDef(35, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(36, 1, xls.AddFormat(fmt));
            xls.SetCellValue(36, 1, "Vivero");

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));
            xls.SetCellValue(36, 2, new TFormula("=Budget_Establecimiento!B34"));

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(37, 1, xls.AddFormat(fmt));
            xls.SetCellValue(37, 1, "Preparación Terreno y Siembra");

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));
            xls.SetCellValue(37, 2, new TFormula("=Budget_Establecimiento!B51"));

            fmt = xls.GetCellVisibleFormatDef(37, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(38, 1, xls.AddFormat(fmt));
            xls.SetCellValue(38, 1, "Levante o plantilla");

            fmt = xls.GetCellVisibleFormatDef(38, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));
            xls.SetCellValue(38, 3, new TFormula("=Budget_Establecimiento!B69"));

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(38, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(39, 1, xls.AddFormat(fmt));
            xls.SetCellValue(39, 1, new TFormula("=Budget_Sostenemiento!$A$3"));

            fmt = xls.GetCellVisibleFormatDef(39, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));
            xls.SetCellValue(39, 4, new TFormula("=Budget_Sostenemiento!D27"));

            fmt = xls.GetCellVisibleFormatDef(39, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 5, xls.AddFormat(fmt));
            xls.SetCellValue(39, 5, new TFormula("=Budget_Sostenemiento!E27"));

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));
            xls.SetCellValue(39, 6, new TFormula("=Budget_Sostenemiento!F27"));

            fmt = xls.GetCellVisibleFormatDef(39, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 7, xls.AddFormat(fmt));
            xls.SetCellValue(39, 7, new TFormula("=Budget_Sostenemiento!G27"));

            fmt = xls.GetCellVisibleFormatDef(39, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 8, xls.AddFormat(fmt));
            xls.SetCellValue(39, 8, new TFormula("=Budget_Sostenemiento!H27"));

            fmt = xls.GetCellVisibleFormatDef(39, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 9, xls.AddFormat(fmt));
            xls.SetCellValue(39, 9, new TFormula("=Budget_Sostenemiento!I27"));

            fmt = xls.GetCellVisibleFormatDef(39, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 10, xls.AddFormat(fmt));
            xls.SetCellValue(39, 10, new TFormula("=Budget_Sostenemiento!J27"));

            fmt = xls.GetCellVisibleFormatDef(39, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 11, xls.AddFormat(fmt));
            xls.SetCellValue(39, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D39+Budget_Presupuesto!E39)/2))+(Proportions!$E$6*((Budget_Presupuesto!F39+Budget_Presupuesto!G39+Budget_Presupuesto!H39)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I39+Budget_Presupuesto!J39)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(40, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(40, 1, xls.AddFormat(fmt));
            xls.SetCellValue(40, 1, "Cosecha");

            fmt = xls.GetCellVisibleFormatDef(40, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));
            xls.SetCellValue(40, 4, new TFormula("=Budget_Sostenemiento!D35"));

            fmt = xls.GetCellVisibleFormatDef(40, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 5, xls.AddFormat(fmt));
            xls.SetCellValue(40, 5, new TFormula("=Budget_Sostenemiento!E35"));

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));
            xls.SetCellValue(40, 6, new TFormula("=Budget_Sostenemiento!F35"));

            fmt = xls.GetCellVisibleFormatDef(40, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 7, xls.AddFormat(fmt));
            xls.SetCellValue(40, 7, new TFormula("=Budget_Sostenemiento!G35"));

            fmt = xls.GetCellVisibleFormatDef(40, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 8, xls.AddFormat(fmt));
            xls.SetCellValue(40, 8, new TFormula("=Budget_Sostenemiento!H35"));

            fmt = xls.GetCellVisibleFormatDef(40, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 9, xls.AddFormat(fmt));
            xls.SetCellValue(40, 9, new TFormula("=Budget_Sostenemiento!I35"));

            fmt = xls.GetCellVisibleFormatDef(40, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 10, xls.AddFormat(fmt));
            xls.SetCellValue(40, 10, new TFormula("=Budget_Sostenemiento!J35"));

            fmt = xls.GetCellVisibleFormatDef(40, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 11, xls.AddFormat(fmt));
            xls.SetCellValue(40, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D40+Budget_Presupuesto!E40)/2))+(Proportions!$E$6*((Budget_Presupuesto!F40+Budget_Presupuesto!G40+Budget_Presupuesto!H40)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I40+Budget_Presupuesto!J40)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(41, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(41, 1, xls.AddFormat(fmt));
            xls.SetCellValue(41, 1, "Beneficio");

            fmt = xls.GetCellVisibleFormatDef(41, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));
            xls.SetCellValue(41, 4, new TFormula("=Budget_Sostenemiento!D40"));

            fmt = xls.GetCellVisibleFormatDef(41, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 5, xls.AddFormat(fmt));
            xls.SetCellValue(41, 5, new TFormula("=Budget_Sostenemiento!E40"));

            fmt = xls.GetCellVisibleFormatDef(41, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 6, xls.AddFormat(fmt));
            xls.SetCellValue(41, 6, new TFormula("=Budget_Sostenemiento!F40"));

            fmt = xls.GetCellVisibleFormatDef(41, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 7, xls.AddFormat(fmt));
            xls.SetCellValue(41, 7, new TFormula("=Budget_Sostenemiento!G40"));

            fmt = xls.GetCellVisibleFormatDef(41, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 8, xls.AddFormat(fmt));
            xls.SetCellValue(41, 8, new TFormula("=Budget_Sostenemiento!H40"));

            fmt = xls.GetCellVisibleFormatDef(41, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 9, xls.AddFormat(fmt));
            xls.SetCellValue(41, 9, new TFormula("=Budget_Sostenemiento!I40"));

            fmt = xls.GetCellVisibleFormatDef(41, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 10, xls.AddFormat(fmt));
            xls.SetCellValue(41, 10, new TFormula("=Budget_Sostenemiento!J40"));

            fmt = xls.GetCellVisibleFormatDef(41, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 11, xls.AddFormat(fmt));
            xls.SetCellValue(41, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D41+Budget_Presupuesto!E41)/2))+(Proportions!$E$6*((Budget_Presupuesto!F41+Budget_Presupuesto!G41+Budget_Presupuesto!H41)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I41+Budget_Presupuesto!J41)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(42, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(42, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 1);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(43, 1, xls.AddFormat(fmt));
            xls.SetCellValue(43, 1, "Overhead (5% of VC?)");

            fmt = xls.GetCellVisibleFormatDef(43, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 2, xls.AddFormat(fmt));
            xls.SetCellValue(43, 2, new TFormula("=SUM(B35:B42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));
            xls.SetCellValue(43, 3, new TFormula("=SUM(C35:C42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));
            xls.SetCellValue(43, 4, new TFormula("=SUM(D36:D42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 5, xls.AddFormat(fmt));
            xls.SetCellValue(43, 5, new TFormula("=SUM(E35:E42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 6, xls.AddFormat(fmt));
            xls.SetCellValue(43, 6, new TFormula("=SUM(F35:F42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 7, xls.AddFormat(fmt));
            xls.SetCellValue(43, 7, new TFormula("=SUM(G35:G42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 8, xls.AddFormat(fmt));
            xls.SetCellValue(43, 8, new TFormula("=SUM(H35:H42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 9);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 9, xls.AddFormat(fmt));
            xls.SetCellValue(43, 9, new TFormula("=SUM(I35:I42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 10, xls.AddFormat(fmt));
            xls.SetCellValue(43, 10, new TFormula("=SUM(J35:J42)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(43, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 11, xls.AddFormat(fmt));
            xls.SetCellValue(43, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D43+Budget_Presupuesto!E43)/2))+(Proportions!$E$6*((Budget_Presupuesto!F43+Budget_Presupuesto!G43+Budget_Presupuesto!H43)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I43+Budget_Presupuesto!J43)/2)))"));
            xls.SetCellValue(43, 14, "Estos assumptions son discutibles");

            fmt = xls.GetCellVisibleFormatDef(44, 1);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(44, 1, xls.AddFormat(fmt));
            xls.SetCellValue(44, 1, "Interest (5% of VC?)");

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));
            xls.SetCellValue(44, 2, new TFormula("=SUM(B36:B43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));
            xls.SetCellValue(44, 3, new TFormula("=SUM(C36:C43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));
            xls.SetCellValue(44, 4, new TFormula("=SUM(D36:D43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 5, xls.AddFormat(fmt));
            xls.SetCellValue(44, 5, new TFormula("=SUM(E36:E43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 6, xls.AddFormat(fmt));
            xls.SetCellValue(44, 6, new TFormula("=SUM(F36:F43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 7, xls.AddFormat(fmt));
            xls.SetCellValue(44, 7, new TFormula("=SUM(G36:G43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 8, xls.AddFormat(fmt));
            xls.SetCellValue(44, 8, new TFormula("=SUM(H36:H43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 9);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 9, xls.AddFormat(fmt));
            xls.SetCellValue(44, 9, new TFormula("=SUM(I36:I43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 10, xls.AddFormat(fmt));
            xls.SetCellValue(44, 10, new TFormula("=SUM(J36:J43)*0.05"));

            fmt = xls.GetCellVisibleFormatDef(44, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 11, xls.AddFormat(fmt));
            xls.SetCellValue(44, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D44+Budget_Presupuesto!E44)/2))+(Proportions!$E$6*((Budget_Presupuesto!F44+Budget_Presupuesto!G44+Budget_Presupuesto!H44)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I44+Budget_Presupuesto!J44)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(44, 12);
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 12, xls.AddFormat(fmt));
            xls.SetCellValue(44, 12, new TFormula("=K43+K44"));

            fmt = xls.GetCellVisibleFormatDef(45, 1);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(45, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(46, 1, xls.AddFormat(fmt));
            xls.SetCellValue(46, 1, "Total Costos Variables");

            fmt = xls.GetCellVisibleFormatDef(46, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 2, xls.AddFormat(fmt));
            xls.SetCellValue(46, 2, new TFormula("=SUM(B36:B44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));
            xls.SetCellValue(46, 3, new TFormula("=SUM(C36:C44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 4, xls.AddFormat(fmt));
            xls.SetCellValue(46, 4, new TFormula("=SUM(D36:D44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 5, xls.AddFormat(fmt));
            xls.SetCellValue(46, 5, new TFormula("=SUM(E36:E44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 6, xls.AddFormat(fmt));
            xls.SetCellValue(46, 6, new TFormula("=SUM(F36:F44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 7, xls.AddFormat(fmt));
            xls.SetCellValue(46, 7, new TFormula("=SUM(G36:G44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 8, xls.AddFormat(fmt));
            xls.SetCellValue(46, 8, new TFormula("=SUM(H36:H44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 9, xls.AddFormat(fmt));
            xls.SetCellValue(46, 9, new TFormula("=SUM(I36:I44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 10, xls.AddFormat(fmt));
            xls.SetCellValue(46, 10, new TFormula("=SUM(J36:J44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 11);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 11, xls.AddFormat(fmt));
            xls.SetCellValue(46, 11, new TFormula("=SUM(K36:K44)"));

            fmt = xls.GetCellVisibleFormatDef(46, 12);
            fmt.Format = "0";
            xls.SetCellFormat(46, 12, xls.AddFormat(fmt));
            xls.SetCellValue(46, 12, new TFormula("=AVERAGE(D46:J46)"));

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(48, 1, xls.AddFormat(fmt));
            xls.SetCellValue(48, 1, "Fixed Costs ($/hectare)");

            fmt = xls.GetCellVisibleFormatDef(48, 2);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(49, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 1);
            fmt.Font.Underline = TFlxUnderline.Single;
            xls.SetCellFormat(50, 1, xls.AddFormat(fmt));
            xls.SetCellValue(50, 1, "Costos prestamos, certificaciones y membresias");

            fmt = xls.GetCellVisibleFormatDef(50, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 11, xls.AddFormat(fmt));
            xls.SetCellValue(51, 1, "     Application Fee");

            fmt = xls.GetCellVisibleFormatDef(51, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 2, xls.AddFormat(fmt));
            xls.SetCellValue(51, 2, new TFormula("=Budget_Supuestos!$B$380"));

            fmt = xls.GetCellVisibleFormatDef(51, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 11, xls.AddFormat(fmt));
            xls.SetCellValue(52, 1, "     Membresia annual");

            fmt = xls.GetCellVisibleFormatDef(52, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 3, xls.AddFormat(fmt));
            xls.SetCellValue(52, 3, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 4, xls.AddFormat(fmt));
            xls.SetCellValue(52, 4, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 5, xls.AddFormat(fmt));
            xls.SetCellValue(52, 5, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 6, xls.AddFormat(fmt));
            xls.SetCellValue(52, 6, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 7, xls.AddFormat(fmt));
            xls.SetCellValue(52, 7, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 8, xls.AddFormat(fmt));
            xls.SetCellValue(52, 8, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 9, xls.AddFormat(fmt));
            xls.SetCellValue(52, 9, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 10, xls.AddFormat(fmt));
            xls.SetCellValue(52, 10, new TFormula("=Budget_Supuestos!$C$381"));

            fmt = xls.GetCellVisibleFormatDef(52, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 11, xls.AddFormat(fmt));
            xls.SetCellValue(52, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D52+Budget_Presupuesto!E52)/2))+(Proportions!$E$6*((Budget_Presupuesto!F52+Budget_Presupuesto!G52+Budget_Presupuesto!H52)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I52+Budget_Presupuesto!J52)/2)))"));
            xls.SetCellValue(53, 1, "     Seguro de vida");

            fmt = xls.GetCellVisibleFormatDef(53, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 3, xls.AddFormat(fmt));
            xls.SetCellValue(53, 3, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 4, xls.AddFormat(fmt));
            xls.SetCellValue(53, 4, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 5, xls.AddFormat(fmt));
            xls.SetCellValue(53, 5, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 6, xls.AddFormat(fmt));
            xls.SetCellValue(53, 6, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 7, xls.AddFormat(fmt));
            xls.SetCellValue(53, 7, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 8, xls.AddFormat(fmt));
            xls.SetCellValue(53, 8, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 9, xls.AddFormat(fmt));
            xls.SetCellValue(53, 9, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 10, xls.AddFormat(fmt));
            xls.SetCellValue(53, 10, new TFormula("=Budget_Supuestos!$C$382"));

            fmt = xls.GetCellVisibleFormatDef(53, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 11, xls.AddFormat(fmt));
            xls.SetCellValue(53, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D53+Budget_Presupuesto!E53)/2))+(Proportions!$E$6*((Budget_Presupuesto!F53+Budget_Presupuesto!G53+Budget_Presupuesto!H53)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I53+Budget_Presupuesto!J53)/2)))"));
            xls.SetCellValue(54, 1, "     FLO Certificatoin");

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 3, xls.AddFormat(fmt));
            xls.SetCellValue(54, 3, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 4, xls.AddFormat(fmt));
            xls.SetCellValue(54, 4, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 5, xls.AddFormat(fmt));
            xls.SetCellValue(54, 5, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 6, xls.AddFormat(fmt));
            xls.SetCellValue(54, 6, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 7, xls.AddFormat(fmt));
            xls.SetCellValue(54, 7, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 8, xls.AddFormat(fmt));
            xls.SetCellValue(54, 8, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 9, xls.AddFormat(fmt));
            xls.SetCellValue(54, 9, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 10, xls.AddFormat(fmt));
            xls.SetCellValue(54, 10, new TFormula("=Budget_Supuestos!$C$383"));

            fmt = xls.GetCellVisibleFormatDef(54, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 11, xls.AddFormat(fmt));
            xls.SetCellValue(54, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D54+Budget_Presupuesto!E54)/2))+(Proportions!$E$6*((Budget_Presupuesto!F54+Budget_Presupuesto!G54+Budget_Presupuesto!H54)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I54+Budget_Presupuesto!J54)/2)))"));
            xls.SetCellValue(55, 1, "     Organic Certification");

            fmt = xls.GetCellVisibleFormatDef(55, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Font.Underline = TFlxUnderline.Single;
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 3, xls.AddFormat(fmt));
            xls.SetCellValue(55, 3, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 4, xls.AddFormat(fmt));
            xls.SetCellValue(55, 4, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 5, xls.AddFormat(fmt));
            xls.SetCellValue(55, 5, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 6, xls.AddFormat(fmt));
            xls.SetCellValue(55, 6, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 7, xls.AddFormat(fmt));
            xls.SetCellValue(55, 7, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 8, xls.AddFormat(fmt));
            xls.SetCellValue(55, 8, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 9, xls.AddFormat(fmt));
            xls.SetCellValue(55, 9, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 10, xls.AddFormat(fmt));
            xls.SetCellValue(55, 10, new TFormula("=Budget_Supuestos!$C$384"));

            fmt = xls.GetCellVisibleFormatDef(55, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 11, xls.AddFormat(fmt));
            xls.SetCellValue(55, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D55+Budget_Presupuesto!E55)/2))+(Proportions!$E$6*((Budget_Presupuesto!F55+Budget_Presupuesto!G55+Budget_Presupuesto!H55)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I55+Budget_Presupuesto!J55)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(56, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(56, 1, xls.AddFormat(fmt));
            xls.SetCellValue(56, 1, "     Pago prestamo a la cooperativa");

            fmt = xls.GetCellVisibleFormatDef(56, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 2, xls.AddFormat(fmt));
            xls.SetCellValue(56, 2, new TFormula("=Budget_Supuestos!N119"));

            fmt = xls.GetCellVisibleFormatDef(56, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 3, xls.AddFormat(fmt));
            xls.SetCellValue(56, 3, new TFormula("=Budget_Supuestos!O119"));

            fmt = xls.GetCellVisibleFormatDef(56, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 4, xls.AddFormat(fmt));
            xls.SetCellValue(56, 4, new TFormula("=Budget_Supuestos!P119"));

            fmt = xls.GetCellVisibleFormatDef(56, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 5, xls.AddFormat(fmt));
            xls.SetCellValue(56, 5, new TFormula("=Budget_Supuestos!Q119"));

            fmt = xls.GetCellVisibleFormatDef(56, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 6, xls.AddFormat(fmt));
            xls.SetCellValue(56, 6, new TFormula("=Budget_Supuestos!R119"));

            fmt = xls.GetCellVisibleFormatDef(56, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 7, xls.AddFormat(fmt));
            xls.SetCellValue(56, 7, new TFormula("=Budget_Supuestos!S119"));

            fmt = xls.GetCellVisibleFormatDef(56, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 8, xls.AddFormat(fmt));
            xls.SetCellValue(56, 8, new TFormula("=Budget_Supuestos!T119"));

            fmt = xls.GetCellVisibleFormatDef(56, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 9, xls.AddFormat(fmt));
            xls.SetCellValue(56, 9, new TFormula("=Budget_Supuestos!U119"));

            fmt = xls.GetCellVisibleFormatDef(56, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 10, xls.AddFormat(fmt));
            xls.SetCellValue(56, 10, new TFormula("=Budget_Supuestos!V119"));

            fmt = xls.GetCellVisibleFormatDef(56, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 11, xls.AddFormat(fmt));
            xls.SetCellValue(56, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D56+Budget_Presupuesto!E56)/2))+(Proportions!$E$6*((Budget_Presupuesto!F56+Budget_Presupuesto!G56+Budget_Presupuesto!H56)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I56+Budget_Presupuesto!J56)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(57, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(57, 1, xls.AddFormat(fmt));
            xls.SetCellValue(57, 1, "     Pago prestamo otros prestamistas");

            fmt = xls.GetCellVisibleFormatDef(57, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 2, xls.AddFormat(fmt));
            xls.SetCellValue(57, 2, new TFormula("=Budget_Supuestos!N135"));

            fmt = xls.GetCellVisibleFormatDef(57, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 3, xls.AddFormat(fmt));
            xls.SetCellValue(57, 3, new TFormula("=Budget_Supuestos!O135"));

            fmt = xls.GetCellVisibleFormatDef(57, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 4, xls.AddFormat(fmt));
            xls.SetCellValue(57, 4, new TFormula("=Budget_Supuestos!P135"));

            fmt = xls.GetCellVisibleFormatDef(57, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 5, xls.AddFormat(fmt));
            xls.SetCellValue(57, 5, new TFormula("=Budget_Supuestos!Q135"));

            fmt = xls.GetCellVisibleFormatDef(57, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 6, xls.AddFormat(fmt));
            xls.SetCellValue(57, 6, new TFormula("=Budget_Supuestos!R135"));

            fmt = xls.GetCellVisibleFormatDef(57, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 7, xls.AddFormat(fmt));
            xls.SetCellValue(57, 7, new TFormula("=Budget_Supuestos!S135"));

            fmt = xls.GetCellVisibleFormatDef(57, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 8, xls.AddFormat(fmt));
            xls.SetCellValue(57, 8, new TFormula("=Budget_Supuestos!T135"));

            fmt = xls.GetCellVisibleFormatDef(57, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 9, xls.AddFormat(fmt));
            xls.SetCellValue(57, 9, new TFormula("=Budget_Supuestos!U135"));

            fmt = xls.GetCellVisibleFormatDef(57, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 10, xls.AddFormat(fmt));
            xls.SetCellValue(57, 10, new TFormula("=Budget_Supuestos!V135"));

            fmt = xls.GetCellVisibleFormatDef(57, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 11, xls.AddFormat(fmt));
            xls.SetCellValue(57, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D57+Budget_Presupuesto!E57)/2))+(Proportions!$E$6*((Budget_Presupuesto!F57+Budget_Presupuesto!G57+Budget_Presupuesto!H57)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I57+Budget_Presupuesto!J57)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(58, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(58, 1, xls.AddFormat(fmt));
            xls.SetCellValue(58, 1, "Total costos prestamos, certificaciones y membresias");

            fmt = xls.GetCellVisibleFormatDef(58, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 2, xls.AddFormat(fmt));
            xls.SetCellValue(58, 2, new TFormula("=SUM(B51:B57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 3, xls.AddFormat(fmt));
            xls.SetCellValue(58, 3, new TFormula("=SUM(C51:C57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 4, xls.AddFormat(fmt));
            xls.SetCellValue(58, 4, new TFormula("=SUM(D51:D57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 5, xls.AddFormat(fmt));
            xls.SetCellValue(58, 5, new TFormula("=SUM(E51:E57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 6, xls.AddFormat(fmt));
            xls.SetCellValue(58, 6, new TFormula("=SUM(F51:F57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 7, xls.AddFormat(fmt));
            xls.SetCellValue(58, 7, new TFormula("=SUM(G51:G57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 8, xls.AddFormat(fmt));
            xls.SetCellValue(58, 8, new TFormula("=SUM(H51:H57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 9, xls.AddFormat(fmt));
            xls.SetCellValue(58, 9, new TFormula("=SUM(I51:I57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 10, xls.AddFormat(fmt));
            xls.SetCellValue(58, 10, new TFormula("=SUM(J51:J57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 11);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 11, xls.AddFormat(fmt));
            xls.SetCellValue(58, 11, new TFormula("=SUM(K51:K57)"));

            fmt = xls.GetCellVisibleFormatDef(58, 12);
            fmt.Format = "0";
            xls.SetCellFormat(58, 12, xls.AddFormat(fmt));
            xls.SetCellValue(58, 12, new TFormula("=AVERAGE(D58:J58)"));

            fmt = xls.GetCellVisibleFormatDef(59, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 1);
            fmt.Font.Underline = TFlxUnderline.Single;
            xls.SetCellFormat(60, 1, xls.AddFormat(fmt));
            xls.SetCellValue(60, 1, "Depreciación");

            fmt = xls.GetCellVisibleFormatDef(60, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(61, 1, xls.AddFormat(fmt));
            xls.SetCellValue(61, 1, "Total depreciación herramientas generales");

            fmt = xls.GetCellVisibleFormatDef(61, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 2, xls.AddFormat(fmt));
            xls.SetCellValue(61, 2, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 3, xls.AddFormat(fmt));
            xls.SetCellValue(61, 3, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 4, xls.AddFormat(fmt));
            xls.SetCellValue(61, 4, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 5, xls.AddFormat(fmt));
            xls.SetCellValue(61, 5, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 6, xls.AddFormat(fmt));
            xls.SetCellValue(61, 6, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 7, xls.AddFormat(fmt));
            xls.SetCellValue(61, 7, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 8, xls.AddFormat(fmt));
            xls.SetCellValue(61, 8, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 9, xls.AddFormat(fmt));
            xls.SetCellValue(61, 9, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(61, 10, xls.AddFormat(fmt));
            xls.SetCellValue(61, 10, new TFormula("=Budget_Equipo!$I$20"));

            fmt = xls.GetCellVisibleFormatDef(61, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 11, xls.AddFormat(fmt));
            xls.SetCellValue(61, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D61+Budget_Presupuesto!E61)/2))+(Proportions!$E$6*((Budget_Presupuesto!F61+Budget_Presupuesto!G61+Budget_Presupuesto!H61)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I61+Budget_Presupuesto!J61)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(62, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(62, 1, xls.AddFormat(fmt));
            xls.SetCellValue(62, 1, "Total depreciación equipos para el beneficio");

            fmt = xls.GetCellVisibleFormatDef(62, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 4, xls.AddFormat(fmt));
            xls.SetCellValue(62, 4, new TFormula("=Budget_Equipo!$I$42"));

            fmt = xls.GetCellVisibleFormatDef(62, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 5, xls.AddFormat(fmt));
            xls.SetCellValue(62, 5, new TFormula("=Budget_Equipo!$I$42"));

            fmt = xls.GetCellVisibleFormatDef(62, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 6, xls.AddFormat(fmt));
            xls.SetCellValue(62, 6, new TFormula("=Budget_Equipo!$I$42"));

            fmt = xls.GetCellVisibleFormatDef(62, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 7, xls.AddFormat(fmt));
            xls.SetCellValue(62, 7, new TFormula("=Budget_Equipo!$I$42"));

            fmt = xls.GetCellVisibleFormatDef(62, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 8, xls.AddFormat(fmt));
            xls.SetCellValue(62, 8, new TFormula("=Budget_Equipo!$I$42"));

            fmt = xls.GetCellVisibleFormatDef(62, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 9, xls.AddFormat(fmt));
            xls.SetCellValue(62, 9, new TFormula("=Budget_Equipo!$I$42"));

            fmt = xls.GetCellVisibleFormatDef(62, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 10, xls.AddFormat(fmt));
            xls.SetCellValue(62, 10, new TFormula("=Budget_Equipo!$I$42"));

            fmt = xls.GetCellVisibleFormatDef(62, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 11, xls.AddFormat(fmt));
            xls.SetCellValue(62, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D62+Budget_Presupuesto!E62)/2))+(Proportions!$E$6*((Budget_Presupuesto!F62+Budget_Presupuesto!G62+Budget_Presupuesto!H62)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I62+Budget_Presupuesto!J62)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(63, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(63, 1, xls.AddFormat(fmt));
            xls.SetCellValue(63, 1, "Total depreciación otros equipos  y/o materiales reutilizables");

            fmt = xls.GetCellVisibleFormatDef(63, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 4, xls.AddFormat(fmt));
            xls.SetCellValue(63, 4, new TFormula("=Budget_Equipo!$I$53"));

            fmt = xls.GetCellVisibleFormatDef(63, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 5, xls.AddFormat(fmt));
            xls.SetCellValue(63, 5, new TFormula("=Budget_Equipo!$I$53"));

            fmt = xls.GetCellVisibleFormatDef(63, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 6, xls.AddFormat(fmt));
            xls.SetCellValue(63, 6, new TFormula("=Budget_Equipo!$I$53"));

            fmt = xls.GetCellVisibleFormatDef(63, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 7, xls.AddFormat(fmt));
            xls.SetCellValue(63, 7, new TFormula("=Budget_Equipo!$I$53"));

            fmt = xls.GetCellVisibleFormatDef(63, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 8, xls.AddFormat(fmt));
            xls.SetCellValue(63, 8, new TFormula("=Budget_Equipo!$I$53"));

            fmt = xls.GetCellVisibleFormatDef(63, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 9, xls.AddFormat(fmt));
            xls.SetCellValue(63, 9, new TFormula("=Budget_Equipo!$I$53"));

            fmt = xls.GetCellVisibleFormatDef(63, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 10, xls.AddFormat(fmt));
            xls.SetCellValue(63, 10, new TFormula("=Budget_Equipo!$I$53"));

            fmt = xls.GetCellVisibleFormatDef(63, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 11, xls.AddFormat(fmt));
            xls.SetCellValue(63, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D63+Budget_Presupuesto!E63)/2))+(Proportions!$E$6*((Budget_Presupuesto!F63+Budget_Presupuesto!G63+Budget_Presupuesto!H63)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I63+Budget_Presupuesto!J63)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(64, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(64, 1, xls.AddFormat(fmt));
            xls.SetCellValue(64, 1, "Total costos depreciación");

            fmt = xls.GetCellVisibleFormatDef(64, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 2, xls.AddFormat(fmt));
            xls.SetCellValue(64, 2, new TFormula("=SUM(B61:B63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 3, xls.AddFormat(fmt));
            xls.SetCellValue(64, 3, new TFormula("=SUM(C61:C63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 4, xls.AddFormat(fmt));
            xls.SetCellValue(64, 4, new TFormula("=SUM(D61:D63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 5, xls.AddFormat(fmt));
            xls.SetCellValue(64, 5, new TFormula("=SUM(E61:E63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 6, xls.AddFormat(fmt));
            xls.SetCellValue(64, 6, new TFormula("=SUM(F61:F63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 7, xls.AddFormat(fmt));
            xls.SetCellValue(64, 7, new TFormula("=SUM(G61:G63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 8, xls.AddFormat(fmt));
            xls.SetCellValue(64, 8, new TFormula("=SUM(H61:H63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 9, xls.AddFormat(fmt));
            xls.SetCellValue(64, 9, new TFormula("=SUM(I61:I63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 10, xls.AddFormat(fmt));
            xls.SetCellValue(64, 10, new TFormula("=SUM(J61:J63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 11);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 11, xls.AddFormat(fmt));
            xls.SetCellValue(64, 11, new TFormula("=SUM(K61:K63)"));

            fmt = xls.GetCellVisibleFormatDef(64, 12);
            fmt.Format = "0";
            xls.SetCellFormat(64, 12, xls.AddFormat(fmt));
            xls.SetCellValue(64, 12, new TFormula("=AVERAGE(D64:J64)"));

            fmt = xls.GetCellVisibleFormatDef(65, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 1);
            fmt.Font.Underline = TFlxUnderline.Single;
            xls.SetCellFormat(66, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[1];
            Runs[0].FirstChar = 26;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fnt.Underline = TFlxUnderline.Single;
            Runs[0].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(66, 1, new TRichString("Interest/Opportunity Cost (4%)", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(66, 1, "Interest/Opportunity Cost&nbsp;<font color = 'red'>(4%)</font>")


            fmt = xls.GetCellVisibleFormatDef(66, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 11, xls.AddFormat(fmt));
            xls.SetCellValue(67, 1, "     Tierra");

            fmt = xls.GetCellVisibleFormatDef(67, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 3, xls.AddFormat(fmt));
            xls.SetCellValue(67, 3, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 4, xls.AddFormat(fmt));
            xls.SetCellValue(67, 4, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 5, xls.AddFormat(fmt));
            xls.SetCellValue(67, 5, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 6, xls.AddFormat(fmt));
            xls.SetCellValue(67, 6, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 7, xls.AddFormat(fmt));
            xls.SetCellValue(67, 7, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 8, xls.AddFormat(fmt));
            xls.SetCellValue(67, 8, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 9);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 9, xls.AddFormat(fmt));
            xls.SetCellValue(67, 9, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 10, xls.AddFormat(fmt));
            xls.SetCellValue(67, 10, new TFormula("=Budget_Supuestos!$C$389"));

            fmt = xls.GetCellVisibleFormatDef(67, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 11, xls.AddFormat(fmt));
            xls.SetCellValue(67, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D67+Budget_Presupuesto!E67)/2))+(Proportions!$E$6*((Budget_Presupuesto!F67+Budget_Presupuesto!G67+Budget_Presupuesto!H67)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I67+Budget_Presupuesto!J67)/2)))"));
            xls.SetCellValue(68, 1, "     Equipo");

            fmt = xls.GetCellVisibleFormatDef(68, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 3, xls.AddFormat(fmt));
            xls.SetCellValue(68, 3, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 4, xls.AddFormat(fmt));
            xls.SetCellValue(68, 4, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 5, xls.AddFormat(fmt));
            xls.SetCellValue(68, 5, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 6, xls.AddFormat(fmt));
            xls.SetCellValue(68, 6, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 7, xls.AddFormat(fmt));
            xls.SetCellValue(68, 7, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 8, xls.AddFormat(fmt));
            xls.SetCellValue(68, 8, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 9);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 9, xls.AddFormat(fmt));
            xls.SetCellValue(68, 9, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFB, 0xA9);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 10, xls.AddFormat(fmt));
            xls.SetCellValue(68, 10, new TFormula("=Budget_Equipo!$H$57"));

            fmt = xls.GetCellVisibleFormatDef(68, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 11, xls.AddFormat(fmt));
            xls.SetCellValue(68, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D68+Budget_Presupuesto!E68)/2))+(Proportions!$E$6*((Budget_Presupuesto!F68+Budget_Presupuesto!G68+Budget_Presupuesto!H68)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I68+Budget_Presupuesto!J68)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(69, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(69, 1, xls.AddFormat(fmt));
            xls.SetCellValue(69, 1, "Total costo oportunidad");

            fmt = xls.GetCellVisibleFormatDef(69, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 2, xls.AddFormat(fmt));
            xls.SetCellValue(69, 2, new TFormula("=SUM(B67:B68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 3, xls.AddFormat(fmt));
            xls.SetCellValue(69, 3, new TFormula("=SUM(C67:C68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 4, xls.AddFormat(fmt));
            xls.SetCellValue(69, 4, new TFormula("=SUM(D67:D68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 5, xls.AddFormat(fmt));
            xls.SetCellValue(69, 5, new TFormula("=SUM(E67:E68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 6, xls.AddFormat(fmt));
            xls.SetCellValue(69, 6, new TFormula("=SUM(F67:F68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 7, xls.AddFormat(fmt));
            xls.SetCellValue(69, 7, new TFormula("=SUM(G67:G68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 8, xls.AddFormat(fmt));
            xls.SetCellValue(69, 8, new TFormula("=SUM(H67:H68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 9, xls.AddFormat(fmt));
            xls.SetCellValue(69, 9, new TFormula("=SUM(I67:I68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 10, xls.AddFormat(fmt));
            xls.SetCellValue(69, 10, new TFormula("=SUM(J67:J68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 11);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 11, xls.AddFormat(fmt));
            xls.SetCellValue(69, 11, new TFormula("=SUM(K67:K68)"));

            fmt = xls.GetCellVisibleFormatDef(69, 12);
            fmt.Format = "0";
            xls.SetCellFormat(69, 12, xls.AddFormat(fmt));
            xls.SetCellValue(69, 12, new TFormula("=AVERAGE(D69:J69)"));

            fmt = xls.GetCellVisibleFormatDef(70, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 1);
            fmt.Font.Underline = TFlxUnderline.Single;
            xls.SetCellFormat(71, 1, xls.AddFormat(fmt));
            xls.SetCellValue(71, 1, "Other Fixed Costs");

            fmt = xls.GetCellVisibleFormatDef(71, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 11, xls.AddFormat(fmt));
            xls.SetCellValue(72, 1, "     Miscellaneous Supplies (10% of variable costs)");

            fmt = xls.GetCellVisibleFormatDef(72, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 3, xls.AddFormat(fmt));
            xls.SetCellValue(72, 3, new TFormula("=0.1*C46"));

            fmt = xls.GetCellVisibleFormatDef(72, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 4, xls.AddFormat(fmt));
            xls.SetCellValue(72, 4, new TFormula("=0.1*D46"));

            fmt = xls.GetCellVisibleFormatDef(72, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 5, xls.AddFormat(fmt));
            xls.SetCellValue(72, 5, new TFormula("=0.1*E46"));

            fmt = xls.GetCellVisibleFormatDef(72, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 6, xls.AddFormat(fmt));
            xls.SetCellValue(72, 6, new TFormula("=0.1*F46"));

            fmt = xls.GetCellVisibleFormatDef(72, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 7, xls.AddFormat(fmt));
            xls.SetCellValue(72, 7, new TFormula("=0.1*G46"));

            fmt = xls.GetCellVisibleFormatDef(72, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 8, xls.AddFormat(fmt));
            xls.SetCellValue(72, 8, new TFormula("=0.1*H46"));

            fmt = xls.GetCellVisibleFormatDef(72, 9);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 9, xls.AddFormat(fmt));
            xls.SetCellValue(72, 9, new TFormula("=0.1*I46"));

            fmt = xls.GetCellVisibleFormatDef(72, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 10, xls.AddFormat(fmt));
            xls.SetCellValue(72, 10, new TFormula("=0.1*J46"));

            fmt = xls.GetCellVisibleFormatDef(72, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 11, xls.AddFormat(fmt));
            xls.SetCellValue(72, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D72+Budget_Presupuesto!E72)/2))+(Proportions!$E$6*((Budget_Presupuesto!F72+Budget_Presupuesto!G72+Budget_Presupuesto!H72)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I72+Budget_Presupuesto!J72)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(72, 12);
            fmt.Format = "0";
            xls.SetCellFormat(72, 12, xls.AddFormat(fmt));
            xls.SetCellValue(73, 1, "     Impuestos de la tierra");

            fmt = xls.GetCellVisibleFormatDef(73, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 3, xls.AddFormat(fmt));
            xls.SetCellValue(73, 3, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 4, xls.AddFormat(fmt));
            xls.SetCellValue(73, 4, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 5, xls.AddFormat(fmt));
            xls.SetCellValue(73, 5, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 6, xls.AddFormat(fmt));
            xls.SetCellValue(73, 6, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 7, xls.AddFormat(fmt));
            xls.SetCellValue(73, 7, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 8, xls.AddFormat(fmt));
            xls.SetCellValue(73, 8, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 9, xls.AddFormat(fmt));
            xls.SetCellValue(73, 9, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 10, xls.AddFormat(fmt));
            xls.SetCellValue(73, 10, new TFormula("=Budget_Supuestos!$B$393"));

            fmt = xls.GetCellVisibleFormatDef(73, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(73, 11, xls.AddFormat(fmt));
            xls.SetCellValue(73, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D73+Budget_Presupuesto!E73)/2))+(Proportions!$E$6*((Budget_Presupuesto!F73+Budget_Presupuesto!G73+Budget_Presupuesto!H73)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I73+Budget_Presupuesto!J73)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(74, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            xls.SetCellFormat(74, 1, xls.AddFormat(fmt));
            xls.SetCellValue(74, 1, "     Management Cost ");

            fmt = xls.GetCellVisibleFormatDef(74, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 2, xls.AddFormat(fmt));
            xls.SetCellValue(74, 2, new TFormula("='Budget_Valor de M Obra'!B89"));

            fmt = xls.GetCellVisibleFormatDef(74, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 3, xls.AddFormat(fmt));
            xls.SetCellValue(74, 3, new TFormula("='Budget_Valor de M Obra'!C89"));

            fmt = xls.GetCellVisibleFormatDef(74, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 4, xls.AddFormat(fmt));
            xls.SetCellValue(74, 4, new TFormula("='Budget_Valor de M Obra'!D89"));

            fmt = xls.GetCellVisibleFormatDef(74, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 5, xls.AddFormat(fmt));
            xls.SetCellValue(74, 5, new TFormula("='Budget_Valor de M Obra'!E89"));

            fmt = xls.GetCellVisibleFormatDef(74, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 6, xls.AddFormat(fmt));
            xls.SetCellValue(74, 6, new TFormula("='Budget_Valor de M Obra'!F89"));

            fmt = xls.GetCellVisibleFormatDef(74, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 7, xls.AddFormat(fmt));
            xls.SetCellValue(74, 7, new TFormula("='Budget_Valor de M Obra'!G89"));

            fmt = xls.GetCellVisibleFormatDef(74, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 8, xls.AddFormat(fmt));
            xls.SetCellValue(74, 8, new TFormula("='Budget_Valor de M Obra'!H89"));

            fmt = xls.GetCellVisibleFormatDef(74, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 9, xls.AddFormat(fmt));
            xls.SetCellValue(74, 9, new TFormula("='Budget_Valor de M Obra'!I89"));

            fmt = xls.GetCellVisibleFormatDef(74, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 10, xls.AddFormat(fmt));
            xls.SetCellValue(74, 10, new TFormula("='Budget_Valor de M Obra'!J89"));

            fmt = xls.GetCellVisibleFormatDef(74, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(74, 11, xls.AddFormat(fmt));
            xls.SetCellValue(74, 11, new TFormula("=(Proportions!$E$5*((Budget_Presupuesto!D74+Budget_Presupuesto!E74)/2))+(Proportions!$E$6*((Budget_Presupuesto!F74+Budget_Presupuesto!G74+Budget_Presupuesto!H74)"
            + "/3)+(Proportions!$E$7*((Budget_Presupuesto!I74+Budget_Presupuesto!J74)/2)))"));

            fmt = xls.GetCellVisibleFormatDef(75, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(75, 1, xls.AddFormat(fmt));
            xls.SetCellValue(75, 1, "Total otros costos fijos");

            fmt = xls.GetCellVisibleFormatDef(75, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 2, xls.AddFormat(fmt));
            xls.SetCellValue(75, 2, new TFormula("=SUM(B72:B74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 3, xls.AddFormat(fmt));
            xls.SetCellValue(75, 3, new TFormula("=SUM(C72:C74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 4, xls.AddFormat(fmt));
            xls.SetCellValue(75, 4, new TFormula("=SUM(D72:D74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 5, xls.AddFormat(fmt));
            xls.SetCellValue(75, 5, new TFormula("=SUM(E72:E74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 6, xls.AddFormat(fmt));
            xls.SetCellValue(75, 6, new TFormula("=SUM(F72:F74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 7, xls.AddFormat(fmt));
            xls.SetCellValue(75, 7, new TFormula("=SUM(G72:G74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 8, xls.AddFormat(fmt));
            xls.SetCellValue(75, 8, new TFormula("=SUM(H72:H74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 9, xls.AddFormat(fmt));
            xls.SetCellValue(75, 9, new TFormula("=SUM(I72:I74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 10);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 10, xls.AddFormat(fmt));
            xls.SetCellValue(75, 10, new TFormula("=SUM(J72:J74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(75, 11, xls.AddFormat(fmt));
            xls.SetCellValue(75, 11, new TFormula("=SUM(K72:K74)"));

            fmt = xls.GetCellVisibleFormatDef(75, 12);
            fmt.Format = "0";
            xls.SetCellFormat(75, 12, xls.AddFormat(fmt));
            xls.SetCellValue(75, 12, new TFormula("=AVERAGE(D75:J75)"));

            fmt = xls.GetCellVisibleFormatDef(76, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(76, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(77, 1, xls.AddFormat(fmt));
            xls.SetCellValue(77, 1, "Total Fixed Costs");

            fmt = xls.GetCellVisibleFormatDef(77, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 2, xls.AddFormat(fmt));
            xls.SetCellValue(77, 2, new TFormula("=B58+B64+B69+B75"));

            fmt = xls.GetCellVisibleFormatDef(77, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 3, xls.AddFormat(fmt));
            xls.SetCellValue(77, 3, new TFormula("=C58+C64+C69+C75"));

            fmt = xls.GetCellVisibleFormatDef(77, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 4, xls.AddFormat(fmt));
            xls.SetCellValue(77, 4, new TFormula("=D58+D64+D69+D75"));

            fmt = xls.GetCellVisibleFormatDef(77, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 5, xls.AddFormat(fmt));
            xls.SetCellValue(77, 5, new TFormula("=E58+E64+E69+E75"));

            fmt = xls.GetCellVisibleFormatDef(77, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 6, xls.AddFormat(fmt));
            xls.SetCellValue(77, 6, new TFormula("=F58+F64+F69+F75"));

            fmt = xls.GetCellVisibleFormatDef(77, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 7, xls.AddFormat(fmt));
            xls.SetCellValue(77, 7, new TFormula("=G58+G64+G69+G75"));

            fmt = xls.GetCellVisibleFormatDef(77, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 8, xls.AddFormat(fmt));
            xls.SetCellValue(77, 8, new TFormula("=H58+H64+H69+H75"));

            fmt = xls.GetCellVisibleFormatDef(77, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 9, xls.AddFormat(fmt));
            xls.SetCellValue(77, 9, new TFormula("=I58+I64+I69+I75"));

            fmt = xls.GetCellVisibleFormatDef(77, 10);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 10, xls.AddFormat(fmt));
            xls.SetCellValue(77, 10, new TFormula("=J58+J64+J69+J75"));

            fmt = xls.GetCellVisibleFormatDef(77, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(77, 11, xls.AddFormat(fmt));
            xls.SetCellValue(77, 11, new TFormula("=K58+K64+K69+K75"));

            fmt = xls.GetCellVisibleFormatDef(77, 12);
            fmt.Format = "0";
            xls.SetCellFormat(77, 12, xls.AddFormat(fmt));
            xls.SetCellValue(77, 12, new TFormula("=AVERAGE(D77:J77)"));

            fmt = xls.GetCellVisibleFormatDef(78, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(78, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(79, 1, xls.AddFormat(fmt));
            xls.SetCellValue(79, 1, "TOTAL COSTS");

            fmt = xls.GetCellVisibleFormatDef(79, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 2, xls.AddFormat(fmt));
            xls.SetCellValue(79, 2, new TFormula("=B46+B77"));

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));
            xls.SetCellValue(79, 3, new TFormula("=C46+C77"));

            fmt = xls.GetCellVisibleFormatDef(79, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 4, xls.AddFormat(fmt));
            xls.SetCellValue(79, 4, new TFormula("=D46+D77"));

            fmt = xls.GetCellVisibleFormatDef(79, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 5, xls.AddFormat(fmt));
            xls.SetCellValue(79, 5, new TFormula("=E46+E77"));

            fmt = xls.GetCellVisibleFormatDef(79, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 6, xls.AddFormat(fmt));
            xls.SetCellValue(79, 6, new TFormula("=F46+F77"));

            fmt = xls.GetCellVisibleFormatDef(79, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 7, xls.AddFormat(fmt));
            xls.SetCellValue(79, 7, new TFormula("=G46+G77"));

            fmt = xls.GetCellVisibleFormatDef(79, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 8, xls.AddFormat(fmt));
            xls.SetCellValue(79, 8, new TFormula("=H46+H77"));

            fmt = xls.GetCellVisibleFormatDef(79, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 9, xls.AddFormat(fmt));
            xls.SetCellValue(79, 9, new TFormula("=I46+I77"));

            fmt = xls.GetCellVisibleFormatDef(79, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 10, xls.AddFormat(fmt));
            xls.SetCellValue(79, 10, new TFormula("=J46+J77"));

            fmt = xls.GetCellVisibleFormatDef(79, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(79, 11, xls.AddFormat(fmt));
            xls.SetCellValue(79, 11, new TFormula("=K46+K77"));

            fmt = xls.GetCellVisibleFormatDef(79, 12);
            fmt.Format = "0";
            xls.SetCellFormat(79, 12, xls.AddFormat(fmt));
            xls.SetCellValue(79, 12, new TFormula("=AVERAGE(D79:J79)"));

            fmt = xls.GetCellVisibleFormatDef(80, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(80, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(81, 1, xls.AddFormat(fmt));
            xls.SetCellValue(81, 1, "NET REVENUE");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 2, xls.AddFormat(fmt));
            xls.SetCellValue(81, 2, new TFormula("=B31-B79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 3, xls.AddFormat(fmt));
            xls.SetCellValue(81, 3, new TFormula("=C31-C79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 4, xls.AddFormat(fmt));
            xls.SetCellValue(81, 4, new TFormula("=D31-D79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 5, xls.AddFormat(fmt));
            xls.SetCellValue(81, 5, new TFormula("=E31-E79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 6, xls.AddFormat(fmt));
            xls.SetCellValue(81, 6, new TFormula("=F31-F79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 7, xls.AddFormat(fmt));
            xls.SetCellValue(81, 7, new TFormula("=G31-G79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 8, xls.AddFormat(fmt));
            xls.SetCellValue(81, 8, new TFormula("=H31-H79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 9, xls.AddFormat(fmt));
            xls.SetCellValue(81, 9, new TFormula("=I31-I79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 10, xls.AddFormat(fmt));
            xls.SetCellValue(81, 10, new TFormula("=J31-J79"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Currency, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(81, 11, xls.AddFormat(fmt));
            xls.SetCellValue(81, 11, new TFormula("=K31-K79"));

            fmt = xls.GetCellVisibleFormatDef(81, 12);
            fmt.Format = "0";
            xls.SetCellFormat(81, 12, xls.AddFormat(fmt));
            xls.SetCellValue(81, 12, new TFormula("=AVERAGE(D81:J81)"));

            fmt = xls.GetCellVisibleFormatDef(82, 1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(82, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 2);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 3);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 4);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 5);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 6);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 7);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 8);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 9);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 10);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(82, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 3);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 4);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 6);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 7);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 8);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 9);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 10);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 11);
            fmt.Format = "#,##0";
            xls.SetCellFormat(83, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 2);
            fmt.Format = "0.0";
            xls.SetCellFormat(86, 2, xls.AddFormat(fmt));

            //Freeze Panes
            xls.FreezePanes(new TCellAddress("B2"));

            //Cell selection and scroll position.
            xls.SelectCell(51, 2, false);

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
