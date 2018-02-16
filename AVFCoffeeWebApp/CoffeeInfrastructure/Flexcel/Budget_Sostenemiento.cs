using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class Budget_Sostenemiento
    {

        public void BudgetSostenemiento(ExcelFile xls)
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

            xls.ActiveSheet = 15;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Budget_Sostenemiento";

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
            xls.PrintScale = 49;
            xls.PrintXResolution = 600;
            xls.PrintYResolution = 600;
            xls.PrintOptions = TPrintOptions.None;
            xls.PrintPaperSize = TPaperSize.Letter;

            //Set up rows and columns
            xls.DefaultColWidth = 2816;

            xls.SetColWidth(1, 1, 13610);    //(52.41 + 0.75) * 256

            xls.SetColWidth(2, 5, 3541);    //(13.08 + 0.75) * 256

            xls.SetColWidth(6, 6, 3712);    //(13.75 + 0.75) * 256

            xls.SetColWidth(7, 10, 3541);    //(13.08 + 0.75) * 256

            xls.SetColWidth(11, 11, 7424);    //(28.25 + 0.75) * 256

            xls.SetColWidth(15, 15, 4608);    //(17.25 + 0.75) * 256

            xls.SetRowHeight(3, 370);    //18.50 * 20
            xls.SetRowHeight(27, 740);    //37.00 * 20
            xls.SetRowHeight(28, 370);    //18.50 * 20
            xls.SetRowHeight(35, 370);    //18.50 * 20
            xls.SetRowHeight(36, 370);    //18.50 * 20
            xls.SetRowHeight(40, 370);    //18.50 * 20
            xls.SetRowHeight(53, 620);    //31.00 * 20

            //Merged Cells
            xls.MergeCells(55, 1, 55, 5);
            xls.MergeCells(58, 1, 58, 5);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(1, 1, xls.AddFormat(fmt));
            xls.SetCellValue(1, 1, "Cuadro. Sostenimiento. Costos Variables detallados");

            fmt = xls.GetCellVisibleFormatDef(1, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(1, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(1, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 1, xls.AddFormat(fmt));
            xls.SetCellValue(2, 1, "Año 2- 8");

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));
            xls.SetCellValue(3, 1, "Mantenimiento, fertilización y control de plagas");

            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));
            xls.SetCellValue(3, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));
            xls.SetCellValue(3, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));
            xls.SetCellValue(3, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));
            xls.SetCellValue(3, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(3, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 6, xls.AddFormat(fmt));
            xls.SetCellValue(3, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(3, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 7, xls.AddFormat(fmt));
            xls.SetCellValue(3, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));
            xls.SetCellValue(3, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(3, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 9, xls.AddFormat(fmt));
            xls.SetCellValue(3, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(3, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 10, xls.AddFormat(fmt));
            xls.SetCellValue(3, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(3, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 11, xls.AddFormat(fmt));
            xls.SetCellValue(3, 11, "Datos para ajuste");

            fmt = xls.GetCellVisibleFormatDef(4, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 1, xls.AddFormat(fmt));
            xls.SetCellValue(4, 1, "Mano de obra mantenimiento, fertilización y control de plagas");

            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));
            xls.SetCellValue(4, 4, new TFormula("='Budget_Valor de M Obra'!D59"));

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));
            xls.SetCellValue(4, 5, new TFormula("='Budget_Valor de M Obra'!E59"));

            fmt = xls.GetCellVisibleFormatDef(4, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(4, 6, xls.AddFormat(fmt));
            xls.SetCellValue(4, 6, new TFormula("='Budget_Valor de M Obra'!F59"));

            fmt = xls.GetCellVisibleFormatDef(4, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(4, 7, xls.AddFormat(fmt));
            xls.SetCellValue(4, 7, new TFormula("='Budget_Valor de M Obra'!G59"));

            fmt = xls.GetCellVisibleFormatDef(4, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(4, 8, xls.AddFormat(fmt));
            xls.SetCellValue(4, 8, new TFormula("='Budget_Valor de M Obra'!H59"));

            fmt = xls.GetCellVisibleFormatDef(4, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(4, 9, xls.AddFormat(fmt));
            xls.SetCellValue(4, 9, new TFormula("='Budget_Valor de M Obra'!I59"));

            fmt = xls.GetCellVisibleFormatDef(4, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(4, 10, xls.AddFormat(fmt));
            xls.SetCellValue(4, 10, new TFormula("='Budget_Valor de M Obra'!J59"));

            fmt = xls.GetCellVisibleFormatDef(4, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(5, 1, xls.AddFormat(fmt));
            xls.SetCellValue(5, 1, "Materiales para fertilización y control plagas:");

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(6, 1, xls.AddFormat(fmt));
            xls.SetCellValue(6, 1, "Abonos");

            fmt = xls.GetCellVisibleFormatDef(7, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(7, 1, xls.AddFormat(fmt));
            xls.SetCellValue(7, 1, new TFormula("=Budget_Supuestos!A279"));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));
            xls.SetCellValue(7, 4, new TFormula("=Budget_Supuestos!B279"));

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));
            xls.SetCellValue(7, 5, new TFormula("=Budget_Supuestos!B279"));

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));
            xls.SetCellValue(7, 6, new TFormula("=Budget_Supuestos!B279"));

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));
            xls.SetCellValue(7, 7, new TFormula("=Budget_Supuestos!B279"));

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));
            xls.SetCellValue(7, 8, new TFormula("=Budget_Supuestos!B279"));

            fmt = xls.GetCellVisibleFormatDef(7, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 9, xls.AddFormat(fmt));
            xls.SetCellValue(7, 9, new TFormula("=Budget_Supuestos!B279"));

            fmt = xls.GetCellVisibleFormatDef(7, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(7, 10, xls.AddFormat(fmt));
            xls.SetCellValue(7, 10, new TFormula("=Budget_Supuestos!B279"));

            fmt = xls.GetCellVisibleFormatDef(8, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(8, 1, xls.AddFormat(fmt));
            xls.SetCellValue(8, 1, new TFormula("=Budget_Supuestos!A280"));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));
            xls.SetCellValue(8, 4, new TFormula("=Budget_Supuestos!B280"));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));
            xls.SetCellValue(8, 5, new TFormula("=Budget_Supuestos!B280"));

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));
            xls.SetCellValue(8, 6, new TFormula("=Budget_Supuestos!B280"));

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));
            xls.SetCellValue(8, 7, new TFormula("=Budget_Supuestos!B280"));

            fmt = xls.GetCellVisibleFormatDef(8, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 8, xls.AddFormat(fmt));
            xls.SetCellValue(8, 8, new TFormula("=Budget_Supuestos!B280"));

            fmt = xls.GetCellVisibleFormatDef(8, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 9, xls.AddFormat(fmt));
            xls.SetCellValue(8, 9, new TFormula("=Budget_Supuestos!B280"));

            fmt = xls.GetCellVisibleFormatDef(8, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(8, 10, xls.AddFormat(fmt));
            xls.SetCellValue(8, 10, new TFormula("=Budget_Supuestos!B280"));

            fmt = xls.GetCellVisibleFormatDef(9, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(9, 1, xls.AddFormat(fmt));
            xls.SetCellValue(9, 1, new TFormula("=Budget_Supuestos!A281"));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));
            xls.SetCellValue(9, 4, new TFormula("=Budget_Supuestos!B281"));

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));
            xls.SetCellValue(9, 5, new TFormula("=Budget_Supuestos!B281"));

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));
            xls.SetCellValue(9, 6, new TFormula("=Budget_Supuestos!B281"));

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));
            xls.SetCellValue(9, 7, new TFormula("=Budget_Supuestos!B281"));

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));
            xls.SetCellValue(9, 8, new TFormula("=Budget_Supuestos!B281"));

            fmt = xls.GetCellVisibleFormatDef(9, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 9, xls.AddFormat(fmt));
            xls.SetCellValue(9, 9, new TFormula("=Budget_Supuestos!B281"));

            fmt = xls.GetCellVisibleFormatDef(9, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 10, xls.AddFormat(fmt));
            xls.SetCellValue(9, 10, new TFormula("=Budget_Supuestos!B281"));

            fmt = xls.GetCellVisibleFormatDef(10, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(10, 1, xls.AddFormat(fmt));
            xls.SetCellValue(10, 1, new TFormula("=Budget_Supuestos!A282"));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));
            xls.SetCellValue(10, 4, new TFormula("=Budget_Supuestos!B282"));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
            xls.SetCellValue(10, 5, new TFormula("=Budget_Supuestos!B282"));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));
            xls.SetCellValue(10, 6, new TFormula("=Budget_Supuestos!B282"));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));
            xls.SetCellValue(10, 7, new TFormula("=Budget_Supuestos!B282"));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));
            xls.SetCellValue(10, 8, new TFormula("=Budget_Supuestos!B282"));

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));
            xls.SetCellValue(10, 9, new TFormula("=Budget_Supuestos!B282"));

            fmt = xls.GetCellVisibleFormatDef(10, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(10, 10, xls.AddFormat(fmt));
            xls.SetCellValue(10, 10, new TFormula("=Budget_Supuestos!B282"));

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));
            xls.SetCellValue(11, 1, new TFormula("=Budget_Supuestos!A283"));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));
            xls.SetCellValue(11, 4, new TFormula("=Budget_Supuestos!B283"));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));
            xls.SetCellValue(11, 5, new TFormula("=Budget_Supuestos!B283"));

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));
            xls.SetCellValue(11, 6, new TFormula("=Budget_Supuestos!B283"));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));
            xls.SetCellValue(11, 7, new TFormula("=Budget_Supuestos!B283"));

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));
            xls.SetCellValue(11, 8, new TFormula("=Budget_Supuestos!B283"));

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));
            xls.SetCellValue(11, 9, new TFormula("=Budget_Supuestos!B283"));

            fmt = xls.GetCellVisibleFormatDef(11, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(11, 10, xls.AddFormat(fmt));
            xls.SetCellValue(11, 10, new TFormula("=Budget_Supuestos!B283"));

            fmt = xls.GetCellVisibleFormatDef(12, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(12, 1, xls.AddFormat(fmt));
            xls.SetCellValue(12, 1, new TFormula("=Budget_Supuestos!A284"));

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));
            xls.SetCellValue(12, 4, new TFormula("=Budget_Supuestos!B284"));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));
            xls.SetCellValue(12, 5, new TFormula("=Budget_Supuestos!B284"));

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));
            xls.SetCellValue(12, 6, new TFormula("=Budget_Supuestos!B284"));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));
            xls.SetCellValue(12, 7, new TFormula("=Budget_Supuestos!B284"));

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));
            xls.SetCellValue(12, 8, new TFormula("=Budget_Supuestos!B284"));

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));
            xls.SetCellValue(12, 9, new TFormula("=Budget_Supuestos!B284"));

            fmt = xls.GetCellVisibleFormatDef(12, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(12, 10, xls.AddFormat(fmt));
            xls.SetCellValue(12, 10, new TFormula("=Budget_Supuestos!B284"));

            fmt = xls.GetCellVisibleFormatDef(13, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(13, 1, xls.AddFormat(fmt));
            xls.SetCellValue(13, 1, new TFormula("=Budget_Supuestos!A285"));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, new TFormula("=Budget_Supuestos!B285"));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));
            xls.SetCellValue(13, 5, new TFormula("=Budget_Supuestos!B285"));

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));
            xls.SetCellValue(13, 6, new TFormula("=Budget_Supuestos!B285"));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));
            xls.SetCellValue(13, 7, new TFormula("=Budget_Supuestos!B285"));

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));
            xls.SetCellValue(13, 8, new TFormula("=Budget_Supuestos!B285"));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));
            xls.SetCellValue(13, 9, new TFormula("=Budget_Supuestos!B285"));

            fmt = xls.GetCellVisibleFormatDef(13, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 10, xls.AddFormat(fmt));
            xls.SetCellValue(13, 10, new TFormula("=Budget_Supuestos!B285"));

            fmt = xls.GetCellVisibleFormatDef(14, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 1, xls.AddFormat(fmt));
            xls.SetCellValue(14, 1, new TFormula("=Budget_Supuestos!A286"));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, new TFormula("=Budget_Supuestos!B286"));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));
            xls.SetCellValue(14, 5, new TFormula("=Budget_Supuestos!B286"));

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));
            xls.SetCellValue(14, 6, new TFormula("=Budget_Supuestos!B286"));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));
            xls.SetCellValue(14, 7, new TFormula("=Budget_Supuestos!B286"));

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));
            xls.SetCellValue(14, 8, new TFormula("=Budget_Supuestos!B286"));

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));
            xls.SetCellValue(14, 9, new TFormula("=Budget_Supuestos!B286"));

            fmt = xls.GetCellVisibleFormatDef(14, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(14, 10, xls.AddFormat(fmt));
            xls.SetCellValue(14, 10, new TFormula("=Budget_Supuestos!B286"));

            fmt = xls.GetCellVisibleFormatDef(15, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 1, xls.AddFormat(fmt));
            xls.SetCellValue(15, 1, new TFormula("=Budget_Supuestos!A287"));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));
            xls.SetCellValue(15, 4, new TFormula("=Budget_Supuestos!B287"));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));
            xls.SetCellValue(15, 5, new TFormula("=Budget_Supuestos!B287"));

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));
            xls.SetCellValue(15, 6, new TFormula("=Budget_Supuestos!B287"));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));
            xls.SetCellValue(15, 7, new TFormula("=Budget_Supuestos!B287"));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));
            xls.SetCellValue(15, 8, new TFormula("=Budget_Supuestos!B287"));

            fmt = xls.GetCellVisibleFormatDef(15, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(15, 9, xls.AddFormat(fmt));
            xls.SetCellValue(15, 9, new TFormula("=Budget_Supuestos!B287"));

            fmt = xls.GetCellVisibleFormatDef(15, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(15, 10, xls.AddFormat(fmt));
            xls.SetCellValue(15, 10, new TFormula("=Budget_Supuestos!B287"));

            fmt = xls.GetCellVisibleFormatDef(16, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(16, 1, xls.AddFormat(fmt));
            xls.SetCellValue(16, 1, new TFormula("=Budget_Supuestos!A288"));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));
            xls.SetCellValue(16, 4, new TFormula("=Budget_Supuestos!B288"));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));
            xls.SetCellValue(16, 5, new TFormula("=Budget_Supuestos!B288"));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));
            xls.SetCellValue(16, 6, new TFormula("=Budget_Supuestos!B288"));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));
            xls.SetCellValue(16, 7, new TFormula("=Budget_Supuestos!B288"));

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));
            xls.SetCellValue(16, 8, new TFormula("=Budget_Supuestos!B288"));

            fmt = xls.GetCellVisibleFormatDef(16, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(16, 9, xls.AddFormat(fmt));
            xls.SetCellValue(16, 9, new TFormula("=Budget_Supuestos!B288"));

            fmt = xls.GetCellVisibleFormatDef(16, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(16, 10, xls.AddFormat(fmt));
            xls.SetCellValue(16, 10, new TFormula("=Budget_Supuestos!B288"));

            fmt = xls.GetCellVisibleFormatDef(17, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 1, xls.AddFormat(fmt));
            xls.SetCellValue(17, 1, new TFormula("=Budget_Supuestos!A289"));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));
            xls.SetCellValue(17, 4, new TFormula("=Budget_Supuestos!B289"));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));
            xls.SetCellValue(17, 5, new TFormula("=Budget_Supuestos!B289"));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));
            xls.SetCellValue(17, 6, new TFormula("=Budget_Supuestos!B289"));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));
            xls.SetCellValue(17, 7, new TFormula("=Budget_Supuestos!B289"));

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));
            xls.SetCellValue(17, 8, new TFormula("=Budget_Supuestos!B289"));

            fmt = xls.GetCellVisibleFormatDef(17, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(17, 9, xls.AddFormat(fmt));
            xls.SetCellValue(17, 9, new TFormula("=Budget_Supuestos!B289"));

            fmt = xls.GetCellVisibleFormatDef(17, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(17, 10, xls.AddFormat(fmt));
            xls.SetCellValue(17, 10, new TFormula("=Budget_Supuestos!B289"));

            fmt = xls.GetCellVisibleFormatDef(18, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(18, 1, xls.AddFormat(fmt));
            xls.SetCellValue(18, 1, new TFormula("=Budget_Supuestos!A290"));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));
            xls.SetCellValue(18, 4, new TFormula("=Budget_Supuestos!B290"));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));
            xls.SetCellValue(18, 5, new TFormula("=Budget_Supuestos!B290"));

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));
            xls.SetCellValue(18, 6, new TFormula("=Budget_Supuestos!B290"));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));
            xls.SetCellValue(18, 7, new TFormula("=Budget_Supuestos!B290"));

            fmt = xls.GetCellVisibleFormatDef(18, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(18, 8, xls.AddFormat(fmt));
            xls.SetCellValue(18, 8, new TFormula("=Budget_Supuestos!B290"));

            fmt = xls.GetCellVisibleFormatDef(18, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(18, 9, xls.AddFormat(fmt));
            xls.SetCellValue(18, 9, new TFormula("=Budget_Supuestos!B290"));

            fmt = xls.GetCellVisibleFormatDef(18, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(18, 10, xls.AddFormat(fmt));
            xls.SetCellValue(18, 10, new TFormula("=Budget_Supuestos!B290"));

            fmt = xls.GetCellVisibleFormatDef(19, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(19, 1, xls.AddFormat(fmt));
            xls.SetCellValue(19, 1, new TFormula("=Budget_Supuestos!A291"));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));
            xls.SetCellValue(19, 4, new TFormula("=Budget_Supuestos!B291"));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));
            xls.SetCellValue(19, 5, new TFormula("=Budget_Supuestos!B291"));

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));
            xls.SetCellValue(19, 6, new TFormula("=Budget_Supuestos!B291"));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));
            xls.SetCellValue(19, 7, new TFormula("=Budget_Supuestos!B291"));

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));
            xls.SetCellValue(19, 8, new TFormula("=Budget_Supuestos!B291"));

            fmt = xls.GetCellVisibleFormatDef(19, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(19, 9, xls.AddFormat(fmt));
            xls.SetCellValue(19, 9, new TFormula("=Budget_Supuestos!B291"));

            fmt = xls.GetCellVisibleFormatDef(19, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(19, 10, xls.AddFormat(fmt));
            xls.SetCellValue(19, 10, new TFormula("=Budget_Supuestos!B291"));

            fmt = xls.GetCellVisibleFormatDef(20, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(20, 1, xls.AddFormat(fmt));
            xls.SetCellValue(20, 1, new TFormula("=Budget_Supuestos!A292"));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));
            xls.SetCellValue(20, 4, new TFormula("=Budget_Supuestos!B292"));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));
            xls.SetCellValue(20, 5, new TFormula("=Budget_Supuestos!B292"));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));
            xls.SetCellValue(20, 6, new TFormula("=Budget_Supuestos!B292"));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));
            xls.SetCellValue(20, 7, new TFormula("=Budget_Supuestos!B292"));

            fmt = xls.GetCellVisibleFormatDef(20, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(20, 8, xls.AddFormat(fmt));
            xls.SetCellValue(20, 8, new TFormula("=Budget_Supuestos!B292"));

            fmt = xls.GetCellVisibleFormatDef(20, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(20, 9, xls.AddFormat(fmt));
            xls.SetCellValue(20, 9, new TFormula("=Budget_Supuestos!B292"));

            fmt = xls.GetCellVisibleFormatDef(20, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(20, 10, xls.AddFormat(fmt));
            xls.SetCellValue(20, 10, new TFormula("=Budget_Supuestos!B292"));

            fmt = xls.GetCellVisibleFormatDef(21, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(21, 1, xls.AddFormat(fmt));
            xls.SetCellValue(21, 1, new TFormula("=Budget_Supuestos!A293"));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));
            xls.SetCellValue(21, 4, new TFormula("=Budget_Supuestos!B293"));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));
            xls.SetCellValue(21, 5, new TFormula("=Budget_Supuestos!B293"));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));
            xls.SetCellValue(21, 6, new TFormula("=Budget_Supuestos!B293"));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));
            xls.SetCellValue(21, 7, new TFormula("=Budget_Supuestos!B293"));

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));
            xls.SetCellValue(21, 8, new TFormula("=Budget_Supuestos!B293"));

            fmt = xls.GetCellVisibleFormatDef(21, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(21, 9, xls.AddFormat(fmt));
            xls.SetCellValue(21, 9, new TFormula("=Budget_Supuestos!B293"));

            fmt = xls.GetCellVisibleFormatDef(21, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(21, 10, xls.AddFormat(fmt));
            xls.SetCellValue(21, 10, new TFormula("=Budget_Supuestos!B293"));

            fmt = xls.GetCellVisibleFormatDef(22, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(22, 1, xls.AddFormat(fmt));
            xls.SetCellValue(22, 1, new TFormula("=Budget_Supuestos!A294"));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));
            xls.SetCellValue(22, 4, new TFormula("=Budget_Supuestos!B294"));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));
            xls.SetCellValue(22, 5, new TFormula("=Budget_Supuestos!B294"));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));
            xls.SetCellValue(22, 6, new TFormula("=Budget_Supuestos!B294"));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));
            xls.SetCellValue(22, 7, new TFormula("=Budget_Supuestos!B294"));

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));
            xls.SetCellValue(22, 8, new TFormula("=Budget_Supuestos!B294"));

            fmt = xls.GetCellVisibleFormatDef(22, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(22, 9, xls.AddFormat(fmt));
            xls.SetCellValue(22, 9, new TFormula("=Budget_Supuestos!B294"));

            fmt = xls.GetCellVisibleFormatDef(22, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(22, 10, xls.AddFormat(fmt));
            xls.SetCellValue(22, 10, new TFormula("=Budget_Supuestos!B294"));

            fmt = xls.GetCellVisibleFormatDef(23, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(23, 1, xls.AddFormat(fmt));
            xls.SetCellValue(23, 1, "Total Fertilizaciones");

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));
            xls.SetCellValue(23, 4, new TFormula("=Budget_Supuestos!$B$296"));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));
            xls.SetCellValue(23, 5, new TFormula("=Budget_Supuestos!$B$296"));

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));
            xls.SetCellValue(23, 6, new TFormula("=Budget_Supuestos!$B$296"));

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));
            xls.SetCellValue(23, 7, new TFormula("=Budget_Supuestos!$B$296"));

            fmt = xls.GetCellVisibleFormatDef(23, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 8, xls.AddFormat(fmt));
            xls.SetCellValue(23, 8, new TFormula("=Budget_Supuestos!$B$296"));

            fmt = xls.GetCellVisibleFormatDef(23, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 9, xls.AddFormat(fmt));
            xls.SetCellValue(23, 9, new TFormula("=Budget_Supuestos!$B$296"));

            fmt = xls.GetCellVisibleFormatDef(23, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 10, xls.AddFormat(fmt));
            xls.SetCellValue(23, 10, new TFormula("=Budget_Supuestos!$B$296"));

            fmt = xls.GetCellVisibleFormatDef(24, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(24, 1, xls.AddFormat(fmt));
            xls.SetCellValue(24, 1, new TFormula("=Budget_Supuestos!A295"));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));
            xls.SetCellValue(24, 4, new TFormula("=Budget_Supuestos!B295"));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));
            xls.SetCellValue(24, 5, new TFormula("=Budget_Supuestos!B295"));

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));
            xls.SetCellValue(24, 6, new TFormula("=Budget_Supuestos!B295"));

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));
            xls.SetCellValue(24, 7, new TFormula("=Budget_Supuestos!B295"));

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));
            xls.SetCellValue(24, 8, new TFormula("=Budget_Supuestos!B295"));

            fmt = xls.GetCellVisibleFormatDef(24, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(24, 9, xls.AddFormat(fmt));
            xls.SetCellValue(24, 9, new TFormula("=Budget_Supuestos!B295"));

            fmt = xls.GetCellVisibleFormatDef(24, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(24, 10, xls.AddFormat(fmt));
            xls.SetCellValue(24, 10, new TFormula("=Budget_Supuestos!B295"));

            fmt = xls.GetCellVisibleFormatDef(25, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 1, xls.AddFormat(fmt));
            xls.SetCellValue(25, 1, "Total materiales fertilización y control de plagas");

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));
            xls.SetCellValue(25, 4, new TFormula("=SUM(D7:D22)+D24"));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));
            xls.SetCellValue(25, 5, new TFormula("=SUM(E7:E22)+E24"));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));
            xls.SetCellValue(25, 6, new TFormula("=SUM(F7:F22)+F24"));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));
            xls.SetCellValue(25, 7, new TFormula("=SUM(G7:G22)+G24"));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));
            xls.SetCellValue(25, 8, new TFormula("=SUM(H7:H22)+H24"));

            fmt = xls.GetCellVisibleFormatDef(25, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(25, 9, xls.AddFormat(fmt));
            xls.SetCellValue(25, 9, new TFormula("=SUM(I7:I22)+I24"));

            fmt = xls.GetCellVisibleFormatDef(25, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(25, 10, xls.AddFormat(fmt));
            xls.SetCellValue(25, 10, new TFormula("=SUM(J7:J22)+J24"));

            fmt = xls.GetCellVisibleFormatDef(26, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 1, xls.AddFormat(fmt));
            xls.SetCellValue(26, 1, "Total costos transporte mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));
            xls.SetCellValue(26, 4, new TFormula("=$F$55"));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));
            xls.SetCellValue(26, 5, new TFormula("=$F$55"));

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));
            xls.SetCellValue(26, 6, new TFormula("=$F$55"));

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));
            xls.SetCellValue(26, 7, new TFormula("=$F$55"));

            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));
            xls.SetCellValue(26, 8, new TFormula("=$F$55"));

            fmt = xls.GetCellVisibleFormatDef(26, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 9, xls.AddFormat(fmt));
            xls.SetCellValue(26, 9, new TFormula("=$F$55"));

            fmt = xls.GetCellVisibleFormatDef(26, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(26, 10, xls.AddFormat(fmt));
            xls.SetCellValue(26, 10, new TFormula("=$F$55"));

            fmt = xls.GetCellVisibleFormatDef(27, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(27, 1, xls.AddFormat(fmt));
            xls.SetCellValue(27, 1, "Total costos variables mantenimiento, fertilización y control de plagas");

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));
            xls.SetCellValue(27, 4, new TFormula("=D25+D4+D26"));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));
            xls.SetCellValue(27, 5, new TFormula("=E25+E4+E26"));

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));
            xls.SetCellValue(27, 6, new TFormula("=F25+F4+F26"));

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));
            xls.SetCellValue(27, 7, new TFormula("=G25+G4+G26"));

            fmt = xls.GetCellVisibleFormatDef(27, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(27, 8, xls.AddFormat(fmt));
            xls.SetCellValue(27, 8, new TFormula("=H25+H4+H26"));

            fmt = xls.GetCellVisibleFormatDef(27, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(27, 9, xls.AddFormat(fmt));
            xls.SetCellValue(27, 9, new TFormula("=I25+I4+I26"));

            fmt = xls.GetCellVisibleFormatDef(27, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(27, 10, xls.AddFormat(fmt));
            xls.SetCellValue(27, 10, new TFormula("=J25+J4+J26"));

            fmt = xls.GetCellVisibleFormatDef(28, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(28, 1, xls.AddFormat(fmt));
            xls.SetCellValue(28, 1, "Cosecha");

            fmt = xls.GetCellVisibleFormatDef(28, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));
            xls.SetCellValue(28, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));
            xls.SetCellValue(28, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));
            xls.SetCellValue(28, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));
            xls.SetCellValue(28, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));
            xls.SetCellValue(28, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));
            xls.SetCellValue(28, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));
            xls.SetCellValue(28, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(28, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 9, xls.AddFormat(fmt));
            xls.SetCellValue(28, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(28, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 10, xls.AddFormat(fmt));
            xls.SetCellValue(28, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(29, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(29, 1, xls.AddFormat(fmt));
            xls.SetCellValue(29, 1, "Mano de obra cosecha");

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));
            xls.SetCellValue(29, 4, new TFormula("='Budget_Valor de M Obra'!D64"));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));
            xls.SetCellValue(29, 5, new TFormula("='Budget_Valor de M Obra'!E64"));

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));
            xls.SetCellValue(29, 6, new TFormula("='Budget_Valor de M Obra'!F64"));

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));
            xls.SetCellValue(29, 7, new TFormula("='Budget_Valor de M Obra'!G64"));

            fmt = xls.GetCellVisibleFormatDef(29, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(29, 8, xls.AddFormat(fmt));
            xls.SetCellValue(29, 8, new TFormula("='Budget_Valor de M Obra'!H64"));

            fmt = xls.GetCellVisibleFormatDef(29, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(29, 9, xls.AddFormat(fmt));
            xls.SetCellValue(29, 9, new TFormula("='Budget_Valor de M Obra'!I64"));

            fmt = xls.GetCellVisibleFormatDef(29, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(29, 10, xls.AddFormat(fmt));
            xls.SetCellValue(29, 10, new TFormula("='Budget_Valor de M Obra'!J64"));

            fmt = xls.GetCellVisibleFormatDef(30, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(30, 1, xls.AddFormat(fmt));
            xls.SetCellValue(30, 1, "Materiales para la cosecha:");

            fmt = xls.GetCellVisibleFormatDef(31, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(31, 1, xls.AddFormat(fmt));
            xls.SetCellValue(31, 1, "Sacos para la recoleccion");

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));
            xls.SetCellValue(31, 4, new TFormula("=K31*$D$48"));

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));
            xls.SetCellValue(31, 5, new TFormula("=K31*$E$48"));

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));
            xls.SetCellValue(31, 6, new TFormula("=K31*$F$48"));

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));
            xls.SetCellValue(31, 7, new TFormula("=K31*$G$48"));

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));
            xls.SetCellValue(31, 8, new TFormula("=K31*$H$48"));

            fmt = xls.GetCellVisibleFormatDef(31, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 9, xls.AddFormat(fmt));
            xls.SetCellValue(31, 9, new TFormula("=K31*$I$48"));

            fmt = xls.GetCellVisibleFormatDef(31, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 10, xls.AddFormat(fmt));
            xls.SetCellValue(31, 10, new TFormula("=K31*$J$48"));

            fmt = xls.GetCellVisibleFormatDef(31, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(31, 11, xls.AddFormat(fmt));
            xls.SetCellValue(31, 11, new TFormula("=Budget_Supuestos!B339"));

            fmt = xls.GetCellVisibleFormatDef(32, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(32, 1, xls.AddFormat(fmt));
            xls.SetCellValue(32, 1, "Sacos Pergamino");

            fmt = xls.GetCellVisibleFormatDef(32, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 4, xls.AddFormat(fmt));
            xls.SetCellValue(32, 4, new TFormula("=K32*$D$48"));

            fmt = xls.GetCellVisibleFormatDef(32, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 5, xls.AddFormat(fmt));
            xls.SetCellValue(32, 5, new TFormula("=K32*$E$48"));

            fmt = xls.GetCellVisibleFormatDef(32, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 6, xls.AddFormat(fmt));
            xls.SetCellValue(32, 6, new TFormula("=K32*$F$48"));

            fmt = xls.GetCellVisibleFormatDef(32, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 7, xls.AddFormat(fmt));
            xls.SetCellValue(32, 7, new TFormula("=K32*$G$48"));

            fmt = xls.GetCellVisibleFormatDef(32, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 8, xls.AddFormat(fmt));
            xls.SetCellValue(32, 8, new TFormula("=K32*$H$48"));

            fmt = xls.GetCellVisibleFormatDef(32, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 9, xls.AddFormat(fmt));
            xls.SetCellValue(32, 9, new TFormula("=K32*$I$48"));

            fmt = xls.GetCellVisibleFormatDef(32, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 10, xls.AddFormat(fmt));
            xls.SetCellValue(32, 10, new TFormula("=K32*$J$48"));

            fmt = xls.GetCellVisibleFormatDef(32, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(32, 11, xls.AddFormat(fmt));
            xls.SetCellValue(32, 11, new TFormula("=Budget_Supuestos!B340"));

            fmt = xls.GetCellVisibleFormatDef(33, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(33, 1, xls.AddFormat(fmt));
            xls.SetCellValue(33, 1, "Cabuya:");

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));
            xls.SetCellValue(33, 4, new TFormula("=K33*$D$48"));

            fmt = xls.GetCellVisibleFormatDef(33, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 5, xls.AddFormat(fmt));
            xls.SetCellValue(33, 5, new TFormula("=K33*$E$48"));

            fmt = xls.GetCellVisibleFormatDef(33, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 6, xls.AddFormat(fmt));
            xls.SetCellValue(33, 6, new TFormula("=K33*$F$48"));

            fmt = xls.GetCellVisibleFormatDef(33, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 7, xls.AddFormat(fmt));
            xls.SetCellValue(33, 7, new TFormula("=K33*$G$48"));

            fmt = xls.GetCellVisibleFormatDef(33, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 8, xls.AddFormat(fmt));
            xls.SetCellValue(33, 8, new TFormula("=K33*$H$48"));

            fmt = xls.GetCellVisibleFormatDef(33, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 9, xls.AddFormat(fmt));
            xls.SetCellValue(33, 9, new TFormula("=K33*$I$48"));

            fmt = xls.GetCellVisibleFormatDef(33, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 10, xls.AddFormat(fmt));
            xls.SetCellValue(33, 10, new TFormula("=K33*$J$48"));

            fmt = xls.GetCellVisibleFormatDef(33, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(33, 11, xls.AddFormat(fmt));
            xls.SetCellValue(33, 11, new TFormula("=Budget_Supuestos!B341"));

            fmt = xls.GetCellVisibleFormatDef(34, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 1, xls.AddFormat(fmt));
            xls.SetCellValue(34, 1, "Total materiales cosecha");

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));
            xls.SetCellValue(34, 4, new TFormula("=SUM(D31:D32)"));

            fmt = xls.GetCellVisibleFormatDef(34, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(34, 5, xls.AddFormat(fmt));
            xls.SetCellValue(34, 5, new TFormula("=SUM(E31:E32)"));

            fmt = xls.GetCellVisibleFormatDef(34, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(34, 6, xls.AddFormat(fmt));
            xls.SetCellValue(34, 6, new TFormula("=SUM(F31:F32)"));

            fmt = xls.GetCellVisibleFormatDef(34, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(34, 7, xls.AddFormat(fmt));
            xls.SetCellValue(34, 7, new TFormula("=SUM(G31:G32)"));

            fmt = xls.GetCellVisibleFormatDef(34, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(34, 8, xls.AddFormat(fmt));
            xls.SetCellValue(34, 8, new TFormula("=SUM(H31:H32)"));

            fmt = xls.GetCellVisibleFormatDef(34, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(34, 9, xls.AddFormat(fmt));
            xls.SetCellValue(34, 9, new TFormula("=SUM(I31:I32)"));

            fmt = xls.GetCellVisibleFormatDef(34, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(34, 10, xls.AddFormat(fmt));
            xls.SetCellValue(34, 10, new TFormula("=SUM(J31:J32)"));

            fmt = xls.GetCellVisibleFormatDef(35, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(35, 1, xls.AddFormat(fmt));
            xls.SetCellValue(35, 1, "Total costos variables cosecha");

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));
            xls.SetCellValue(35, 4, new TFormula("=D34+D29"));

            fmt = xls.GetCellVisibleFormatDef(35, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(35, 5, xls.AddFormat(fmt));
            xls.SetCellValue(35, 5, new TFormula("=E34+E29"));

            fmt = xls.GetCellVisibleFormatDef(35, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(35, 6, xls.AddFormat(fmt));
            xls.SetCellValue(35, 6, new TFormula("=F34+F29"));

            fmt = xls.GetCellVisibleFormatDef(35, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(35, 7, xls.AddFormat(fmt));
            xls.SetCellValue(35, 7, new TFormula("=G34+G29"));

            fmt = xls.GetCellVisibleFormatDef(35, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(35, 8, xls.AddFormat(fmt));
            xls.SetCellValue(35, 8, new TFormula("=H34+H29"));

            fmt = xls.GetCellVisibleFormatDef(35, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(35, 9, xls.AddFormat(fmt));
            xls.SetCellValue(35, 9, new TFormula("=I34+I29"));

            fmt = xls.GetCellVisibleFormatDef(35, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(35, 10, xls.AddFormat(fmt));
            xls.SetCellValue(35, 10, new TFormula("=J34+J29"));

            fmt = xls.GetCellVisibleFormatDef(36, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(36, 1, xls.AddFormat(fmt));
            xls.SetCellValue(36, 1, "Beneficio");

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));
            xls.SetCellValue(36, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));
            xls.SetCellValue(36, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));
            xls.SetCellValue(36, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(36, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 5, xls.AddFormat(fmt));
            xls.SetCellValue(36, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(36, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 6, xls.AddFormat(fmt));
            xls.SetCellValue(36, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(36, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 7, xls.AddFormat(fmt));
            xls.SetCellValue(36, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(36, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 8, xls.AddFormat(fmt));
            xls.SetCellValue(36, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(36, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 9, xls.AddFormat(fmt));
            xls.SetCellValue(36, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(36, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 10, xls.AddFormat(fmt));
            xls.SetCellValue(36, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(37, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(37, 1, xls.AddFormat(fmt));
            xls.SetCellValue(37, 1, "Mano de obra beneficio humedo");

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));
            xls.SetCellValue(37, 4, new TFormula("='Budget_Valor de M Obra'!D69"));

            fmt = xls.GetCellVisibleFormatDef(37, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 5, xls.AddFormat(fmt));
            xls.SetCellValue(37, 5, new TFormula("='Budget_Valor de M Obra'!E69"));

            fmt = xls.GetCellVisibleFormatDef(37, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 6, xls.AddFormat(fmt));
            xls.SetCellValue(37, 6, new TFormula("='Budget_Valor de M Obra'!F69"));

            fmt = xls.GetCellVisibleFormatDef(37, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 7, xls.AddFormat(fmt));
            xls.SetCellValue(37, 7, new TFormula("='Budget_Valor de M Obra'!G69"));

            fmt = xls.GetCellVisibleFormatDef(37, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 8, xls.AddFormat(fmt));
            xls.SetCellValue(37, 8, new TFormula("='Budget_Valor de M Obra'!H69"));

            fmt = xls.GetCellVisibleFormatDef(37, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 9, xls.AddFormat(fmt));
            xls.SetCellValue(37, 9, new TFormula("='Budget_Valor de M Obra'!I69"));

            fmt = xls.GetCellVisibleFormatDef(37, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(37, 10, xls.AddFormat(fmt));
            xls.SetCellValue(37, 10, new TFormula("='Budget_Valor de M Obra'!J69"));

            fmt = xls.GetCellVisibleFormatDef(38, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(38, 1, xls.AddFormat(fmt));
            xls.SetCellValue(38, 1, "Mano de obra beneficio seco");

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));
            xls.SetCellValue(38, 4, new TFormula("='Budget_Valor de M Obra'!D78"));

            fmt = xls.GetCellVisibleFormatDef(38, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 5, xls.AddFormat(fmt));
            xls.SetCellValue(38, 5, new TFormula("='Budget_Valor de M Obra'!E78"));

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));
            xls.SetCellValue(38, 6, new TFormula("='Budget_Valor de M Obra'!F78"));

            fmt = xls.GetCellVisibleFormatDef(38, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 7, xls.AddFormat(fmt));
            xls.SetCellValue(38, 7, new TFormula("='Budget_Valor de M Obra'!G78"));

            fmt = xls.GetCellVisibleFormatDef(38, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 8, xls.AddFormat(fmt));
            xls.SetCellValue(38, 8, new TFormula("='Budget_Valor de M Obra'!H78"));

            fmt = xls.GetCellVisibleFormatDef(38, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 9, xls.AddFormat(fmt));
            xls.SetCellValue(38, 9, new TFormula("='Budget_Valor de M Obra'!I78"));

            fmt = xls.GetCellVisibleFormatDef(38, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 10, xls.AddFormat(fmt));
            xls.SetCellValue(38, 10, new TFormula("='Budget_Valor de M Obra'!J78"));

            fmt = xls.GetCellVisibleFormatDef(39, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 1, xls.AddFormat(fmt));
            xls.SetCellValue(39, 1, "Total costos transporte cosecha/pergamino");

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));
            xls.SetCellValue(39, 4, new TFormula("=$K$39*D48"));

            fmt = xls.GetCellVisibleFormatDef(39, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 5, xls.AddFormat(fmt));
            xls.SetCellValue(39, 5, new TFormula("=$K$39*E48"));

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));
            xls.SetCellValue(39, 6, new TFormula("=$K$39*F48"));

            fmt = xls.GetCellVisibleFormatDef(39, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 7, xls.AddFormat(fmt));
            xls.SetCellValue(39, 7, new TFormula("=$K$39*G48"));

            fmt = xls.GetCellVisibleFormatDef(39, 8);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 8, xls.AddFormat(fmt));
            xls.SetCellValue(39, 8, new TFormula("=$K$39*H48"));

            fmt = xls.GetCellVisibleFormatDef(39, 9);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 9, xls.AddFormat(fmt));
            xls.SetCellValue(39, 9, new TFormula("=$K$39*I48"));

            fmt = xls.GetCellVisibleFormatDef(39, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(39, 10, xls.AddFormat(fmt));
            xls.SetCellValue(39, 10, new TFormula("=$K$39*J48"));

            fmt = xls.GetCellVisibleFormatDef(39, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(39, 11, xls.AddFormat(fmt));
            xls.SetCellValue(39, 11, new TFormula("=F58"));

            fmt = xls.GetCellVisibleFormatDef(40, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(40, 1, xls.AddFormat(fmt));
            xls.SetCellValue(40, 1, "Total costos variables beneficio");

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));
            xls.SetCellValue(40, 4, new TFormula("=D37+D38+D39"));

            fmt = xls.GetCellVisibleFormatDef(40, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 5, xls.AddFormat(fmt));
            xls.SetCellValue(40, 5, new TFormula("=E37+E38+E39"));

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));
            xls.SetCellValue(40, 6, new TFormula("=F37+F38+F39"));

            fmt = xls.GetCellVisibleFormatDef(40, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 7, xls.AddFormat(fmt));
            xls.SetCellValue(40, 7, new TFormula("=G37+G38+G39"));

            fmt = xls.GetCellVisibleFormatDef(40, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 8, xls.AddFormat(fmt));
            xls.SetCellValue(40, 8, new TFormula("=H37+H38+H39"));

            fmt = xls.GetCellVisibleFormatDef(40, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 9, xls.AddFormat(fmt));
            xls.SetCellValue(40, 9, new TFormula("=I37+I38+I39"));

            fmt = xls.GetCellVisibleFormatDef(40, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 10, xls.AddFormat(fmt));
            xls.SetCellValue(40, 10, new TFormula("=J37+J38+J39"));

            fmt = xls.GetCellVisibleFormatDef(41, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(41, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(42, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(43, 1, xls.AddFormat(fmt));
            xls.SetCellValue(43, 1, "Cuadro auxiliar de ajuste");

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));
            xls.SetCellValue(44, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));
            xls.SetCellValue(44, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));
            xls.SetCellValue(44, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(44, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 5, xls.AddFormat(fmt));
            xls.SetCellValue(44, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(44, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 6, xls.AddFormat(fmt));
            xls.SetCellValue(44, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(44, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 7, xls.AddFormat(fmt));
            xls.SetCellValue(44, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(44, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 8, xls.AddFormat(fmt));
            xls.SetCellValue(44, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(44, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 9, xls.AddFormat(fmt));
            xls.SetCellValue(44, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(44, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 10, xls.AddFormat(fmt));
            xls.SetCellValue(44, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(45, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(45, 1, xls.AddFormat(fmt));
            xls.SetCellValue(45, 1, "Cosecha");

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));
            xls.SetCellValue(45, 2, new TFormula("=Budget_Supuestos!L145"));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));
            xls.SetCellValue(45, 3, new TFormula("=Budget_Supuestos!L146"));

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));
            xls.SetCellValue(45, 4, new TFormula("=Budget_Supuestos!L147"));

            fmt = xls.GetCellVisibleFormatDef(45, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 5, xls.AddFormat(fmt));
            xls.SetCellValue(45, 5, new TFormula("=Budget_Supuestos!L148"));

            fmt = xls.GetCellVisibleFormatDef(45, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 6, xls.AddFormat(fmt));
            xls.SetCellValue(45, 6, new TFormula("=Budget_Supuestos!L149"));

            fmt = xls.GetCellVisibleFormatDef(45, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 7, xls.AddFormat(fmt));
            xls.SetCellValue(45, 7, new TFormula("=Budget_Supuestos!L150"));

            fmt = xls.GetCellVisibleFormatDef(45, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 8, xls.AddFormat(fmt));
            xls.SetCellValue(45, 8, new TFormula("=Budget_Supuestos!L151"));

            fmt = xls.GetCellVisibleFormatDef(45, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 9, xls.AddFormat(fmt));
            xls.SetCellValue(45, 9, new TFormula("=Budget_Supuestos!L152"));

            fmt = xls.GetCellVisibleFormatDef(45, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(45, 10, xls.AddFormat(fmt));
            xls.SetCellValue(45, 10, new TFormula("=Budget_Supuestos!L153"));
            xls.SetCellValue(46, 1, "Crecimiento anual de la cosecha");

            fmt = xls.GetCellVisibleFormatDef(46, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 5, xls.AddFormat(fmt));
            xls.SetCellValue(46, 5, new TFormula("=(E45-D45)/D45"));

            fmt = xls.GetCellVisibleFormatDef(46, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 6, xls.AddFormat(fmt));
            xls.SetCellValue(46, 6, new TFormula("=(F45-E45)/E45"));

            fmt = xls.GetCellVisibleFormatDef(46, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 7, xls.AddFormat(fmt));
            xls.SetCellValue(46, 7, new TFormula("=(G45-F45)/F45"));

            fmt = xls.GetCellVisibleFormatDef(46, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 8, xls.AddFormat(fmt));
            xls.SetCellValue(46, 8, new TFormula("=(H45-G45)/G45"));

            fmt = xls.GetCellVisibleFormatDef(46, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 9, xls.AddFormat(fmt));
            xls.SetCellValue(46, 9, new TFormula("=(I45-H45)/H45"));

            fmt = xls.GetCellVisibleFormatDef(46, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 10, xls.AddFormat(fmt));
            xls.SetCellValue(46, 10, new TFormula("=(J45-I45)/I45"));
            xls.SetCellValue(47, 1, "Promedio");

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));
            xls.SetCellValue(47, 4, new TFormula("=AVERAGE(D45:J45)"));

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));
            xls.SetCellValue(47, 5, new TFormula("=D47"));

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));
            xls.SetCellValue(47, 6, new TFormula("=E47"));

            fmt = xls.GetCellVisibleFormatDef(47, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(47, 7, xls.AddFormat(fmt));
            xls.SetCellValue(47, 7, new TFormula("=F47"));

            fmt = xls.GetCellVisibleFormatDef(47, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(47, 8, xls.AddFormat(fmt));
            xls.SetCellValue(47, 8, new TFormula("=G47"));

            fmt = xls.GetCellVisibleFormatDef(47, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(47, 9, xls.AddFormat(fmt));
            xls.SetCellValue(47, 9, new TFormula("=H47"));

            fmt = xls.GetCellVisibleFormatDef(47, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(47, 10, xls.AddFormat(fmt));
            xls.SetCellValue(47, 10, new TFormula("=I47"));
            xls.SetCellValue(48, 1, "Participacion anual en relación al promedio");
            xls.SetCellValue(48, 4, new TFormula("=D45/D47"));
            xls.SetCellValue(48, 5, new TFormula("=E45/E47"));
            xls.SetCellValue(48, 6, new TFormula("=F45/F47"));
            xls.SetCellValue(48, 7, new TFormula("=G45/G47"));
            xls.SetCellValue(48, 8, new TFormula("=H45/H47"));
            xls.SetCellValue(48, 9, new TFormula("=I45/I47"));
            xls.SetCellValue(48, 10, new TFormula("=J45/J47"));

            fmt = xls.GetCellVisibleFormatDef(50, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(50, 1, xls.AddFormat(fmt));
            xls.SetCellValue(50, 1, "Costos transporte mantenimiento, fertilización y control plagas");

            fmt = xls.GetCellVisibleFormatDef(50, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(50, 6, xls.AddFormat(fmt));
            xls.SetCellValue(50, 6, "Costo en transporte");
            xls.SetCellValue(51, 1, "Transporte equipo y herramientas");

            fmt = xls.GetCellVisibleFormatDef(51, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(51, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(51, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(51, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(51, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(51, 6, xls.AddFormat(fmt));
            xls.SetCellValue(51, 6, new TFormula("=Budget_Supuestos!B369"));

            fmt = xls.GetCellVisibleFormatDef(52, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(52, 1, xls.AddFormat(fmt));
            xls.SetCellValue(52, 1, "Transporte mano de obra (no pagada en el jornal)");

            fmt = xls.GetCellVisibleFormatDef(52, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(52, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(52, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(52, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(52, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(52, 6, xls.AddFormat(fmt));
            xls.SetCellValue(52, 6, new TFormula("=Budget_Supuestos!B370"));

            fmt = xls.GetCellVisibleFormatDef(53, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(53, 1, xls.AddFormat(fmt));
            xls.SetCellValue(53, 1, "Transporte para ir a supervisas actividades (limpias, manejos, podas, obras conservación)");

            fmt = xls.GetCellVisibleFormatDef(53, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(53, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(53, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(53, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(53, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(53, 6, xls.AddFormat(fmt));
            xls.SetCellValue(53, 6, new TFormula("=Budget_Supuestos!B372"));

            fmt = xls.GetCellVisibleFormatDef(54, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(54, 1, xls.AddFormat(fmt));
            xls.SetCellValue(54, 1, "Otro(s) transportes no considerados:");

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(54, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(54, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(54, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(54, 6, xls.AddFormat(fmt));
            xls.SetCellValue(54, 6, new TFormula("=Budget_Supuestos!B373"));

            fmt = xls.GetCellVisibleFormatDef(55, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(55, 1, xls.AddFormat(fmt));
            xls.SetCellValue(55, 1, "Total costos transporte mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(55, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(55, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(55, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(55, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(55, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            xls.SetCellFormat(55, 6, xls.AddFormat(fmt));
            xls.SetCellValue(55, 6, new TFormula("=SUM(F51:F52)"));

            fmt = xls.GetCellVisibleFormatDef(56, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(56, 1, xls.AddFormat(fmt));
            xls.SetCellValue(56, 1, "Costos transporte cosecha");

            fmt = xls.GetCellVisibleFormatDef(57, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(57, 1, xls.AddFormat(fmt));
            xls.SetCellValue(57, 1, "Cosecha");

            fmt = xls.GetCellVisibleFormatDef(57, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 2, xls.AddFormat(fmt));
            xls.SetCellValue(57, 2, new TFormula("=Budget_Supuestos!#REF!"));

            fmt = xls.GetCellVisibleFormatDef(57, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 3, xls.AddFormat(fmt));
            xls.SetCellValue(57, 3, new TFormula("=Budget_Supuestos!C371"));

            fmt = xls.GetCellVisibleFormatDef(57, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 4, xls.AddFormat(fmt));
            xls.SetCellValue(57, 4, new TFormula("=Budget_Supuestos!D371"));

            fmt = xls.GetCellVisibleFormatDef(57, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 5, xls.AddFormat(fmt));
            xls.SetCellValue(57, 5, new TFormula("=Budget_Supuestos!E371"));

            fmt = xls.GetCellVisibleFormatDef(57, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 6, xls.AddFormat(fmt));
            xls.SetCellValue(57, 6, new TFormula("=Budget_Supuestos!B371"));

            fmt = xls.GetCellVisibleFormatDef(58, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(58, 1, xls.AddFormat(fmt));
            xls.SetCellValue(58, 1, "Total costos transporte cosecha");

            fmt = xls.GetCellVisibleFormatDef(58, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(58, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(58, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(58, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(58, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            xls.SetCellFormat(58, 6, xls.AddFormat(fmt));
            xls.SetCellValue(58, 6, new TFormula("=SUM(F57)"));

            //Cell selection and scroll position.
            xls.SelectCell(23, 4, false);

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
