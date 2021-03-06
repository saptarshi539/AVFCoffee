﻿using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;


namespace CoffeeInfrastructure.Flexcel
{
    public class Prporcion_de_productividad
    {
        public void ProporcionDeProductividad(ExcelFile xls)
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

            xls.ActiveSheet = 19;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Proporción de productividad";
            xls.SheetZoom = 86;
            xls.SheetView = new TSheetView(TSheetViewType.Normal, true, true, 86, 86, 0);

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

            xls.SetColWidth(1, 1, 810);    //(2.41 + 0.75) * 256

            TFlxFormat ColFmt;
            ColFmt = xls.GetFormat(xls.GetColFormat(1));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(1, 1, xls.AddFormat(ColFmt));

            xls.SetColWidth(2, 2, 3029);    //(11.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(2));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(2, 2, xls.AddFormat(ColFmt));

            xls.SetColWidth(3, 3, 3754);    //(13.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(3));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(3, 3, xls.AddFormat(ColFmt));

            xls.SetColWidth(4, 4, 3242);    //(11.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(4));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(4, 4, xls.AddFormat(ColFmt));

            xls.SetColWidth(5, 5, 3072);    //(11.25 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(5));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(5, 5, xls.AddFormat(ColFmt));

            xls.SetColWidth(6, 6, 3626);    //(13.41 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(6));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(6, 6, xls.AddFormat(ColFmt));

            xls.SetColWidth(7, 8, 2261);    //(8.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(7));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(7, 8, xls.AddFormat(ColFmt));

            xls.SetColWidth(9, 9, 3968);    //(14.75 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(9));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(9, 9, xls.AddFormat(ColFmt));

            xls.SetColWidth(10, 11, 2261);    //(8.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(10));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(10, 11, xls.AddFormat(ColFmt));

            xls.SetColWidth(12, 12, 5034);    //(18.91 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(12));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(12, 12, xls.AddFormat(ColFmt));

            xls.SetColWidth(13, 14, 2261);    //(8.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(13));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(13, 14, xls.AddFormat(ColFmt));

            xls.SetColWidth(15, 15, 3285);    //(12.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(15));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(15, 15, xls.AddFormat(ColFmt));

            xls.SetColWidth(16, 16384, 2261);    //(8.08 + 0.75) * 256

            ColFmt = xls.GetFormat(xls.GetColFormat(16));
            ColFmt.Font.Size20 = 200;
            xls.SetColFormat(16, 16384, xls.AddFormat(ColFmt));
            xls.DefaultRowHeight = 280;

            xls.SetRowHeight(4, 270);    //13.50 * 20
            xls.SetRowHeight(5, 780);    //39.00 * 20
            xls.SetRowHeight(7, 270);    //13.50 * 20
            xls.SetRowHeight(14, 780);    //39.00 * 20

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(2, 2);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 8);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 9);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 10);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 11);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 12);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 13);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 14);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 15);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 9);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 10);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 11);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 12);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 13);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 14);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 15);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));
            xls.SetCellValue(4, 2, "Productividad");

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));
            xls.SetCellValue(4, 3, "Perú");

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 6, xls.AddFormat(fmt));
            xls.SetCellValue(4, 6, "Mexico CESMACH");

            fmt = xls.GetCellVisibleFormatDef(4, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 8);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 9);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 9, xls.AddFormat(fmt));
            xls.SetCellValue(4, 9, "FCC");

            fmt = xls.GetCellVisibleFormatDef(4, 10);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 11);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 12);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 12, xls.AddFormat(fmt));
            xls.SetCellValue(4, 12, "COMSA");

            fmt = xls.GetCellVisibleFormatDef(4, 13);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 14);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 15);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));
            xls.SetCellValue(5, 3, "Pergamino seco (kilos / hectarea)");

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));
            xls.SetCellValue(5, 4, "cambio %");

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));
            xls.SetCellValue(5, 6, "Pergamino seco (quintales/ha)");

            fmt = xls.GetCellVisibleFormatDef(5, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 7, xls.AddFormat(fmt));
            xls.SetCellValue(5, 7, "cambio %");

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 9);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 9, xls.AddFormat(fmt));
            xls.SetCellValue(5, 9, "Pergamino seco (kilos / Manzana)");

            fmt = xls.GetCellVisibleFormatDef(5, 10);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 10, xls.AddFormat(fmt));
            xls.SetCellValue(5, 10, "cambio %");

            fmt = xls.GetCellVisibleFormatDef(5, 11);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 12);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 12, xls.AddFormat(fmt));
            xls.SetCellValue(5, 12, "UVA \n(Quintales/ Manzana)");

            fmt = xls.GetCellVisibleFormatDef(5, 13);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 13, xls.AddFormat(fmt));
            xls.SetCellValue(5, 13, "cambio %");

            fmt = xls.GetCellVisibleFormatDef(5, 14);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(5, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 15);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 15, xls.AddFormat(fmt));
            xls.SetCellValue(5, 15, "Promedios crecimientos %");

            fmt = xls.GetCellVisibleFormatDef(5, 16);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(5, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, "Años 2,3");

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));
            xls.SetCellValue(6, 3, new TFormula("=AVERAGE(C16:C17)"));

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));
            xls.SetCellValue(6, 4, new TFormula("=(C7-C6)/100"));

            fmt = xls.GetCellVisibleFormatDef(6, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(6, 6, xls.AddFormat(fmt));
            xls.SetCellValue(6, 6, new TFormula("=AVERAGE(D16:D17)"));

            fmt = xls.GetCellVisibleFormatDef(6, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(6, 7, xls.AddFormat(fmt));
            xls.SetCellValue(6, 7, new TFormula("=(F7-F6)/F6"));

            fmt = xls.GetCellVisibleFormatDef(6, 8);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(6, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(6, 9, xls.AddFormat(fmt));
            xls.SetCellValue(6, 9, new TFormula("=AVERAGE(E16:E17)"));

            fmt = xls.GetCellVisibleFormatDef(6, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(6, 10, xls.AddFormat(fmt));
            xls.SetCellValue(6, 10, new TFormula("=(I7-I6)/I6"));

            fmt = xls.GetCellVisibleFormatDef(6, 11);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(6, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(6, 12, xls.AddFormat(fmt));
            xls.SetCellValue(6, 12, new TFormula("=AVERAGE(F16:F17)"));

            fmt = xls.GetCellVisibleFormatDef(6, 13);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(6, 13, xls.AddFormat(fmt));
            xls.SetCellValue(6, 13, new TFormula("=(L7-L6)/L6"));

            fmt = xls.GetCellVisibleFormatDef(6, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(6, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(6, 15, xls.AddFormat(fmt));
            xls.SetCellValue(6, 15, new TFormula("=AVERAGE(D6,G6,J6,M6)"));

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, "Años 4,5,6");

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));
            xls.SetCellValue(7, 3, new TFormula("=AVERAGE(C18:C20)"));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));
            xls.SetCellValue(7, 4, new TFormula("=(C8-C7)/C7"));

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));
            xls.SetCellValue(7, 6, new TFormula("=AVERAGE(D18:D20)"));

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));
            xls.SetCellValue(7, 7, new TFormula("=(F8-F7)/F7"));

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(7, 9, xls.AddFormat(fmt));
            xls.SetCellValue(7, 9, new TFormula("=AVERAGE(E18:E20)"));

            fmt = xls.GetCellVisibleFormatDef(7, 10);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 10, xls.AddFormat(fmt));
            xls.SetCellValue(7, 10, new TFormula("=(I8-I7)/I7"));

            fmt = xls.GetCellVisibleFormatDef(7, 11);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(7, 12, xls.AddFormat(fmt));
            xls.SetCellValue(7, 12, new TFormula("=AVERAGE(F18:F20)"));

            fmt = xls.GetCellVisibleFormatDef(7, 13);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 13, xls.AddFormat(fmt));
            xls.SetCellValue(7, 13, new TFormula("=(L8-L7)/L7"));

            fmt = xls.GetCellVisibleFormatDef(7, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(7, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 15);
            fmt.Font.Size20 = 200;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent2, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.000";
            xls.SetCellFormat(7, 15, xls.AddFormat(fmt));
            xls.SetCellValue(7, 15, new TFormula("=AVERAGE(D7,G7,J7,M7)"));

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));
            xls.SetCellValue(8, 2, "Años 7, 8");

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));
            xls.SetCellValue(8, 3, 2250);

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));
            xls.SetCellValue(8, 6, new TFormula("=AVERAGE(D21:D22)"));

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 8);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(8, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 9, xls.AddFormat(fmt));
            xls.SetCellValue(8, 9, new TFormula("=AVERAGE(E21:E22)"));

            fmt = xls.GetCellVisibleFormatDef(8, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 11);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 12, xls.AddFormat(fmt));
            xls.SetCellValue(8, 12, new TFormula("=AVERAGE(F21:F22)"));

            fmt = xls.GetCellVisibleFormatDef(8, 13);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(8, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 11);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 13);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(9, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(10, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 11);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(10, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(10, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 13);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(10, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(10, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 11);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 13);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(11, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.Font.Size20 = 200;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(12, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 11);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(12, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(12, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 13);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(12, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(12, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
            xls.SetCellValue(13, 3, "Perú");

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, "Mexico CESMACH");

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));
            xls.SetCellValue(13, 5, "FCC");

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));
            xls.SetCellValue(13, 6, "COMSA");

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 10);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 11);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 12);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 13);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 14);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(13, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));
            xls.SetCellValue(14, 2, "Valores originales");

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));
            xls.SetCellValue(14, 3, "Pergamino seco (kilos / hectarea)");

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, "Pergamino seco (quintales/ha)");

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));
            xls.SetCellValue(14, 5, "Pergamino seco (kilos / Manzana)");

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));
            xls.SetCellValue(14, 6, "UVA \n(Quintales/ Manzana)");

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));
            xls.SetCellValue(16, 2, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));
            xls.SetCellValue(16, 3, 500);

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));
            xls.SetCellValue(16, 4, 2.3047619047619);

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));
            xls.SetCellValue(16, 5, 553.771929824561);

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));
            xls.SetCellValue(16, 6, 47.6318681318681);

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));
            xls.SetCellValue(17, 3, 1500);

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));
            xls.SetCellValue(17, 4, 5.5359375);

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));
            xls.SetCellValue(17, 5, 1051.84210526316);

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));
            xls.SetCellValue(17, 6, 91.978947368421);

            fmt = xls.GetCellVisibleFormatDef(17, 12);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            xls.SetCellFormat(17, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));
            xls.SetCellValue(18, 2, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));
            xls.SetCellValue(18, 3, 1750);

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));
            xls.SetCellValue(18, 4, 8.828125);

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));
            xls.SetCellValue(18, 5, 1371.51785714286);

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));
            xls.SetCellValue(18, 6, 136.774193548387);

            fmt = xls.GetCellVisibleFormatDef(18, 12);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            xls.SetCellFormat(18, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));
            xls.SetCellValue(19, 2, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));
            xls.SetCellValue(19, 3, 2500);

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));
            xls.SetCellValue(19, 4, 11.7272727272727);

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));
            xls.SetCellValue(19, 5, 1487.83018867925);

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));
            xls.SetCellValue(19, 6, 161.129032258065);

            fmt = xls.GetCellVisibleFormatDef(19, 12);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            xls.SetCellFormat(19, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));
            xls.SetCellValue(20, 2, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));
            xls.SetCellValue(20, 3, 2500);

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));
            xls.SetCellValue(20, 4, 15.175);

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));
            xls.SetCellValue(20, 5, 1515.44230769231);

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));
            xls.SetCellValue(20, 6, 171.326086956522);

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));
            xls.SetCellValue(21, 3, 2500);

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));
            xls.SetCellValue(21, 4, 17.2121212121212);

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));
            xls.SetCellValue(21, 5, 1221.47169811321);

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));
            xls.SetCellValue(21, 6, 167.989130434783);

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));
            xls.SetCellValue(22, 2, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));
            xls.SetCellValue(22, 3, 2000);

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));
            xls.SetCellValue(22, 4, 19.8939393939394);

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));
            xls.SetCellValue(22, 5, 1017);

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));
            xls.SetCellValue(22, 6, 169.923076923077);

            //Objects
            TShapeProperties ShapeOptions1 = new TShapeProperties();
            ShapeOptions1.Anchor = new TClientAnchor(TFlxAnchorType.MoveAndResize, 10, 68, 10, 145, 26, 164, 16, 942);
            ShapeOptions1.ShapeType = TShapeType.Rectangle;
            ShapeOptions1.ObjectType = TObjectType.MicrosoftOfficeDrawing;
            ShapeOptions1.ShapeName = "TextBox 1";
            ShapeOptions1.Text = "\n\n\n\n\n\n\n\ncon ejemplo del primer cambio porcental (4.15)\n\n\n\n";
            ShapeOptions1.TextFlags = 530;
            ShapeOptions1.RotateTextWithShape = true;
            ShapeOptions1.ShapeThemeFont = new TShapeFont(TFontScheme.Minor, TDrawingColor.FromTheme(TThemeColor.Foreground1));
            ShapeOptions1.Print = true;
            ShapeOptions1.Visible = true;
            ShapeOptions1.ShapeGeometry = "<?xml version=\"1.0\" encoding=\"utf-8\" standalone=\"yes\"?><a:shapeGeom xmlns:a=\"http://schemas.openxmlformats.org/drawingml/2006/main\"><a:prstGeom"
            + " prst=\"rect\"><a:avLst /></a:prstGeom></a:shapeGeom>";
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.fillColor, 16777215);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.fillBackColor, 134217808);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.fFilled, true);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.lineColor, 12369084);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.shadowColor, 0);
            ShapeOptions1.ShapeOptions.SetValue(TShapeOption.wzName, "TextBox 1");
            xls.AddAutoShape(ShapeOptions1);


            //Cell selection and scroll position.
            xls.SelectCell(6, 15, false);

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
