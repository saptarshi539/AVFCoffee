using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;
namespace CoffeeInfrastructure.Flexcel
{
    public class Budget_Establecimiento
    {

        public void BudgetEstablecimiento(ExcelFile xls)
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

            xls.ActiveSheet = 14;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsCheckCompatibility = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Budget_Establecimiento";
            xls.SheetZoom = 64;
            xls.SheetView = new TSheetView(TSheetViewType.Normal, true, true, 64, 64, 0);

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
            xls.PrintXResolution = 600;
            xls.PrintYResolution = 600;
            xls.PrintOptions = TPrintOptions.Orientation;
            xls.PrintPaperSize = TPaperSize.Letter;

            //Theme - You might use GetTheme/SetTheme methods here instead.
            xls.SetColorTheme(TThemeColor.Background2, TUIColor.FromArgb(0xEE, 0xEC, 0xE1));
            xls.SetColorTheme(TThemeColor.Foreground2, TUIColor.FromArgb(0x1F, 0x49, 0x7D));
            xls.SetColorTheme(TThemeColor.Accent1, TUIColor.FromArgb(0x4F, 0x81, 0xBD));
            xls.SetColorTheme(TThemeColor.Accent2, TUIColor.FromArgb(0xC0, 0x50, 0x4D));
            xls.SetColorTheme(TThemeColor.Accent3, TUIColor.FromArgb(0x9B, 0xBB, 0x59));
            xls.SetColorTheme(TThemeColor.Accent4, TUIColor.FromArgb(0x80, 0x64, 0xA2));
            xls.SetColorTheme(TThemeColor.Accent5, TUIColor.FromArgb(0x4B, 0xAC, 0xC6));
            xls.SetColorTheme(TThemeColor.Accent6, TUIColor.FromArgb(0xF7, 0x96, 0x46));
            xls.SetColorTheme(TThemeColor.HyperLink, TUIColor.FromArgb(0x00, 0x00, 0xFF));
            xls.SetColorTheme(TThemeColor.FollowedHyperLink, TUIColor.FromArgb(0x80, 0x00, 0x80));

            //Major font
            TThemeTextFont MajorLatin = new TThemeTextFont("Cambria", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            TThemeTextFont MajorEastAsian = new TThemeTextFont("", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            TThemeTextFont MajorComplexScript = new TThemeTextFont("", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            TThemeFont MajorFont = new TThemeFont(MajorLatin, MajorEastAsian, MajorComplexScript);
            MajorFont.AddFont("Jpan", "ＭＳ Ｐゴシック");
            MajorFont.AddFont("Hang", "맑은 고딕");
            MajorFont.AddFont("Hans", "宋体");
            MajorFont.AddFont("Hant", "新細明體");
            MajorFont.AddFont("Arab", "Times New Roman");
            MajorFont.AddFont("Hebr", "Times New Roman");
            MajorFont.AddFont("Thai", "Tahoma");
            MajorFont.AddFont("Ethi", "Nyala");
            MajorFont.AddFont("Beng", "Vrinda");
            MajorFont.AddFont("Gujr", "Shruti");
            MajorFont.AddFont("Khmr", "MoolBoran");
            MajorFont.AddFont("Knda", "Tunga");
            MajorFont.AddFont("Guru", "Raavi");
            MajorFont.AddFont("Cans", "Euphemia");
            MajorFont.AddFont("Cher", "Plantagenet Cherokee");
            MajorFont.AddFont("Yiii", "Microsoft Yi Baiti");
            MajorFont.AddFont("Tibt", "Microsoft Himalaya");
            MajorFont.AddFont("Thaa", "MV Boli");
            MajorFont.AddFont("Deva", "Mangal");
            MajorFont.AddFont("Telu", "Gautami");
            MajorFont.AddFont("Taml", "Latha");
            MajorFont.AddFont("Syrc", "Estrangelo Edessa");
            MajorFont.AddFont("Orya", "Kalinga");
            MajorFont.AddFont("Mlym", "Kartika");
            MajorFont.AddFont("Laoo", "DokChampa");
            MajorFont.AddFont("Sinh", "Iskoola Pota");
            MajorFont.AddFont("Mong", "Mongolian Baiti");
            MajorFont.AddFont("Viet", "Times New Roman");
            MajorFont.AddFont("Uigh", "Microsoft Uighur");
            MajorFont.AddFont("Geor", "Sylfaen");
            xls.SetThemeFont(TFontScheme.Major, MajorFont);

            //Minor font
            TThemeTextFont MinorLatin = new TThemeTextFont("Calibri", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            TThemeTextFont MinorEastAsian = new TThemeTextFont("", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            TThemeTextFont MinorComplexScript = new TThemeTextFont("", null, TPitchFamily.DEFAULT_PITCH__UNKNOWN_FONT_FAMILY, TFontCharSet.Default);
            TThemeFont MinorFont = new TThemeFont(MinorLatin, MinorEastAsian, MinorComplexScript);
            MinorFont.AddFont("Jpan", "ＭＳ Ｐゴシック");
            MinorFont.AddFont("Hang", "맑은 고딕");
            MinorFont.AddFont("Hans", "宋体");
            MinorFont.AddFont("Hant", "新細明體");
            MinorFont.AddFont("Arab", "Arial");
            MinorFont.AddFont("Hebr", "Arial");
            MinorFont.AddFont("Thai", "Tahoma");
            MinorFont.AddFont("Ethi", "Nyala");
            MinorFont.AddFont("Beng", "Vrinda");
            MinorFont.AddFont("Gujr", "Shruti");
            MinorFont.AddFont("Khmr", "DaunPenh");
            MinorFont.AddFont("Knda", "Tunga");
            MinorFont.AddFont("Guru", "Raavi");
            MinorFont.AddFont("Cans", "Euphemia");
            MinorFont.AddFont("Cher", "Plantagenet Cherokee");
            MinorFont.AddFont("Yiii", "Microsoft Yi Baiti");
            MinorFont.AddFont("Tibt", "Microsoft Himalaya");
            MinorFont.AddFont("Thaa", "MV Boli");
            MinorFont.AddFont("Deva", "Mangal");
            MinorFont.AddFont("Telu", "Gautami");
            MinorFont.AddFont("Taml", "Latha");
            MinorFont.AddFont("Syrc", "Estrangelo Edessa");
            MinorFont.AddFont("Orya", "Kalinga");
            MinorFont.AddFont("Mlym", "Kartika");
            MinorFont.AddFont("Laoo", "DokChampa");
            MinorFont.AddFont("Sinh", "Iskoola Pota");
            MinorFont.AddFont("Mong", "Mongolian Baiti");
            MinorFont.AddFont("Viet", "Arial");
            MinorFont.AddFont("Uigh", "Microsoft Uighur");
            MinorFont.AddFont("Geor", "Sylfaen");
            xls.SetThemeFont(TFontScheme.Minor, MinorFont);

            //Set up rows and columns
            xls.DefaultColWidth = 0;

            xls.SetColWidth(1, 1, 12320);    //(47.38 + 0.75) * 256

            xls.SetColWidth(2, 2, 3584);    //(13.25 + 0.75) * 256

            xls.SetColWidth(3, 3, 6016);    //(22.75 + 0.75) * 256

            xls.SetColWidth(4, 5, 2816);    //(10.25 + 0.75) * 256

            xls.SetColWidth(6, 6, 4320);    //(16.13 + 0.75) * 256

            xls.SetColWidth(7, 13, 2816);    //(10.25 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(3, 375);    //18.75 * 20
            xls.SetRowHeight(15, 375);    //18.75 * 20
            xls.SetRowHeight(16, 375);    //18.75 * 20
            xls.SetRowHeight(34, 375);    //18.75 * 20
            xls.SetRowHeight(36, 375);    //18.75 * 20

            TFlxFormat RowFmt;
            RowFmt = xls.GetFormat(xls.GetRowFormat(36));
            RowFmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetRowFormat(36, xls.AddFormat(RowFmt));
            xls.SetRowHeight(50, 630);    //31.50 * 20
            xls.SetRowHeight(51, 375);    //18.75 * 20
            xls.SetRowHeight(55, 375);    //18.75 * 20
            xls.SetRowHeight(58, 630);    //31.50 * 20
            xls.SetRowHeight(69, 375);    //18.75 * 20
            xls.SetRowHeight(95, 615);    //30.75 * 20

            //Merged Cells
            xls.MergeCells(80, 1, 80, 5);
            xls.MergeCells(85, 1, 85, 5);
            xls.MergeCells(91, 1, 91, 5);
            xls.MergeCells(97, 1, 97, 5);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(1, 1, xls.AddFormat(fmt));
            xls.SetCellValue(1, 1, "Cuadro. Establecimiento. Costos variables detallados");

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
            xls.SetCellValue(2, 1, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));
            xls.SetCellValue(3, 1, "Germinador");

            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));
            xls.SetCellValue(3, 2, "Precio");

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));
            xls.SetCellValue(3, 3, "Notas");

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));
            xls.SetCellValue(3, 4, "Preguntas control");

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 1, xls.AddFormat(fmt));
            xls.SetCellValue(4, 1, "Mano de obra germinador");

            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));
            xls.SetCellValue(4, 2, new TFormula("='Budget_Valor de M Obra'!$B$11"));

            fmt = xls.GetCellVisibleFormatDef(5, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(5, 1, xls.AddFormat(fmt));
            xls.SetCellValue(5, 1, "Materiales germinador:");

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(6, 1, xls.AddFormat(fmt));
            xls.SetCellValue(6, 1, new TFormula("=Budget_Supuestos!A232"));

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, new TFormula("=Budget_Supuestos!B232"));
            xls.SetCellValue(6, 3, "2 kilos semilla por hectaria");

            fmt = xls.GetCellVisibleFormatDef(7, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(7, 1, xls.AddFormat(fmt));
            xls.SetCellValue(7, 1, new TFormula("=Budget_Supuestos!A233"));

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, new TFormula("=Budget_Supuestos!B233"));
            xls.SetCellValue(7, 3, "1 m x 2 m x 40 cm");
            xls.SetCellValue(7, 4, "Mano obra construccion germinador");

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));
            xls.SetCellValue(7, 8, new TFormula("='Budget_Valor de M Obra'!B8+'Budget_Valor de M Obra'!B9"));
            xls.SetCellValue(7, 12, new TFormula("=H7"));

            fmt = xls.GetCellVisibleFormatDef(8, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(8, 1, xls.AddFormat(fmt));
            xls.SetCellValue(8, 1, new TFormula("=Budget_Supuestos!A234"));

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));
            xls.SetCellValue(8, 2, new TFormula("=Budget_Supuestos!B234"));
            xls.SetCellValue(8, 3, "2 cm2");
            xls.SetCellValue(8, 4, "Materiales para estructura germinador");
            xls.SetCellValue(8, 8, new TFormula("=B7+B8+B9+B10+B11+B6"));
            xls.SetCellValue(8, 12, new TFormula("=H8-B6"));

            fmt = xls.GetCellVisibleFormatDef(9, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(9, 1, xls.AddFormat(fmt));
            xls.SetCellValue(9, 1, new TFormula("=Budget_Supuestos!A235"));

            fmt = xls.GetCellVisibleFormatDef(9, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));
            xls.SetCellValue(9, 2, new TFormula("=Budget_Supuestos!B235"));
            xls.SetCellValue(9, 3, "1 m3 de arena");

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));
            xls.SetCellValue(9, 8, new TFormula("=SUM(H7:H8)"));
            xls.SetCellValue(9, 12, new TFormula("=SUM(L7:L8)"));

            fmt = xls.GetCellVisibleFormatDef(10, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(10, 1, xls.AddFormat(fmt));
            xls.SetCellValue(10, 1, new TFormula("=Budget_Supuestos!A236"));

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));
            xls.SetCellValue(10, 2, new TFormula("=Budget_Supuestos!B236"));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));
            xls.SetCellValue(10, 10, 1);
            xls.SetCellValue(10, 12, 1);

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));
            xls.SetCellValue(11, 1, new TFormula("=Budget_Supuestos!A237"));

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));
            xls.SetCellValue(11, 2, new TFormula("=Budget_Supuestos!B237"));

            fmt = xls.GetCellVisibleFormatDef(12, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(12, 1, xls.AddFormat(fmt));
            xls.SetCellValue(12, 1, new TFormula("=Budget_Supuestos!A238"));

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));
            xls.SetCellValue(12, 2, new TFormula("=Budget_Supuestos!B238"));

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 1, xls.AddFormat(fmt));
            xls.SetCellValue(13, 1, "Total materiales germinador");

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));
            xls.SetCellValue(13, 2, new TFormula("=SUM(B6:B12)"));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, "Valor de mercado que Ud. Le pondria al germinador?");

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));
            xls.SetCellValue(13, 8, 0);

            fmt = xls.GetCellVisibleFormatDef(14, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 1, xls.AddFormat(fmt));
            xls.SetCellValue(14, 1, "Total costos transporte germinador");

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));
            xls.SetCellValue(14, 2, new TFormula("=F80"));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 1, xls.AddFormat(fmt));
            xls.SetCellValue(15, 1, "Total costos variables germinador");

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));
            xls.SetCellValue(15, 2, new TFormula("=B13+B4+B14"));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(16, 1, xls.AddFormat(fmt));
            xls.SetCellValue(16, 1, "Vivero");

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(17, 1, xls.AddFormat(fmt));
            xls.SetCellValue(17, 1, "Mano de obra vivero");

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, new TFormula("='Budget_Valor de M Obra'!$B$23"));

            fmt = xls.GetCellVisibleFormatDef(18, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(18, 1, xls.AddFormat(fmt));
            xls.SetCellValue(18, 1, "Materiales vivero:");

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(19, 1, xls.AddFormat(fmt));
            xls.SetCellValue(19, 1, new TFormula("=Budget_Supuestos!A241"));

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.249977111117893);
            fmt.Format = "#,##0";
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));
            xls.SetCellValue(19, 2, new TFormula("=Budget_Supuestos!B241"));

            fmt = xls.GetCellVisibleFormatDef(20, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(20, 1, xls.AddFormat(fmt));
            xls.SetCellValue(20, 1, new TFormula("=Budget_Supuestos!A242"));

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.249977111117893);
            fmt.Format = "#,##0";
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));
            xls.SetCellValue(20, 2, new TFormula("=Budget_Supuestos!B242"));
            xls.SetCellValue(20, 3, "6000 bolsitas");
            xls.SetCellValue(20, 4, "Mano obra construccion vivero");
            xls.SetCellValue(20, 8, new TFormula("='Budget_Valor de M Obra'!B13"));
            xls.SetCellValue(20, 12, new TFormula("=H20"));

            fmt = xls.GetCellVisibleFormatDef(21, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(21, 1, xls.AddFormat(fmt));
            xls.SetCellValue(21, 1, new TFormula("=Budget_Supuestos!A243"));

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, new TFormula("=Budget_Supuestos!B243"));
            xls.SetCellValue(21, 3, "4 m x 15 m ");
            xls.SetCellValue(21, 4, "Materiales para estructura vivero");
            xls.SetCellValue(21, 8, new TFormula("=B21+B22+B23+B24+B25+B19+B20"));
            xls.SetCellValue(21, 12, new TFormula("=H21-B19-B20"));

            fmt = xls.GetCellVisibleFormatDef(22, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(22, 1, xls.AddFormat(fmt));
            xls.SetCellValue(22, 1, new TFormula("=Budget_Supuestos!A244"));

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));
            xls.SetCellValue(22, 2, new TFormula("=Budget_Supuestos!B244"));
            xls.SetCellValue(22, 3, "20 postes");
            xls.SetCellValue(22, 8, new TFormula("=SUM(H20:H21)"));
            xls.SetCellValue(22, 12, new TFormula("=SUM(L20:L21)"));

            fmt = xls.GetCellVisibleFormatDef(23, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(23, 1, xls.AddFormat(fmt));
            xls.SetCellValue(23, 1, new TFormula("=Budget_Supuestos!A245"));

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));
            xls.SetCellValue(23, 2, new TFormula("=Budget_Supuestos!B245"));
            xls.SetCellValue(23, 8, ".");
            xls.SetCellValue(23, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(24, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(24, 1, xls.AddFormat(fmt));
            xls.SetCellValue(24, 1, new TFormula("=Budget_Supuestos!A246"));

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));
            xls.SetCellValue(24, 2, new TFormula("=Budget_Supuestos!B246"));

            fmt = xls.GetCellVisibleFormatDef(25, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(25, 1, xls.AddFormat(fmt));
            xls.SetCellValue(25, 1, new TFormula("=Budget_Supuestos!A247"));

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));
            xls.SetCellValue(25, 2, new TFormula("=Budget_Supuestos!B247"));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Font.Style = TFlxFontStyles.StrikeOut;
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(26, 1, xls.AddFormat(fmt));
            xls.SetCellValue(26, 1, new TFormula("=Budget_Supuestos!A248"));

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));
            xls.SetCellValue(26, 2, new TFormula("=Budget_Supuestos!B248"));
            xls.SetCellValue(26, 3, "90 kilos");

            fmt = xls.GetCellVisibleFormatDef(26, 4);
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
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));

            TRTFRun[] Runs;
            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 23;
            TFlxFont fnt;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 39;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(26, 4, new TRichString("Valor estimado vivero (de la estructura)", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(26, 4, "Valor estimado vivero (<font color = 'blue'>de la estructura</font>)")


            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));
            xls.SetCellValue(26, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(27, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(27, 1, xls.AddFormat(fmt));
            xls.SetCellValue(27, 1, new TFormula("=Budget_Supuestos!A249"));

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));
            xls.SetCellValue(27, 2, new TFormula("=Budget_Supuestos!B249"));

            fmt = xls.GetCellVisibleFormatDef(28, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(28, 1, xls.AddFormat(fmt));
            xls.SetCellValue(28, 1, new TFormula("=Budget_Supuestos!A250"));

            fmt = xls.GetCellVisibleFormatDef(28, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));
            xls.SetCellValue(28, 2, new TFormula("=Budget_Supuestos!B250"));

            fmt = xls.GetCellVisibleFormatDef(29, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(29, 1, xls.AddFormat(fmt));
            xls.SetCellValue(29, 1, new TFormula("=Budget_Supuestos!A251"));

            fmt = xls.GetCellVisibleFormatDef(29, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, new TFormula("=Budget_Supuestos!B251"));

            fmt = xls.GetCellVisibleFormatDef(30, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(30, 1, xls.AddFormat(fmt));
            xls.SetCellValue(30, 1, new TFormula("=Budget_Supuestos!A252"));

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));
            xls.SetCellValue(30, 2, new TFormula("=Budget_Supuestos!B252"));

            fmt = xls.GetCellVisibleFormatDef(31, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(31, 1, xls.AddFormat(fmt));
            xls.SetCellValue(31, 1, new TFormula("=Budget_Supuestos!A253"));

            fmt = xls.GetCellVisibleFormatDef(31, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));
            xls.SetCellValue(31, 2, new TFormula("=Budget_Supuestos!B253"));

            fmt = xls.GetCellVisibleFormatDef(32, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 1, xls.AddFormat(fmt));
            xls.SetCellValue(32, 1, "Total materiales vivero");

            fmt = xls.GetCellVisibleFormatDef(32, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(32, 2, xls.AddFormat(fmt));
            xls.SetCellValue(32, 2, new TFormula("=SUM(B19:B31)"));

            fmt = xls.GetCellVisibleFormatDef(33, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 1, xls.AddFormat(fmt));
            xls.SetCellValue(33, 1, "Total costos transporte vivero");

            fmt = xls.GetCellVisibleFormatDef(33, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, new TFormula("=F85"));

            fmt = xls.GetCellVisibleFormatDef(34, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 1, xls.AddFormat(fmt));
            xls.SetCellValue(34, 1, "Total costos variables vivero");

            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));
            xls.SetCellValue(34, 2, new TFormula("=B32+B17+B33"));

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(36, 1, xls.AddFormat(fmt));
            xls.SetCellValue(36, 1, "Preparación Terreno y Siembra");

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));
            xls.SetCellValue(36, 2, "Precio");

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));
            xls.SetCellValue(36, 3, "Notas");

            fmt = xls.GetCellVisibleFormatDef(37, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(37, 1, xls.AddFormat(fmt));
            xls.SetCellValue(37, 1, "Mano de obra preparacion terreno y siembra");

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));
            xls.SetCellValue(37, 2, new TFormula("='Budget_Valor de M Obra'!$B$36"));

            fmt = xls.GetCellVisibleFormatDef(38, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 1, xls.AddFormat(fmt));
            xls.SetCellValue(38, 1, "Materiales preparcion terreno y siembra:");

            fmt = xls.GetCellVisibleFormatDef(38, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 11);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 12);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(38, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(39, 1, xls.AddFormat(fmt));
            xls.SetCellValue(39, 1, "Abono organico para los Hoyos");

            fmt = xls.GetCellVisibleFormatDef(39, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "#,##0";
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));
            xls.SetCellValue(39, 2, new TFormula("=Budget_Supuestos!B256"));

            fmt = xls.GetCellVisibleFormatDef(40, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(40, 1, xls.AddFormat(fmt));
            xls.SetCellValue(40, 1, "Especificos:");

            fmt = xls.GetCellVisibleFormatDef(40, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));
            xls.SetCellValue(40, 2, new TFormula("=Budget_Supuestos!B257"));

            fmt = xls.GetCellVisibleFormatDef(41, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(41, 1, xls.AddFormat(fmt));
            xls.SetCellValue(41, 1, "Harina de Roca");

            fmt = xls.GetCellVisibleFormatDef(41, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(41, 2, xls.AddFormat(fmt));
            xls.SetCellValue(41, 2, new TFormula("=Budget_Supuestos!B258"));

            fmt = xls.GetCellVisibleFormatDef(42, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(42, 1, xls.AddFormat(fmt));
            xls.SetCellValue(42, 1, "Cascarilla de Café");

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));
            xls.SetCellValue(42, 2, new TFormula("=Budget_Supuestos!B259"));

            fmt = xls.GetCellVisibleFormatDef(43, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(43, 1, xls.AddFormat(fmt));
            xls.SetCellValue(43, 1, "Gallinaza");

            fmt = xls.GetCellVisibleFormatDef(43, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(43, 2, xls.AddFormat(fmt));
            xls.SetCellValue(43, 2, new TFormula("=Budget_Supuestos!B260"));

            fmt = xls.GetCellVisibleFormatDef(44, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 1, xls.AddFormat(fmt));
            xls.SetCellValue(44, 1, "Abono químico para los hoyos");

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));
            xls.SetCellValue(44, 2, new TFormula("=Budget_Supuestos!B261"));

            fmt = xls.GetCellVisibleFormatDef(45, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 1, xls.AddFormat(fmt));
            xls.SetCellValue(45, 1, "Cal");

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));
            xls.SetCellValue(45, 2, new TFormula("=Budget_Supuestos!B262"));

            fmt = xls.GetCellVisibleFormatDef(46, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 1, xls.AddFormat(fmt));
            xls.SetCellValue(46, 1, "Otros elementos para los hoyos: ");

            fmt = xls.GetCellVisibleFormatDef(46, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(46, 2, xls.AddFormat(fmt));
            xls.SetCellValue(46, 2, new TFormula("=Budget_Supuestos!B263"));

            fmt = xls.GetCellVisibleFormatDef(47, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 1, xls.AddFormat(fmt));
            xls.SetCellValue(47, 1, "Total Fertilizacion");

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));
            xls.SetCellValue(47, 2, new TFormula("=Budget_Supuestos!$B$264"));

            fmt = xls.GetCellVisibleFormatDef(48, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(48, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(49, 1, xls.AddFormat(fmt));
            xls.SetCellValue(49, 1, "Total materiales preparacion terreno y siembra");

            fmt = xls.GetCellVisibleFormatDef(49, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Format = "#,##0";
            xls.SetCellFormat(49, 2, xls.AddFormat(fmt));
            xls.SetCellValue(49, 2, new TFormula("=SUM(B39:B46)"));

            fmt = xls.GetCellVisibleFormatDef(50, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(50, 1, xls.AddFormat(fmt));
            xls.SetCellValue(50, 1, "Total costos transporte preparacion terreno y siembra");

            fmt = xls.GetCellVisibleFormatDef(50, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Format = "#,##0";
            xls.SetCellFormat(50, 2, xls.AddFormat(fmt));
            xls.SetCellValue(50, 2, new TFormula("=F91"));

            fmt = xls.GetCellVisibleFormatDef(51, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(51, 1, xls.AddFormat(fmt));
            xls.SetCellValue(51, 1, "Total Preparacion terreno y siembra");

            fmt = xls.GetCellVisibleFormatDef(51, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(51, 2, xls.AddFormat(fmt));
            xls.SetCellValue(51, 2, new TFormula("=B49+B37+B50"));

            fmt = xls.GetCellVisibleFormatDef(52, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(52, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(53, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(54, 1, xls.AddFormat(fmt));
            xls.SetCellValue(54, 1, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(55, 1, xls.AddFormat(fmt));
            xls.SetCellValue(55, 1, "Plantilla o levante");

            fmt = xls.GetCellVisibleFormatDef(55, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(55, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.WrapText = true;
            xls.SetCellFormat(56, 1, xls.AddFormat(fmt));
            xls.SetCellValue(56, 1, "Mano de obra plantilla o levante");

            fmt = xls.GetCellVisibleFormatDef(56, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2, -0.249977111117893);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(56, 2, xls.AddFormat(fmt));
            xls.SetCellValue(56, 2, new TFormula("='Budget_Valor de M Obra'!$C$45"));

            fmt = xls.GetCellVisibleFormatDef(57, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(57, 1, xls.AddFormat(fmt));
            xls.SetCellValue(57, 1, "Materiales plantilla o levante:");

            fmt = xls.GetCellVisibleFormatDef(57, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(57, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(58, 1, xls.AddFormat(fmt));
            xls.SetCellValue(58, 1, "Abono organico para levante (alrededor de la planta)");

            fmt = xls.GetCellVisibleFormatDef(58, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(58, 2, xls.AddFormat(fmt));
            xls.SetCellValue(58, 2, new TFormula("=Budget_Supuestos!B266"));
            xls.SetCellValue(58, 3, "Notas FCC: 5 bolsas de 37 kilos para 1,000 arboles");

            fmt = xls.GetCellVisibleFormatDef(59, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(59, 1, xls.AddFormat(fmt));
            xls.SetCellValue(59, 1, "Especificos:");

            fmt = xls.GetCellVisibleFormatDef(59, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(59, 2, xls.AddFormat(fmt));
            xls.SetCellValue(59, 2, new TFormula("=Budget_Supuestos!B267"));

            fmt = xls.GetCellVisibleFormatDef(60, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(60, 1, xls.AddFormat(fmt));
            xls.SetCellValue(60, 1, "Harina de Roca");

            fmt = xls.GetCellVisibleFormatDef(60, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(60, 2, xls.AddFormat(fmt));
            xls.SetCellValue(60, 2, new TFormula("=Budget_Supuestos!B268"));

            fmt = xls.GetCellVisibleFormatDef(61, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(61, 1, xls.AddFormat(fmt));
            xls.SetCellValue(61, 1, "Cascarilla de Café");

            fmt = xls.GetCellVisibleFormatDef(61, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(61, 2, xls.AddFormat(fmt));
            xls.SetCellValue(61, 2, new TFormula("=Budget_Supuestos!B269"));

            fmt = xls.GetCellVisibleFormatDef(62, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(62, 1, xls.AddFormat(fmt));
            xls.SetCellValue(62, 1, "Gallinaza");

            fmt = xls.GetCellVisibleFormatDef(62, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(62, 2, xls.AddFormat(fmt));
            xls.SetCellValue(62, 2, new TFormula("=Budget_Supuestos!B270"));

            fmt = xls.GetCellVisibleFormatDef(63, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 1, xls.AddFormat(fmt));
            xls.SetCellValue(63, 1, "Abono químico para levante (alrededor de la planta)");

            fmt = xls.GetCellVisibleFormatDef(63, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(63, 2, xls.AddFormat(fmt));
            xls.SetCellValue(63, 2, new TFormula("=Budget_Supuestos!B271"));

            fmt = xls.GetCellVisibleFormatDef(64, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 1, xls.AddFormat(fmt));
            xls.SetCellValue(64, 1, "Insumos para la foliación en la plantilla");

            fmt = xls.GetCellVisibleFormatDef(64, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(64, 2, xls.AddFormat(fmt));
            xls.SetCellValue(64, 2, new TFormula("=Budget_Supuestos!B272"));

            fmt = xls.GetCellVisibleFormatDef(65, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 1, xls.AddFormat(fmt));
            xls.SetCellValue(65, 1, "Otros elementos para siembra y levante:");

            fmt = xls.GetCellVisibleFormatDef(65, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(65, 2, xls.AddFormat(fmt));
            xls.SetCellValue(65, 2, new TFormula("=Budget_Supuestos!B273"));

            fmt = xls.GetCellVisibleFormatDef(66, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 1, xls.AddFormat(fmt));
            xls.SetCellValue(66, 1, "Total Fertilizaciones");

            fmt = xls.GetCellVisibleFormatDef(66, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(66, 2, xls.AddFormat(fmt));
            xls.SetCellValue(66, 2, new TFormula("=Budget_Supuestos!$B$274"));

            fmt = xls.GetCellVisibleFormatDef(67, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(67, 1, xls.AddFormat(fmt));
            xls.SetCellValue(67, 1, "Total materiales levante");

            fmt = xls.GetCellVisibleFormatDef(67, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(67, 2, xls.AddFormat(fmt));
            xls.SetCellValue(67, 2, new TFormula("=SUM(B58:B65)"));

            fmt = xls.GetCellVisibleFormatDef(68, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(68, 1, xls.AddFormat(fmt));
            xls.SetCellValue(68, 1, "Total costos levante");

            fmt = xls.GetCellVisibleFormatDef(68, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(68, 2, xls.AddFormat(fmt));
            xls.SetCellValue(68, 2, new TFormula("=F97"));

            fmt = xls.GetCellVisibleFormatDef(69, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(69, 1, xls.AddFormat(fmt));
            xls.SetCellValue(69, 1, "Total levante o plantilla");

            fmt = xls.GetCellVisibleFormatDef(69, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Format = "#,##0";
            xls.SetCellFormat(69, 2, xls.AddFormat(fmt));
            xls.SetCellValue(69, 2, new TFormula("=B67+B56+B68"));

            fmt = xls.GetCellVisibleFormatDef(70, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(70, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(71, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(72, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(74, 1, xls.AddFormat(fmt));
            xls.SetCellValue(74, 1, "Costos de Transporte establecimiento Año 0");
            xls.SetCellValue(74, 2, "Tiempo en dias");
            xls.SetCellValue(74, 3, "Costo transporte");
            xls.SetCellValue(74, 4, "Unidad");
            xls.SetCellValue(74, 5, "Frecuencia");

            fmt = xls.GetCellVisibleFormatDef(74, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(74, 6, xls.AddFormat(fmt));
            xls.SetCellValue(74, 6, "Costo en transporte");

            fmt = xls.GetCellVisibleFormatDef(75, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(75, 1, xls.AddFormat(fmt));
            xls.SetCellValue(75, 1, "Semillero:");
            xls.SetCellValue(76, 1, new TFormula("=Budget_Supuestos!A355"));

            fmt = xls.GetCellVisibleFormatDef(76, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(76, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(76, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(76, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(76, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(76, 6, xls.AddFormat(fmt));
            xls.SetCellValue(76, 6, new TFormula("=Budget_Supuestos!B355"));
            xls.SetCellValue(77, 1, new TFormula("=Budget_Supuestos!A356"));

            fmt = xls.GetCellVisibleFormatDef(77, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(77, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(77, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(77, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(77, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(77, 6, xls.AddFormat(fmt));
            xls.SetCellValue(77, 6, new TFormula("=Budget_Supuestos!B356"));
            xls.SetCellValue(78, 1, new TFormula("=Budget_Supuestos!A357"));

            fmt = xls.GetCellVisibleFormatDef(78, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 6, xls.AddFormat(fmt));
            xls.SetCellValue(78, 6, new TFormula("=Budget_Supuestos!B357"));
            xls.SetCellValue(79, 1, new TFormula("=Budget_Supuestos!A358"));

            fmt = xls.GetCellVisibleFormatDef(79, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 6, xls.AddFormat(fmt));
            xls.SetCellValue(79, 6, new TFormula("=Budget_Supuestos!B358"));

            fmt = xls.GetCellVisibleFormatDef(80, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(80, 1, xls.AddFormat(fmt));
            xls.SetCellValue(80, 1, "Total costos transporte germinador");

            fmt = xls.GetCellVisibleFormatDef(80, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(80, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(80, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(80, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(80, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            xls.SetCellFormat(80, 6, xls.AddFormat(fmt));
            xls.SetCellValue(80, 6, new TFormula("=SUM(F76:F79)"));

            fmt = xls.GetCellVisibleFormatDef(81, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 1, xls.AddFormat(fmt));
            xls.SetCellValue(81, 1, "Vivero:");

            fmt = xls.GetCellVisibleFormatDef(81, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(81, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(81, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(81, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(81, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(81, 6, xls.AddFormat(fmt));
            xls.SetCellValue(82, 1, "Jalada de tierra");

            fmt = xls.GetCellVisibleFormatDef(82, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 6, xls.AddFormat(fmt));
            xls.SetCellValue(82, 6, new TFormula("=Budget_Supuestos!B360"));
            xls.SetCellValue(83, 1, "Ir a comprar bolsas y otros insumos para el vivero");

            fmt = xls.GetCellVisibleFormatDef(83, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 6, xls.AddFormat(fmt));
            xls.SetCellValue(83, 6, new TFormula("=Budget_Supuestos!B361"));
            xls.SetCellValue(84, 1, "Otro(s)");

            fmt = xls.GetCellVisibleFormatDef(84, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 6, xls.AddFormat(fmt));
            xls.SetCellValue(84, 6, new TFormula("=Budget_Supuestos!B362"));

            fmt = xls.GetCellVisibleFormatDef(85, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(85, 1, xls.AddFormat(fmt));
            xls.SetCellValue(85, 1, "Total costos transporte  vivero");

            fmt = xls.GetCellVisibleFormatDef(85, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(85, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(85, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(85, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(85, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            xls.SetCellFormat(85, 6, xls.AddFormat(fmt));
            xls.SetCellValue(85, 6, new TFormula("=SUM(F82:F83)"));

            fmt = xls.GetCellVisibleFormatDef(86, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(86, 1, xls.AddFormat(fmt));
            xls.SetCellValue(86, 1, "Preparación terreno y siembra:");

            fmt = xls.GetCellVisibleFormatDef(86, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(86, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(86, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(86, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(86, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(86, 6, xls.AddFormat(fmt));
            xls.SetCellValue(87, 1, "Llevada de leña");

            fmt = xls.GetCellVisibleFormatDef(87, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 6, xls.AddFormat(fmt));
            xls.SetCellValue(87, 6, new TFormula("=Budget_Supuestos!B364"));
            xls.SetCellValue(88, 1, "Lleva del abono");

            fmt = xls.GetCellVisibleFormatDef(88, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 6, xls.AddFormat(fmt));
            xls.SetCellValue(88, 6, new TFormula("=Budget_Supuestos!B365"));
            xls.SetCellValue(89, 1, "Llevar plantas del vivero al campo");

            fmt = xls.GetCellVisibleFormatDef(89, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(89, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(89, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(89, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(89, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(89, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(89, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(89, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(89, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(89, 6, xls.AddFormat(fmt));
            xls.SetCellValue(89, 6, new TFormula("=Budget_Supuestos!B366"));
            xls.SetCellValue(90, 1, "Otro(s)");

            fmt = xls.GetCellVisibleFormatDef(90, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 6, xls.AddFormat(fmt));
            xls.SetCellValue(90, 6, new TFormula("=Budget_Supuestos!B367"));

            fmt = xls.GetCellVisibleFormatDef(91, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(91, 1, xls.AddFormat(fmt));
            xls.SetCellValue(91, 1, "Total costos preparacion terreno y siembra");

            fmt = xls.GetCellVisibleFormatDef(91, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(91, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(91, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(91, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(91, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            xls.SetCellFormat(91, 6, xls.AddFormat(fmt));
            xls.SetCellValue(91, 6, new TFormula("=SUM(F87:F89)"));

            fmt = xls.GetCellVisibleFormatDef(92, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 1, xls.AddFormat(fmt));
            xls.SetCellValue(92, 1, "Costos transporte levante Año 1");
            xls.SetCellValue(92, 2, "Tiempo en dias");
            xls.SetCellValue(92, 3, "Costo transporte");
            xls.SetCellValue(92, 4, "Unidad");
            xls.SetCellValue(92, 5, "Frecuencia");

            fmt = xls.GetCellVisibleFormatDef(92, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 6, xls.AddFormat(fmt));
            xls.SetCellValue(92, 6, "Costo en transporte");
            xls.SetCellValue(93, 1, "Transporte equipo y herramientas");

            fmt = xls.GetCellVisibleFormatDef(93, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 6, xls.AddFormat(fmt));
            xls.SetCellValue(93, 6, new TFormula("=Budget_Supuestos!B369"));

            fmt = xls.GetCellVisibleFormatDef(94, 1);
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
            xls.SetCellFormat(94, 1, xls.AddFormat(fmt));
            xls.SetCellValue(94, 1, "Transporte mano de obra (no pagada en el jornal)");

            fmt = xls.GetCellVisibleFormatDef(94, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 6, xls.AddFormat(fmt));
            xls.SetCellValue(94, 6, new TFormula("=Budget_Supuestos!B370"));

            fmt = xls.GetCellVisibleFormatDef(95, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(95, 1, xls.AddFormat(fmt));
            xls.SetCellValue(95, 1, "Transporte para ir a supervisas actividades (limpias, manejos, podas, obras conservación)");

            fmt = xls.GetCellVisibleFormatDef(95, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 6, xls.AddFormat(fmt));
            xls.SetCellValue(95, 6, new TFormula("=Budget_Supuestos!B372"));

            fmt = xls.GetCellVisibleFormatDef(96, 1);
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
            xls.SetCellFormat(96, 1, xls.AddFormat(fmt));
            xls.SetCellValue(96, 1, "Otro(s) transportes no considerados:");

            fmt = xls.GetCellVisibleFormatDef(96, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 6, xls.AddFormat(fmt));
            xls.SetCellValue(96, 6, new TFormula("=Budget_Supuestos!B373"));

            fmt = xls.GetCellVisibleFormatDef(97, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(97, 1, xls.AddFormat(fmt));
            xls.SetCellValue(97, 1, "Total costos levante");

            fmt = xls.GetCellVisibleFormatDef(97, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(97, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(97, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(97, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(97, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent2);
            xls.SetCellFormat(97, 6, xls.AddFormat(fmt));
            xls.SetCellValue(97, 6, new TFormula("=SUM(F93:F95)"));

            //Cell selection and scroll position.
            xls.SelectCell(44, 11, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-07T22:31:31Z");


        }

    }
}
