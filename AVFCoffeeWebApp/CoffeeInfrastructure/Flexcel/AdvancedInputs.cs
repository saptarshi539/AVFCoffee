
using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;

namespace CoffeeInfrastructure.Flexcel
{
    public class AdvancedInputs {
        public void Budget_Supuestos(ExcelFile xls)
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

            xls.ActiveSheet = 9;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsCheckCompatibility = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Budget_Supuestos";
            xls.SheetZoom = 62;
            xls.SheetView = new TSheetView(TSheetViewType.Normal, true, true, 62, 62, 0);

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
            xls.PrintScale = 61;
            xls.PrintXResolution = 600;
            xls.PrintYResolution = 600;
            xls.PrintOptions = TPrintOptions.None;
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

            xls.SetColWidth(1, 1, 12704);    //(48.88 + 0.75) * 256

            xls.SetColWidth(2, 2, 6368);    //(24.13 + 0.75) * 256

            xls.SetColWidth(3, 3, 8576);    //(32.75 + 0.75) * 256

            xls.SetColWidth(4, 5, 7296);    //(27.75 + 0.75) * 256

            xls.SetColWidth(6, 6, 12512);    //(48.13 + 0.75) * 256

            xls.SetColWidth(7, 7, 7584);    //(28.88 + 0.75) * 256

            xls.SetColWidth(8, 9, 5792);    //(21.88 + 0.75) * 256

            xls.SetColWidth(10, 10, 7968);    //(30.38 + 0.75) * 256

            xls.SetColWidth(11, 11, 6304);    //(23.88 + 0.75) * 256

            xls.SetColWidth(12, 12, 5664);    //(21.38 + 0.75) * 256

            xls.SetColWidth(13, 13, 7200);    //(27.38 + 0.75) * 256

            xls.SetColWidth(14, 14, 7456);    //(28.38 + 0.75) * 256

            xls.SetColWidth(15, 15, 5632);    //(21.25 + 0.75) * 256

            xls.SetColWidth(16, 21, 2816);    //(10.25 + 0.75) * 256

            xls.SetColWidth(22, 22, 3840);    //(14.25 + 0.75) * 256

            xls.SetColWidth(23, 23, 5664);    //(21.38 + 0.75) * 256

            xls.SetColWidth(24, 24, 7040);    //(26.75 + 0.75) * 256

            xls.SetColWidth(25, 25, 6528);    //(24.75 + 0.75) * 256

            xls.SetColWidth(26, 27, 2816);    //(10.25 + 0.75) * 256

            xls.SetColWidth(28, 16384, 0);
            xls.SetColHidden(28, 16384, true);
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(1, 420);    //21.00 * 20
            xls.SetRowHeight(18, 360);    //18.00 * 20
            xls.SetRowHeight(19, 360);    //18.00 * 20
            xls.SetRowHeight(20, 360);    //18.00 * 20
            xls.SetRowHeight(21, 360);    //18.00 * 20
            xls.SetRowHeight(22, 360);    //18.00 * 20
            xls.SetRowHeight(23, 360);    //18.00 * 20
            xls.SetRowHeight(24, 360);    //18.00 * 20
            xls.SetRowHeight(25, 360);    //18.00 * 20
            xls.SetRowHeight(26, 360);    //18.00 * 20
            xls.SetRowHeight(27, 360);    //18.00 * 20
            xls.SetRowHeight(28, 360);    //18.00 * 20
            xls.SetRowHeight(29, 360);    //18.00 * 20
            xls.SetRowHeight(30, 360);    //18.00 * 20
            xls.SetRowHeight(31, 360);    //18.00 * 20
            xls.SetRowHeight(32, 360);    //18.00 * 20
            xls.SetRowHeight(33, 360);    //18.00 * 20
            xls.SetRowHeight(34, 360);    //18.00 * 20
            xls.SetRowHeight(35, 360);    //18.00 * 20
            xls.SetRowHeight(36, 360);    //18.00 * 20
            xls.SetRowHeight(37, 360);    //18.00 * 20
            xls.SetRowHeight(38, 360);    //18.00 * 20
            xls.SetRowHeight(39, 360);    //18.00 * 20
            xls.SetRowHeight(40, 360);    //18.00 * 20
            xls.SetRowHeight(41, 360);    //18.00 * 20
            xls.SetRowHeight(42, 360);    //18.00 * 20
            xls.SetRowHeight(43, 360);    //18.00 * 20
            xls.SetRowHeight(44, 360);    //18.00 * 20
            xls.SetRowHeight(46, 420);    //21.00 * 20
            xls.SetRowHeight(47, 330);    //16.50 * 20
            xls.SetRowHeight(48, 960);    //48.00 * 20
            xls.SetRowHeight(49, 630);    //31.50 * 20
            xls.SetRowHeight(55, 630);    //31.50 * 20
            xls.SetRowHeight(56, 1260);    //63.00 * 20
            xls.SetRowHeight(62, 1575);    //78.75 * 20
            xls.SetRowHeight(76, 420);    //21.00 * 20
            xls.SetRowHeight(79, 945);    //47.25 * 20
            xls.SetRowHeight(81, 342);    //17.10 * 20
            xls.SetRowHeight(109, 945);    //47.25 * 20
            xls.SetRowHeight(111, 319);    //15.95 * 20
            xls.SetRowHeight(125, 945);    //47.25 * 20
            xls.SetRowHeight(127, 319);    //15.95 * 20
            xls.SetRowHeight(139, 420);    //21.00 * 20
            xls.SetRowHeight(143, 630);    //31.50 * 20
            xls.SetRowHeight(157, 375);    //18.75 * 20
            xls.SetRowHeight(158, 630);    //31.50 * 20
            xls.SetRowHeight(170, 630);    //31.50 * 20
            xls.SetRowHeight(183, 600);    //30.00 * 20
            xls.SetRowHeight(184, 300);    //15.00 * 20
            xls.SetRowHeight(226, 420);    //21.00 * 20
            xls.SetRowHeight(228, 945);    //47.25 * 20
            xls.SetRowHeight(229, 300);    //15.00 * 20
            xls.SetRowHeight(230, 300);    //15.00 * 20
            xls.SetRowHeight(266, 630);    //31.50 * 20
            xls.SetRowHeight(276, 900);    //45.00 * 20
            xls.SetRowHeight(348, 630);    //31.50 * 20
            xls.SetRowHeight(349, 630);    //31.50 * 20
            xls.SetRowHeight(350, 630);    //31.50 * 20
            xls.SetRowHeight(371, 720);    //36.00 * 20
            xls.SetRowHeight(372, 615);    //30.75 * 20
            xls.SetRowHeight(393, 630);    //31.50 * 20
            xls.SetRowHeight(395, 630);    //31.50 * 20
            xls.SetRowHeight(400, 630);    //31.50 * 20
            xls.SetRowHeight(408, 615);    //30.75 * 20
            xls.SetRowHeight(416, 1215);    //60.75 * 20
            xls.SetRowHeight(418, 615);    //30.75 * 20
            xls.SetRowHeight(421, 615);    //30.75 * 20
            xls.SetRowHeight(422, 615);    //30.75 * 20
            xls.SetRowHeight(424, 915);    //45.75 * 20
            xls.SetRowHeight(425, 360);    //18.00 * 20
            xls.SetRowHeight(426, 720);    //36.00 * 20
            xls.SetRowHeight(427, 360);    //18.00 * 20
            xls.SetRowHeight(428, 360);    //18.00 * 20
            xls.SetRowHeight(429, 360);    //18.00 * 20
            xls.SetRowHeight(430, 360);    //18.00 * 20
            xls.SetRowHeight(431, 360);    //18.00 * 20
            xls.SetRowHeight(432, 360);    //18.00 * 20
            xls.SetRowHeight(433, 360);    //18.00 * 20

            //Merged Cells
            xls.MergeCells(230, 2, 231, 2);
            xls.MergeCells(56, 16, 56, 19);
            xls.MergeCells(143, 4, 143, 6);
            xls.MergeCells(194, 4, 194, 6);
            xls.MergeCells(158, 9, 158, 10);
            xls.MergeCells(170, 4, 170, 6);
            xls.MergeCells(229, 6, 231, 6);
            xls.MergeCells(183, 4, 183, 6);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 1);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(1, 1, xls.AddFormat(fmt));
            xls.SetCellValue(1, 1, "Finca Cafetera");

            fmt = xls.GetCellVisibleFormatDef(1, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(1, 3, xls.AddFormat(fmt));
            xls.SetCellValue(1, 3, "Info que ya se metio al informe escrito");

            fmt = xls.GetCellVisibleFormatDef(1, 8);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(1, 8, xls.AddFormat(fmt));
            xls.SetCellValue(1, 8, "Columna links");

            fmt = xls.GetCellVisibleFormatDef(2, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(2, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));
            xls.SetCellValue(3, 1, "Supuestos de la Finca");

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(4, 1, xls.AddFormat(fmt));
            xls.SetCellValue(4, 2, "Hectarea");
            xls.SetCellValue(4, 3, "Manzanas");

            fmt = xls.GetCellVisibleFormatDef(4, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(4, 8, xls.AddFormat(fmt));
            xls.SetCellValue(4, 8, "Hectareas");
            xls.SetCellValue(5, 1, "Tamaño area productiva analizada (ha)");

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
            xls.SetCellValue(5, 2, 1);

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));
            xls.SetCellValue(5, 8, 1);
            xls.SetCellValue(6, 1, "Area de café en total");

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, new TFormula("='Inputs advanced'!F152"));

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 8, xls.AddFormat(fmt));
            xls.SetCellValue(6, 8, new TFormula("=IF(B6=\".\",C6/Conversiones!$C$7,B6)"));
            xls.SetCellValue(7, 1, "Area de la finca en total");

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, new TFormula("='Inputs advanced'!F151"));

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));
            xls.SetCellValue(7, 8, new TFormula("=IF(B7=\".\",C7/Conversiones!$C$7,B7)"));

            fmt = xls.GetCellVisibleFormatDef(8, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(8, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(15, 10, xls.AddFormat(fmt));
            xls.SetCellValue(15, 10, "ANTES ERA 5000 COMO QUE REDONDEABA");

            fmt = xls.GetCellVisibleFormatDef(15, 11);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(15, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 12);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(15, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(15, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(16, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 11);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(16, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 12);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(16, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(16, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 1, xls.AddFormat(fmt));
            xls.SetCellValue(17, 1, "Arboles de café");

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, "Procentaje");

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(18, 1, xls.AddFormat(fmt));
            xls.SetCellValue(18, 1, "Arabe");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));
            xls.SetCellValue(18, 2, new TFormula("='Inputs advanced'!F157"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(18, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(19, 1, xls.AddFormat(fmt));
            xls.SetCellValue(19, 1, "Borbon");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));
            xls.SetCellValue(19, 2, new TFormula("='Inputs advanced'!F158"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(20, 1, xls.AddFormat(fmt));
            xls.SetCellValue(20, 1, "Catimore");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));
            xls.SetCellValue(20, 2, new TFormula("='Inputs advanced'!F159"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(20, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(21, 1, xls.AddFormat(fmt));
            xls.SetCellValue(21, 1, "Catuai");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, new TFormula("='Inputs advanced'!F160"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(22, 1, xls.AddFormat(fmt));
            xls.SetCellValue(22, 1, "Caturra");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));
            xls.SetCellValue(22, 2, new TFormula("='Inputs advanced'!F161"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(23, 1, xls.AddFormat(fmt));
            xls.SetCellValue(23, 1, "Colombia");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));
            xls.SetCellValue(23, 2, new TFormula("='Inputs advanced'!F162"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(24, 1, xls.AddFormat(fmt));
            xls.SetCellValue(24, 1, "Costa Rica");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));
            xls.SetCellValue(24, 2, new TFormula("='Inputs advanced'!F163"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(25, 1, xls.AddFormat(fmt));
            xls.SetCellValue(25, 1, "Castillo");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));
            xls.SetCellValue(25, 2, new TFormula("='Inputs advanced'!F164"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(26, 1, xls.AddFormat(fmt));
            xls.SetCellValue(26, 1, "Giesha");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));
            xls.SetCellValue(26, 2, new TFormula("='Inputs advanced'!F165"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(27, 1, xls.AddFormat(fmt));
            xls.SetCellValue(27, 1, "Icafe 90");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));
            xls.SetCellValue(27, 2, new TFormula("='Inputs advanced'!F166"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(28, 1, xls.AddFormat(fmt));
            xls.SetCellValue(28, 1, "Icatu");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));
            xls.SetCellValue(28, 2, new TFormula("='Inputs advanced'!F167"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(29, 1, xls.AddFormat(fmt));
            xls.SetCellValue(29, 1, "Lempira");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, new TFormula("='Inputs advanced'!F168"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(29, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(30, 1, xls.AddFormat(fmt));
            xls.SetCellValue(30, 1, "Maragogype o Marago");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));
            xls.SetCellValue(30, 2, new TFormula("='Inputs advanced'!F169"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(30, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(31, 1, xls.AddFormat(fmt));
            xls.SetCellValue(31, 1, "Pacamara");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));
            xls.SetCellValue(31, 2, new TFormula("='Inputs advanced'!F170"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(32, 1, xls.AddFormat(fmt));
            xls.SetCellValue(32, 1, "Pache");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(32, 2, xls.AddFormat(fmt));
            xls.SetCellValue(32, 2, new TFormula("='Inputs advanced'!F171"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(32, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(32, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(32, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(33, 1, xls.AddFormat(fmt));
            xls.SetCellValue(33, 1, "Parainema");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, new TFormula("='Inputs advanced'!F172"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(33, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(34, 1, xls.AddFormat(fmt));
            xls.SetCellValue(34, 1, "Suprema");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));
            xls.SetCellValue(34, 2, new TFormula("='Inputs advanced'!F173"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(34, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(35, 1, xls.AddFormat(fmt));
            xls.SetCellValue(35, 1, "Tipico");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));
            xls.SetCellValue(35, 2, new TFormula("='Inputs advanced'!F174"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(35, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(36, 1, xls.AddFormat(fmt));
            xls.SetCellValue(36, 1, "Villaserechi");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));
            xls.SetCellValue(36, 2, new TFormula("='Inputs advanced'!F175"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(36, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 1);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(37, 1, xls.AddFormat(fmt));
            xls.SetCellValue(37, 1, "Otra variedad:");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));
            xls.SetCellValue(37, 2, new TFormula("='Inputs advanced'!F176"));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(37, 8, xls.AddFormat(fmt));
            xls.SetCellValue(38, 1, "Arboles en Total");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "";
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));
            xls.SetCellValue(38, 2, new TFormula("=SUM(B18:B30)"));

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(38, 8, xls.AddFormat(fmt));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Format = "0";
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(39, 8, xls.AddFormat(fmt));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Format = "0";
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(40, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(41, 1, xls.AddFormat(fmt));
            xls.SetCellValue(41, 1, "Metodos de Producción");

            fmt = xls.GetCellVisibleFormatDef(41, 2);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(41, 2, xls.AddFormat(fmt));
            xls.SetCellValue(41, 2, "1 = yes");

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));
            xls.SetCellValue(41, 3, "Escriba 1 si desea incluir insumos quimicos u organico. Pero ojo pues eso puede afectar"
            + " productividades");

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(41, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(41, 10, xls.AddFormat(fmt));
            xls.SetCellValue(41, 10, "If both go with the most important");

            fmt = xls.GetCellVisibleFormatDef(41, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(41, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 12);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(41, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 13);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(41, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 1);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(42, 1, xls.AddFormat(fmt));
            xls.SetCellValue(42, 1, "Finca Quimica");

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));
            xls.SetCellValue(42, 2, new TFormula("='Inputs advanced'!F178"));

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Format = "0";
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(42, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 1);
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(43, 1, xls.AddFormat(fmt));
            xls.SetCellValue(43, 1, "Finca Organica");

            fmt = xls.GetCellVisibleFormatDef(43, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(43, 2, xls.AddFormat(fmt));
            xls.SetCellValue(43, 2, new TFormula("='Inputs advanced'!F179"));

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Format = "0";
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));
            xls.SetCellValue(44, 1, "Finca transición");

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x80, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));
            xls.SetCellValue(44, 2, new TFormula("='Inputs advanced'!F180"));

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Format = "0";
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.Indent = 3;
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(44, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Format = "0";
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(45, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 1);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 1, xls.AddFormat(fmt));
            xls.SetCellValue(46, 1, "Ingresos");

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Format = "0";
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(46, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(47, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));
            xls.SetCellValue(47, 2, "Moneda local ");

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Format = "0";
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));
            xls.SetCellValue(47, 4, "USD Dollars");

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 8, xls.AddFormat(fmt));
            xls.SetCellValue(47, 8, "moneda local/kilo");

            fmt = xls.GetCellVisibleFormatDef(48, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(48, 1, xls.AddFormat(fmt));

            TRTFRun[] Runs;
            Runs = new TRTFRun[4];
            Runs[0].FirstChar = 23;
            TFlxFont fnt;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 31;
            fnt = xls.GetDefaultFont;
            Runs[1].FontIndex = xls.AddFont(fnt);
            Runs[2].FirstChar = 34;
            fnt = xls.GetDefaultFont;
            fnt.Style = TFlxFontStyles.Bold;
            Runs[2].FontIndex = xls.AddFont(fnt);
            Runs[3].FirstChar = 48;
            fnt = xls.GetDefaultFont;
            Runs[3].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(48, 1, new TRichString("Cual fue el precio por QUINTAL de café pergamino, que usted recibió en la ultima cosecha"
            + " sin contar ninguna prima? ", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(48, 1, "Cual fue el precio por&nbsp;<font color = 'red'>QUINTAL&nbsp;</font>de&nbsp;<b>caf&eacute;"
            //+" pergamino</b>, que usted recibi&oacute; en la ultima cosecha sin contar ninguna prima?&nbsp;")


    fmt = xls.GetCellVisibleFormatDef(48, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 2, xls.AddFormat(fmt));
            xls.SetCellValue(48, 2, new TFormula("='Inputs advanced'!F196"));

            fmt = xls.GetCellVisibleFormatDef(48, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(48, 3, xls.AddFormat(fmt));
            xls.SetCellValue(48, 3, "moneda local/kilo");

            fmt = xls.GetCellVisibleFormatDef(48, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 4, xls.AddFormat(fmt));
            xls.SetCellValue(48, 4, new TFormula("=B50/B67"));

            fmt = xls.GetCellVisibleFormatDef(48, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 5, xls.AddFormat(fmt));
            xls.SetCellValue(48, 6, "$/kilo");

            fmt = xls.GetCellVisibleFormatDef(48, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0000";
            xls.SetCellFormat(48, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(49, 1, xls.AddFormat(fmt));
            xls.SetCellValue(49, 1, "Si no tiene claro el precio por kilo, precio por medida alternativa (Ex: $/quintal)"
            + " ");

            fmt = xls.GetCellVisibleFormatDef(49, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 2, xls.AddFormat(fmt));
            xls.SetCellValue(49, 2, new TFormula("='Inputs advanced'!F196"));

            fmt = xls.GetCellVisibleFormatDef(49, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(49, 3, xls.AddFormat(fmt));
            xls.SetCellValue(49, 3, "moneda local/medida alternativa");

            fmt = xls.GetCellVisibleFormatDef(49, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(49, 4, xls.AddFormat(fmt));
            xls.SetCellValue(49, 4, new TFormula("=D48/Conversiones!C11"));

            fmt = xls.GetCellVisibleFormatDef(49, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(49, 5, xls.AddFormat(fmt));
            xls.SetCellValue(49, 6, "$/lb");

            fmt = xls.GetCellVisibleFormatDef(49, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(49, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(50, 1, xls.AddFormat(fmt));
            xls.SetCellValue(50, 1, "Conversión a precio por kilo en moneda local");

            fmt = xls.GetCellVisibleFormatDef(50, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(50, 2, xls.AddFormat(fmt));
            xls.SetCellValue(50, 2, new TFormula("=IF(B49=\".\",\".\",B49/Conversiones!D14)"));

            fmt = xls.GetCellVisibleFormatDef(50, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(50, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(51, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(51, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 8);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            xls.SetCellFormat(51, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(52, 1, xls.AddFormat(fmt));
            xls.SetCellValue(52, 1, "Ingresos Cerezo por hectarea");

            fmt = xls.GetCellVisibleFormatDef(52, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(52, 2, xls.AddFormat(fmt));
            xls.SetCellValue(52, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(53, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(53, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(53, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(54, 1, xls.AddFormat(fmt));
            xls.SetCellValue(54, 1, "Primas");

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));
            xls.SetCellValue(54, 2, "Percentage Yes");

            fmt = xls.GetCellVisibleFormatDef(55, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(55, 1, xls.AddFormat(fmt));
            xls.SetCellValue(55, 1, "Ha recibido UD. En algun momento algun premio asociado a su produccion de café?");

            fmt = xls.GetCellVisibleFormatDef(55, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent1, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(55, 2, xls.AddFormat(fmt));
            xls.SetCellValue(55, 2, new TFormula("='Inputs advanced'!F197"));

            fmt = xls.GetCellVisibleFormatDef(56, 1);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(56, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 31;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 91;
            fnt = xls.GetDefaultFont;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(56, 1, new TRichString("Cual prima? (Si cero es por que al menos del 50% de los productores afirma recibir"+ " la prima)", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(56, 1, "Cual prima? (Si cero es por que<font color = 'blue'>&nbsp;al menos del 50% de los"
            //+" productores afirma recibir la prima</font>)")


    fmt = xls.GetCellVisibleFormatDef(56, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(56, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(56, 3, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 22;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 46;
            fnt = xls.GetDefaultFont;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(56, 3, new TRichString("Cual es el valorde esa prima en pergamino seco (en caso inclusive que menos del 50%"
            + " de productores afirme recibirla):", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(56, 3, "Cual es el valorde esa<font color = 'blue'>&nbsp;prima en pergamino seco</font>&nbsp;(en"
            //+" caso inclusive que menos del 50% de productores afirme recibirla):")

    xls.SetCellValue(56, 4, "kilo");
            xls.SetCellValue(56, 6, "libra ");
            xls.SetCellValue(56, 7, "quintal");

            fmt = xls.GetCellVisibleFormatDef(56, 16);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(56, 16, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 5;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(56, 16, new TRichString("Nota: Esta información es para efecto comparativo con lo que diga la cooperativa", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(56, 16, "<b>Nota:</b>&nbsp;Esta informaci&oacute;n es para efecto comparativo con lo que diga"
            //+" la cooperativa")


    fmt = xls.GetCellVisibleFormatDef(56, 17);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(56, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 18);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(56, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 19);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(56, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(57, 1, xls.AddFormat(fmt));
            xls.SetCellValue(57, 1, "Fair Trade ");

            fmt = xls.GetCellVisibleFormatDef(57, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(57, 2, xls.AddFormat(fmt));
            xls.SetCellValue(57, 2, 50);
            xls.SetCellValue(57, 3, "Fair Trade ");

            fmt = xls.GetCellVisibleFormatDef(57, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 4, xls.AddFormat(fmt));
            xls.SetCellValue(57, 4, new TFormula("='Inputs advanced'!F198"));

            fmt = xls.GetCellVisibleFormatDef(57, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 6, xls.AddFormat(fmt));
            xls.SetCellValue(57, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(57, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(57, 7, xls.AddFormat(fmt));
            xls.SetCellValue(57, 7, 0);
            xls.SetCellValue(57, 9, new TFormula("=F57"));
            xls.SetCellValue(57, 10, 0);

            fmt = xls.GetCellVisibleFormatDef(58, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(58, 1, xls.AddFormat(fmt));
            xls.SetCellValue(58, 1, "Organic");

            fmt = xls.GetCellVisibleFormatDef(58, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(58, 2, xls.AddFormat(fmt));
            xls.SetCellValue(58, 2, 0);
            xls.SetCellValue(58, 3, "Organic");

            fmt = xls.GetCellVisibleFormatDef(58, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(58, 4, xls.AddFormat(fmt));
            xls.SetCellValue(58, 4, new TFormula("='Inputs advanced'!F199"));

            fmt = xls.GetCellVisibleFormatDef(58, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(58, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(58, 6, xls.AddFormat(fmt));
            xls.SetCellValue(58, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(58, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(58, 7, xls.AddFormat(fmt));
            xls.SetCellValue(58, 7, 0);
            xls.SetCellValue(58, 9, new TFormula("=F58"));
            xls.SetCellValue(58, 10, 0);

            fmt = xls.GetCellVisibleFormatDef(59, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(59, 1, xls.AddFormat(fmt));
            xls.SetCellValue(59, 1, "Prima \"cooperativa\"");

            fmt = xls.GetCellVisibleFormatDef(59, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(59, 2, xls.AddFormat(fmt));
            xls.SetCellValue(59, 2, 0);
            xls.SetCellValue(59, 3, "Prima \"cooperativa\"");

            fmt = xls.GetCellVisibleFormatDef(59, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(59, 4, xls.AddFormat(fmt));
            xls.SetCellValue(59, 4, new TFormula("='Inputs advanced'!F200"));

            fmt = xls.GetCellVisibleFormatDef(59, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(59, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(59, 6, xls.AddFormat(fmt));
            xls.SetCellValue(59, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(59, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(59, 7, xls.AddFormat(fmt));
            xls.SetCellValue(59, 7, 0);
            xls.SetCellValue(59, 9, new TFormula("=F59"));
            xls.SetCellValue(59, 10, 0);

            fmt = xls.GetCellVisibleFormatDef(60, 1);
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(60, 1, xls.AddFormat(fmt));
            xls.SetCellValue(60, 1, "Otra:                                   .                    ");

            fmt = xls.GetCellVisibleFormatDef(60, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 2, xls.AddFormat(fmt));
            xls.SetCellValue(60, 2, 0);
            xls.SetCellValue(60, 3, "Otra:                                   .                    ");

            fmt = xls.GetCellVisibleFormatDef(60, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 4, xls.AddFormat(fmt));
            xls.SetCellValue(60, 4, new TFormula("='Inputs advanced'!F201"));

            fmt = xls.GetCellVisibleFormatDef(60, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(60, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 1);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(61, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 10, xls.AddFormat(fmt));
            xls.SetCellValue(61, 10, "A");

            fmt = xls.GetCellVisibleFormatDef(61, 11);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 11, xls.AddFormat(fmt));
            xls.SetCellValue(61, 11, "B");

            fmt = xls.GetCellVisibleFormatDef(61, 12);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 12, xls.AddFormat(fmt));
            xls.SetCellValue(61, 12, "C");

            fmt = xls.GetCellVisibleFormatDef(62, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(62, 1, xls.AddFormat(fmt));
            xls.SetCellValue(62, 1, "Nota : Los siguientes datos de prima de café oro en dolares pueden tomarse de lo que"
            + " diga el productor, de la prima que se sabe se da en el mercado, o de una prima que"
            + " la cooperativa reeporta transfiere al productor. Ver vinculo.");

            fmt = xls.GetCellVisibleFormatDef(62, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(62, 2, xls.AddFormat(fmt));
            xls.SetCellValue(62, 2, "Prima libra oro");
            xls.SetCellValue(62, 4, "Oficial");

            fmt = xls.GetCellVisibleFormatDef(62, 8);
            fmt.WrapText = true;
            xls.SetCellFormat(62, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 9);
            fmt.WrapText = true;
            xls.SetCellFormat(62, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 10);
            fmt.WrapText = true;
            xls.SetCellFormat(62, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 12);
            fmt.WrapText = true;
            xls.SetCellFormat(62, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(62, 13, xls.AddFormat(fmt));
            xls.SetCellValue(63, 1, "Premio Fair Tade (centavos por libra CAFÉ ORO)");

            fmt = xls.GetCellVisibleFormatDef(63, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(63, 2, xls.AddFormat(fmt));
            xls.SetCellValue(63, 2, 0);
            xls.SetCellValue(63, 3, "USD/libra");

            fmt = xls.GetCellVisibleFormatDef(63, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(63, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(63, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(63, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(63, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 13);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(63, 13, xls.AddFormat(fmt));
            xls.SetCellValue(64, 1, "Premio Organico (centavos por libra CAFÉ ORO)");

            fmt = xls.GetCellVisibleFormatDef(64, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(64, 2, xls.AddFormat(fmt));
            xls.SetCellValue(64, 2, 0);
            xls.SetCellValue(64, 3, "USD/libra");

            fmt = xls.GetCellVisibleFormatDef(64, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(64, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(64, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(64, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(64, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 13);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(64, 13, xls.AddFormat(fmt));
            xls.SetCellValue(65, 1, "Premio \"Cooperativa\" (centavos por libra CAFÉ ORO)");

            fmt = xls.GetCellVisibleFormatDef(65, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(65, 2, xls.AddFormat(fmt));
            xls.SetCellValue(65, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(65, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 3, xls.AddFormat(fmt));
            xls.SetCellValue(65, 3, "Incluida ya en precio considerado");

            fmt = xls.GetCellVisibleFormatDef(65, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(65, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(65, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 13);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(65, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(66, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(66, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(66, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(66, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(66, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 13);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(66, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(67, 1, xls.AddFormat(fmt));
            xls.SetCellValue(67, 1, "Tasa de cambio");

            fmt = xls.GetCellVisibleFormatDef(67, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.000";
            xls.SetCellFormat(67, 2, xls.AddFormat(fmt));
            xls.SetCellValue(67, 2, new TFormula("=Conversiones!$F$24"));
            xls.SetCellValue(67, 3, "PES/USD");

            fmt = xls.GetCellVisibleFormatDef(68, 1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(68, 1, xls.AddFormat(fmt));

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            xls.SetCellFormat(68, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(69, 1, xls.AddFormat(fmt));
            xls.SetCellValue(69, 1, "Jornal");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Percent, 0), true);
            xls.SetCellFormat(69, 2, xls.AddFormat(fmt));
            xls.SetCellValue(70, 2, "PES");
            xls.SetCellValue(70, 3, "USD");
            xls.SetCellValue(71, 1, "Cual es el valor del Jornal en la zona? (el pago por días)");

            fmt = xls.GetCellVisibleFormatDef(71, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 2, xls.AddFormat(fmt));
            xls.SetCellValue(71, 2, new TFormula("='Inputs advanced'!F202+B75"));

            fmt = xls.GetCellVisibleFormatDef(71, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(71, 3, xls.AddFormat(fmt));
            xls.SetCellValue(71, 3, new TFormula("=B71/$B$67"));

            fmt = xls.GetCellVisibleFormatDef(72, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent5, -0.249977111117893);
            xls.SetCellFormat(72, 1, xls.AddFormat(fmt));
            xls.SetCellValue(72, 1, "Cual es el salario minimo mensual por ley?");

            fmt = xls.GetCellVisibleFormatDef(72, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent1, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(72, 2, xls.AddFormat(fmt));
            xls.SetCellValue(72, 2, new TFormula("='Inputs advanced'!F205"));

            fmt = xls.GetCellVisibleFormatDef(72, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(72, 3, xls.AddFormat(fmt));
            xls.SetCellValue(72, 3, new TFormula("=B72/$B$67"));

            fmt = xls.GetCellVisibleFormatDef(73, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent5, -0.249977111117893);
            xls.SetCellFormat(73, 1, xls.AddFormat(fmt));
            xls.SetCellValue(73, 1, "Ingreso annual aproximado");

            fmt = xls.GetCellVisibleFormatDef(73, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(73, 2, xls.AddFormat(fmt));
            xls.SetCellValue(73, 2, new TFormula("=B72*12"));

            fmt = xls.GetCellVisibleFormatDef(73, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(73, 3, xls.AddFormat(fmt));
            xls.SetCellValue(73, 3, new TFormula("=B73/$B$67"));
            xls.SetCellValue(74, 1, "Valor por caja del Recolector");

            fmt = xls.GetCellVisibleFormatDef(74, 2);
            fmt.Format = "0.00";
            xls.SetCellFormat(74, 2, xls.AddFormat(fmt));
            xls.SetCellValue(74, 2, new TFormula("='Inputs advanced'!F144"));
            xls.SetCellValue(75, 1, "Alimentacion Recolector");
            xls.SetCellValue(75, 2, new TFormula("='Inputs advanced'!F204"));

            fmt = xls.GetCellVisibleFormatDef(76, 1);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xEE, 0x50, 0xAD);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(76, 1, xls.AddFormat(fmt));
            xls.SetCellValue(76, 1, "Ingresos Indirectos");

            fmt = xls.GetCellVisibleFormatDef(76, 2);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 3);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 4);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 5);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 6);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 7);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 8);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 9);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 10);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 11);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 12);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 13);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 14);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 15);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 16);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 17);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 18);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 19);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 20);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 21);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 22);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 23);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 24);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 25);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 25, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 26);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 26, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 27);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(76, 27, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 1, xls.AddFormat(fmt));
            xls.SetCellValue(78, 1, "Transferencias de la cooperativa en dinero o bienes");

            fmt = xls.GetCellVisibleFormatDef(78, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(78, 2, xls.AddFormat(fmt));
            xls.SetCellValue(78, 2, "Percentage Yes");

            fmt = xls.GetCellVisibleFormatDef(79, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(79, 1, xls.AddFormat(fmt));
            xls.SetCellValue(79, 1, "Recibio usted otros ingresos por parte de la cooperativa diferentes a prestamos? (Regalos,"
            + " premios, ayudas). Cuál fue la suma?");

            fmt = xls.GetCellVisibleFormatDef(79, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(79, 2, xls.AddFormat(fmt));
            xls.SetCellValue(79, 2, new TFormula("='Inputs advanced'!F219"));

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));
            xls.SetCellValue(79, 3, "En que Año?:                                           ");

            fmt = xls.GetCellVisibleFormatDef(79, 4);
            fmt.WrapText = true;
            xls.SetCellFormat(79, 4, xls.AddFormat(fmt));
            xls.SetCellValue(79, 4, "Cantidad anual en moneda local");

            fmt = xls.GetCellVisibleFormatDef(79, 5);
            fmt.WrapText = true;
            xls.SetCellFormat(79, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(80, 1, xls.AddFormat(fmt));
            xls.SetCellValue(80, 3, "Preparacion terreno (Año 0) ");

            fmt = xls.GetCellVisibleFormatDef(80, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(80, 4, xls.AddFormat(fmt));
            xls.SetCellValue(80, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(80, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(80, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(81, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(81, 3, xls.AddFormat(fmt));
            xls.SetCellValue(81, 3, "Fertilización y control de plagas (Año 1)");

            fmt = xls.GetCellVisibleFormatDef(81, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(81, 4, xls.AddFormat(fmt));
            xls.SetCellValue(81, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(81, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(81, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(82, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(82, 3, xls.AddFormat(fmt));
            xls.SetCellValue(82, 3, "Cosecha y Postcosecha (Año 2)");

            fmt = xls.GetCellVisibleFormatDef(82, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 4, xls.AddFormat(fmt));
            xls.SetCellValue(82, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(82, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(83, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(83, 3, xls.AddFormat(fmt));
            xls.SetCellValue(83, 3, "Cosecha y Postcosecha (Año 3)");

            fmt = xls.GetCellVisibleFormatDef(83, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 4, xls.AddFormat(fmt));
            xls.SetCellValue(83, 4, new TFormula("='Inputs advanced'!F222"));

            fmt = xls.GetCellVisibleFormatDef(83, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(84, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(84, 3, xls.AddFormat(fmt));
            xls.SetCellValue(84, 3, "Cosecha y Postcosecha (Año 4)");

            fmt = xls.GetCellVisibleFormatDef(84, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 4, xls.AddFormat(fmt));
            xls.SetCellValue(84, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(84, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(84, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(85, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(85, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(85, 3, xls.AddFormat(fmt));
            xls.SetCellValue(85, 3, "Cosecha y Postcosecha (Año 5)");

            fmt = xls.GetCellVisibleFormatDef(85, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(85, 4, xls.AddFormat(fmt));
            xls.SetCellValue(85, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(85, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(85, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(86, 3, xls.AddFormat(fmt));
            xls.SetCellValue(86, 3, "Cosecha y Postcosecha (Año 6)");

            fmt = xls.GetCellVisibleFormatDef(86, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(86, 4, xls.AddFormat(fmt));
            xls.SetCellValue(86, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(86, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(86, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(87, 3, xls.AddFormat(fmt));
            xls.SetCellValue(87, 3, "Cosecha y Postcosecha (Año 7)");

            fmt = xls.GetCellVisibleFormatDef(87, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 4, xls.AddFormat(fmt));
            xls.SetCellValue(87, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(87, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(87, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(88, 3, xls.AddFormat(fmt));
            xls.SetCellValue(88, 3, "Cosecha y Postcosecha (Año 8)");

            fmt = xls.GetCellVisibleFormatDef(88, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 4, xls.AddFormat(fmt));
            xls.SetCellValue(88, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(88, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(88, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(89, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(89, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 1, xls.AddFormat(fmt));
            xls.SetCellValue(90, 1, "Capacitaciones");

            fmt = xls.GetCellVisibleFormatDef(90, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(90, 2, xls.AddFormat(fmt));
            xls.SetCellValue(90, 2, "Percentage Yes");

            fmt = xls.GetCellVisibleFormatDef(90, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(90, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(90, 22, xls.AddFormat(fmt));
            xls.SetCellValue(91, 1, "Recibio por parte de la cooperativa algún tipo de capaciatación?");

            fmt = xls.GetCellVisibleFormatDef(91, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(91, 2, xls.AddFormat(fmt));
            xls.SetCellValue(91, 2, new TFormula("='Inputs advanced'!F223"));

            fmt = xls.GetCellVisibleFormatDef(91, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(91, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(91, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(91, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(92, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(92, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(92, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(93, 1, xls.AddFormat(fmt));
            xls.SetCellValue(93, 1, "Area de capacitacion:");

            fmt = xls.GetCellVisibleFormatDef(93, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(93, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(93, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(93, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(93, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(93, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(93, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 1, xls.AddFormat(fmt));
            xls.SetCellValue(94, 1, "Numero de años");

            fmt = xls.GetCellVisibleFormatDef(94, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(94, 2, xls.AddFormat(fmt));
            xls.SetCellValue(94, 2, new TFormula("='Inputs advanced'!F226"));

            fmt = xls.GetCellVisibleFormatDef(94, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(94, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(94, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(94, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 1, xls.AddFormat(fmt));
            xls.SetCellValue(95, 1, "Número días por año");

            fmt = xls.GetCellVisibleFormatDef(95, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 2, xls.AddFormat(fmt));
            xls.SetCellValue(95, 2, new TFormula("='Inputs advanced'!F227"));

            fmt = xls.GetCellVisibleFormatDef(95, 3);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 4, xls.AddFormat(fmt));
            xls.SetCellValue(95, 4, "Valor monetario de las capacitaciones");

            fmt = xls.GetCellVisibleFormatDef(95, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(95, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(95, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(95, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 2, xls.AddFormat(fmt));
            xls.SetCellValue(96, 3, "Preparacion terreno (Año 0) ");

            fmt = xls.GetCellVisibleFormatDef(96, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(96, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(96, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(97, 3, xls.AddFormat(fmt));
            xls.SetCellValue(97, 3, "Fertilización y control de plagas (Año 1)");

            fmt = xls.GetCellVisibleFormatDef(97, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(97, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(97, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(97, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(98, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(98, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(98, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(98, 3, xls.AddFormat(fmt));
            xls.SetCellValue(98, 3, "Cosecha y Postcosecha (Año 2)");

            fmt = xls.GetCellVisibleFormatDef(98, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(98, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(98, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(98, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(99, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(99, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(99, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(99, 3, xls.AddFormat(fmt));
            xls.SetCellValue(99, 3, "Cosecha y Postcosecha (Año 3)");

            fmt = xls.GetCellVisibleFormatDef(99, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(99, 4, xls.AddFormat(fmt));
            xls.SetCellValue(99, 4, new TFormula("=B$95*B$71"));

            fmt = xls.GetCellVisibleFormatDef(99, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(99, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(100, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(100, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(100, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(100, 3, xls.AddFormat(fmt));
            xls.SetCellValue(100, 3, "Cosecha y Postcosecha (Año 4)");

            fmt = xls.GetCellVisibleFormatDef(100, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(100, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(100, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(100, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(101, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(101, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(101, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(101, 3, xls.AddFormat(fmt));
            xls.SetCellValue(101, 3, "Cosecha y Postcosecha (Año 5)");

            fmt = xls.GetCellVisibleFormatDef(101, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(101, 4, xls.AddFormat(fmt));
            xls.SetCellValue(101, 4, new TFormula("=B$95*B$71"));

            fmt = xls.GetCellVisibleFormatDef(101, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(101, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(102, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(102, 3, xls.AddFormat(fmt));
            xls.SetCellValue(102, 3, "Cosecha y Postcosecha (Año 6)");

            fmt = xls.GetCellVisibleFormatDef(102, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(102, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(102, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(102, 5, xls.AddFormat(fmt));
            xls.SetCellValue(102, 13, new TFormula("=1000*(1+0.2)"));
            xls.SetCellValue(102, 14, new TFormula("=0.2*3"));

            fmt = xls.GetCellVisibleFormatDef(103, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(103, 3, xls.AddFormat(fmt));
            xls.SetCellValue(103, 3, "Cosecha y Postcosecha (Año 7)");

            fmt = xls.GetCellVisibleFormatDef(103, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(103, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(103, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(103, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(104, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(104, 3, xls.AddFormat(fmt));
            xls.SetCellValue(104, 3, "Cosecha y Postcosecha (Año 8)");

            fmt = xls.GetCellVisibleFormatDef(104, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(104, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(104, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(104, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(105, 4, xls.AddFormat(fmt));
            xls.SetCellValue(105, 4, "Nota 1: Un prestamo que se cancela con un intervalo de más de 1.5 años se apunta en"
            + " el siguiente renglón. (1.5 aproxima a 2do renglon contando el renglon del año del"
            + " prestamo)");

            fmt = xls.GetCellVisibleFormatDef(105, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(105, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(105, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(105, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(105, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(105, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(105, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(105, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(105, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(106, 4, xls.AddFormat(fmt));
            xls.SetCellValue(106, 4, "Ejemplo 1: Prestamo Enero 2015, Cancelado Diciembre 2015 mismo renglon (menos de 18"
            + " meses). Ejemplo 2: Prestamo Enero 2015, Cancelacion Julio 2016, devolucion va siguiente"
            + " renglon.   ");

            fmt = xls.GetCellVisibleFormatDef(106, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(106, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(106, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(106, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(106, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(106, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(106, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(106, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(106, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(107, 4, xls.AddFormat(fmt));
            xls.SetCellValue(107, 4, "Nota 2: En caso de más de un prestamo registrar año de recibido y devolución en todos"
            + " los casos con los códigos X-Y, A-B, C-D etc");

            fmt = xls.GetCellVisibleFormatDef(107, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(107, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(107, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(107, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(107, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(107, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(107, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(107, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(108, 2, xls.AddFormat(fmt));
            xls.SetCellValue(108, 2, "Porcentaje afirmativo");

            fmt = xls.GetCellVisibleFormatDef(108, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(108, 4, xls.AddFormat(fmt));
            xls.SetCellValue(108, 4, "Nota 3: Para más de un prestamo en el mismo año sumar los prestamos. Tomar el año"
            + " de devolucion como el año medio (en caso de ser distinto años)");

            fmt = xls.GetCellVisibleFormatDef(108, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(108, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(108, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(108, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(108, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(108, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(108, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(108, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(109, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(109, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 28;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 37;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.Automatic;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(109, 1, new TRichString("Recibio algún prestamo para inversión por parte de la cooperativa?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(109, 1, "Recibio alg&uacute;n prestamo para&nbsp;<font color = 'blue'>inversi&oacute;n</font>&nbsp;por"
            //+" parte de la cooperativa?")


    fmt = xls.GetCellVisibleFormatDef(109, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(109, 2, xls.AddFormat(fmt));
            xls.SetCellValue(109, 2, new TFormula("='Inputs advanced'!F230"));

            fmt = xls.GetCellVisibleFormatDef(109, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(109, 3, xls.AddFormat(fmt));
            xls.SetCellValue(109, 3, "En que Año?:                                           ");

            fmt = xls.GetCellVisibleFormatDef(109, 4);
            fmt.WrapText = true;
            xls.SetCellFormat(109, 4, xls.AddFormat(fmt));
            xls.SetCellValue(109, 4, "Marque X (mayuscula) en el año en que recibio el prestamo");

            fmt = xls.GetCellVisibleFormatDef(109, 5);
            fmt.WrapText = true;
            xls.SetCellFormat(109, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(109, 6);
            fmt.WrapText = true;
            xls.SetCellFormat(109, 6, xls.AddFormat(fmt));
            xls.SetCellValue(109, 6, "Cantidad en moneda local");

            fmt = xls.GetCellVisibleFormatDef(109, 7);
            fmt.WrapText = true;
            xls.SetCellFormat(109, 7, xls.AddFormat(fmt));
            xls.SetCellValue(109, 7, "Cuando termina o terminó de pagar el prestamo? (Marque Y mayuscula)");

            fmt = xls.GetCellVisibleFormatDef(109, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(109, 9, xls.AddFormat(fmt));
            xls.SetCellValue(109, 9, "Resumen Prestamos");

            fmt = xls.GetCellVisibleFormatDef(109, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(109, 10, xls.AddFormat(fmt));
            xls.SetCellValue(109, 10, "Resumen Pagos");

            fmt = xls.GetCellVisibleFormatDef(109, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(109, 11, xls.AddFormat(fmt));
            xls.SetCellValue(109, 11, "Valor prestamo");

            fmt = xls.GetCellVisibleFormatDef(109, 12);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.WrapText = true;
            xls.SetCellFormat(109, 12, xls.AddFormat(fmt));
            xls.SetCellValue(109, 12, "Cual es la tasa de interés annual de dicho prestamo");

            fmt = xls.GetCellVisibleFormatDef(109, 13);
            fmt.WrapText = true;
            xls.SetCellFormat(109, 13, xls.AddFormat(fmt));
            xls.SetCellValue(109, 13, "Cuanto pagó en total por ese prestamo?");

            fmt = xls.GetCellVisibleFormatDef(109, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 14, xls.AddFormat(fmt));
            xls.SetCellValue(109, 14, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(109, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 15, xls.AddFormat(fmt));
            xls.SetCellValue(109, 15, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(109, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 16, xls.AddFormat(fmt));
            xls.SetCellValue(109, 16, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(109, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 17, xls.AddFormat(fmt));
            xls.SetCellValue(109, 17, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(109, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 18, xls.AddFormat(fmt));
            xls.SetCellValue(109, 18, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(109, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 19, xls.AddFormat(fmt));
            xls.SetCellValue(109, 19, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(109, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 20, xls.AddFormat(fmt));
            xls.SetCellValue(109, 20, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(109, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 21, xls.AddFormat(fmt));
            xls.SetCellValue(109, 21, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(109, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(109, 22, xls.AddFormat(fmt));
            xls.SetCellValue(109, 22, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(110, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 1, xls.AddFormat(fmt));
            xls.SetCellValue(110, 1, "Tiempo del prestamo");

            fmt = xls.GetCellVisibleFormatDef(110, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(110, 2, xls.AddFormat(fmt));
            xls.SetCellValue(110, 2, new TFormula("='Inputs advanced'!F235"));
            xls.SetCellValue(110, 3, "Preparacion terreno (Año 0) ");

            fmt = xls.GetCellVisibleFormatDef(110, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 4, xls.AddFormat(fmt));
            xls.SetCellValue(110, 4, "X");

            fmt = xls.GetCellVisibleFormatDef(110, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 6, xls.AddFormat(fmt));
            xls.SetCellValue(110, 6, new TFormula("=B111"));

            fmt = xls.GetCellVisibleFormatDef(110, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 7, xls.AddFormat(fmt));
            xls.SetCellValue(110, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(110, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 8, xls.AddFormat(fmt));
            xls.SetCellValue(110, 8, 0);

            fmt = xls.GetCellVisibleFormatDef(110, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 9, xls.AddFormat(fmt));
            xls.SetCellValue(110, 9, "X");

            fmt = xls.GetCellVisibleFormatDef(110, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 10, xls.AddFormat(fmt));
            xls.SetCellValue(110, 10, "Y");

            fmt = xls.GetCellVisibleFormatDef(110, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 11, xls.AddFormat(fmt));
            xls.SetCellValue(110, 11, new TFormula("=IF($B$115=1,VLOOKUP(I110,$D$110:$F$118,2,FALSE),0)"));

            fmt = xls.GetCellVisibleFormatDef(110, 12);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 12, xls.AddFormat(fmt));
            xls.SetCellValue(110, 12, new TFormula("=B112/100"));

            fmt = xls.GetCellVisibleFormatDef(110, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 13, xls.AddFormat(fmt));
            xls.SetCellValue(110, 13, new TFormula("=B111*(1+(L110*B110))"));

            fmt = xls.GetCellVisibleFormatDef(110, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 14, xls.AddFormat(fmt));
            xls.SetCellValue(110, 14, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 15, xls.AddFormat(fmt));
            xls.SetCellValue(110, 15, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 16, xls.AddFormat(fmt));
            xls.SetCellValue(110, 16, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 17, xls.AddFormat(fmt));
            xls.SetCellValue(110, 17, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 18, xls.AddFormat(fmt));
            xls.SetCellValue(110, 18, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 19);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 19, xls.AddFormat(fmt));
            xls.SetCellValue(110, 19, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 20);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 20, xls.AddFormat(fmt));
            xls.SetCellValue(110, 20, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 21);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 21, xls.AddFormat(fmt));
            xls.SetCellValue(110, 21, new TFormula("=$M110/9"));

            fmt = xls.GetCellVisibleFormatDef(110, 22);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(110, 22, xls.AddFormat(fmt));
            xls.SetCellValue(110, 22, new TFormula("=$M110/9"));
            xls.SetCellValue(110, 24, "Año Prestamo");

            fmt = xls.GetCellVisibleFormatDef(110, 25);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 25, xls.AddFormat(fmt));
            xls.SetCellValue(110, 25, "X");

            fmt = xls.GetCellVisibleFormatDef(110, 26);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 26, xls.AddFormat(fmt));
            xls.SetCellValue(110, 26, new TFormula("=IF(B115=1,VLOOKUP(Y110,D110:H118,4,FALSE),0)"));

            fmt = xls.GetCellVisibleFormatDef(110, 27);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(110, 27, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(111, 1, xls.AddFormat(fmt));
            xls.SetCellValue(111, 1, "Monto en moneda local");

            fmt = xls.GetCellVisibleFormatDef(111, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(111, 2, xls.AddFormat(fmt));
            xls.SetCellValue(111, 2, new TFormula("='Inputs advanced'!F233"));

            fmt = xls.GetCellVisibleFormatDef(111, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(111, 3, xls.AddFormat(fmt));
            xls.SetCellValue(111, 3, "Fertilización y control de plagas (Año 1)");

            fmt = xls.GetCellVisibleFormatDef(111, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 4, xls.AddFormat(fmt));
            xls.SetCellValue(111, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(111, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(111, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 7, xls.AddFormat(fmt));
            xls.SetCellValue(111, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(111, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 8, xls.AddFormat(fmt));
            xls.SetCellValue(111, 8, 1);

            fmt = xls.GetCellVisibleFormatDef(111, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 11);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 13);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(111, 13, xls.AddFormat(fmt));
            xls.SetCellValue(111, 24, "Año Pago");

            fmt = xls.GetCellVisibleFormatDef(111, 25);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 25, xls.AddFormat(fmt));
            xls.SetCellValue(111, 25, "Y");

            fmt = xls.GetCellVisibleFormatDef(111, 26);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 26, xls.AddFormat(fmt));
            xls.SetCellValue(111, 26, new TFormula("=IF(B115=1,VLOOKUP(Y111,G110:H118,2,FALSE),0)"));

            fmt = xls.GetCellVisibleFormatDef(111, 27);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(111, 27, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(112, 1, xls.AddFormat(fmt));
            xls.SetCellValue(112, 1, "Tasa de interes anual");

            fmt = xls.GetCellVisibleFormatDef(112, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(112, 2, xls.AddFormat(fmt));
            xls.SetCellValue(112, 2, new TFormula("='Inputs advanced'!F239"));

            fmt = xls.GetCellVisibleFormatDef(112, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(112, 3, xls.AddFormat(fmt));
            xls.SetCellValue(112, 3, "Cosecha y Postcosecha (Año 2)");

            fmt = xls.GetCellVisibleFormatDef(112, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(112, 4, xls.AddFormat(fmt));
            xls.SetCellValue(112, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(112, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(112, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(112, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(112, 7, xls.AddFormat(fmt));
            xls.SetCellValue(112, 7, "Y");

            fmt = xls.GetCellVisibleFormatDef(112, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(112, 8, xls.AddFormat(fmt));
            xls.SetCellValue(112, 8, 2);

            fmt = xls.GetCellVisibleFormatDef(112, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(112, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(112, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 11);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(112, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(112, 13);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(112, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(113, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(113, 3, xls.AddFormat(fmt));
            xls.SetCellValue(113, 3, "Cosecha y Postcosecha (Año 3)");

            fmt = xls.GetCellVisibleFormatDef(113, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(113, 4, xls.AddFormat(fmt));
            xls.SetCellValue(113, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(113, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(113, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(113, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(113, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(113, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(113, 7, xls.AddFormat(fmt));
            xls.SetCellValue(113, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(113, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(113, 8, xls.AddFormat(fmt));
            xls.SetCellValue(113, 8, 3);

            fmt = xls.GetCellVisibleFormatDef(113, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(113, 9, xls.AddFormat(fmt));
            xls.SetCellValue(113, 25, "Prestamo a cuantos años:");
            xls.SetCellValue(113, 26, new TFormula("=(Z111-Z110)+1"));

            fmt = xls.GetCellVisibleFormatDef(114, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(114, 2, xls.AddFormat(fmt));
            xls.SetCellValue(114, 2, "Si = 1");

            fmt = xls.GetCellVisibleFormatDef(114, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(114, 3, xls.AddFormat(fmt));
            xls.SetCellValue(114, 3, "Cosecha y Postcosecha (Año 4)");

            fmt = xls.GetCellVisibleFormatDef(114, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(114, 4, xls.AddFormat(fmt));
            xls.SetCellValue(114, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(114, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(114, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(114, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(114, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(114, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(114, 7, xls.AddFormat(fmt));
            xls.SetCellValue(114, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(114, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(114, 8, xls.AddFormat(fmt));
            xls.SetCellValue(114, 8, 4);

            fmt = xls.GetCellVisibleFormatDef(114, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(114, 9, xls.AddFormat(fmt));
            xls.SetCellValue(115, 1, "Se considera este prestamo para el productor representativo?");

            fmt = xls.GetCellVisibleFormatDef(115, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(115, 2, xls.AddFormat(fmt));
            xls.SetCellValue(115, 2, new TFormula("=IF(B109>=0.3,1,0)"));

            fmt = xls.GetCellVisibleFormatDef(115, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(115, 3, xls.AddFormat(fmt));
            xls.SetCellValue(115, 3, "Cosecha y Postcosecha (Año 5)");

            fmt = xls.GetCellVisibleFormatDef(115, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(115, 4, xls.AddFormat(fmt));
            xls.SetCellValue(115, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(115, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(115, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(115, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(115, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(115, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(115, 7, xls.AddFormat(fmt));
            xls.SetCellValue(115, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(115, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(115, 8, xls.AddFormat(fmt));
            xls.SetCellValue(115, 8, 5);

            fmt = xls.GetCellVisibleFormatDef(115, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(115, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(115, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(115, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(116, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(116, 3, xls.AddFormat(fmt));
            xls.SetCellValue(116, 3, "Cosecha y Postcosecha (Año 6)");

            fmt = xls.GetCellVisibleFormatDef(116, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(116, 4, xls.AddFormat(fmt));
            xls.SetCellValue(116, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(116, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(116, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(116, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(116, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(116, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(116, 7, xls.AddFormat(fmt));
            xls.SetCellValue(116, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(116, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(116, 8, xls.AddFormat(fmt));
            xls.SetCellValue(116, 8, 6);

            fmt = xls.GetCellVisibleFormatDef(116, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(116, 9, xls.AddFormat(fmt));
            xls.SetCellValue(116, 25, "Tasa implicita anual");

            fmt = xls.GetCellVisibleFormatDef(116, 26);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(116, 26, xls.AddFormat(fmt));
            xls.SetCellValue(116, 26, new TFormula("=IF(B115=1,((M110-K110)/Z113)/K110,0)"));

            fmt = xls.GetCellVisibleFormatDef(117, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(117, 3, xls.AddFormat(fmt));
            xls.SetCellValue(117, 3, "Cosecha y Postcosecha (Año 7)");

            fmt = xls.GetCellVisibleFormatDef(117, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(117, 4, xls.AddFormat(fmt));
            xls.SetCellValue(117, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(117, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(117, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(117, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(117, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(117, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(117, 7, xls.AddFormat(fmt));
            xls.SetCellValue(117, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(117, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(117, 8, xls.AddFormat(fmt));
            xls.SetCellValue(117, 8, 7);

            fmt = xls.GetCellVisibleFormatDef(117, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(117, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(118, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(118, 3, xls.AddFormat(fmt));
            xls.SetCellValue(118, 3, "Cosecha y Postcosecha (Año 8)");

            fmt = xls.GetCellVisibleFormatDef(118, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(118, 4, xls.AddFormat(fmt));
            xls.SetCellValue(118, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(118, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(118, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(118, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(118, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(118, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(118, 7, xls.AddFormat(fmt));
            xls.SetCellValue(118, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(118, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(118, 8, xls.AddFormat(fmt));
            xls.SetCellValue(118, 8, 8);

            fmt = xls.GetCellVisibleFormatDef(118, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(118, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(119, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(119, 2, xls.AddFormat(fmt));
            xls.SetCellValue(119, 2, "Total Pago Prestamos Coop:");

            fmt = xls.GetCellVisibleFormatDef(119, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(119, 13, xls.AddFormat(fmt));
            xls.SetCellValue(119, 13, "Total Pago Prestamos Coop.");

            fmt = xls.GetCellVisibleFormatDef(119, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 14, xls.AddFormat(fmt));
            xls.SetCellValue(119, 14, new TFormula("=SUM(N110:N118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 15, xls.AddFormat(fmt));
            xls.SetCellValue(119, 15, new TFormula("=SUM(O110:O118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 16, xls.AddFormat(fmt));
            xls.SetCellValue(119, 16, new TFormula("=SUM(P110:P118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 17, xls.AddFormat(fmt));
            xls.SetCellValue(119, 17, new TFormula("=SUM(Q110:Q118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 18, xls.AddFormat(fmt));
            xls.SetCellValue(119, 18, new TFormula("=SUM(R110:R118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 19, xls.AddFormat(fmt));
            xls.SetCellValue(119, 19, new TFormula("=SUM(S110:S118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 20, xls.AddFormat(fmt));
            xls.SetCellValue(119, 20, new TFormula("=SUM(T110:T118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 21, xls.AddFormat(fmt));
            xls.SetCellValue(119, 21, new TFormula("=SUM(U110:U118)"));

            fmt = xls.GetCellVisibleFormatDef(119, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(119, 22, xls.AddFormat(fmt));
            xls.SetCellValue(119, 22, new TFormula("=SUM(V110:V118)"));

            fmt = xls.GetCellVisibleFormatDef(120, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(120, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(121, 4, xls.AddFormat(fmt));
            xls.SetCellValue(121, 4, "Nota 1: Un prestamo que se cancela con un intervalo de más de 1.5 años se apunta en"
            + " el siguiente renglón. (1.5 aproxima a 2do renglon contando el renglon del año del"
            + " prestamo)");

            fmt = xls.GetCellVisibleFormatDef(121, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(121, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(121, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(121, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(121, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(121, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 10);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(121, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(121, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(121, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(122, 4, xls.AddFormat(fmt));
            xls.SetCellValue(122, 4, "Ejemplo 1: Prestamo Enero 2015, Cancelado Diciembre 2015 mismo renglon (menos de 18"
            + " meses). Ejemplo 2: Prestamo Enero 2015, Cancelacion Julio 2016, devolucion va siguiente"
            + " renglon.   ");

            fmt = xls.GetCellVisibleFormatDef(122, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(122, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(122, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(122, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(122, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(122, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(122, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(122, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(123, 4, xls.AddFormat(fmt));
            xls.SetCellValue(123, 4, "Nota 2: En caso de más de un prestamo registrar año de recibido y devolución en todos"
            + " los casos con los códigos X-Y, A-B, C-D etc");

            fmt = xls.GetCellVisibleFormatDef(123, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(123, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(123, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(123, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(123, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(123, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(123, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(123, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(123, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(124, 2, xls.AddFormat(fmt));
            xls.SetCellValue(124, 2, "Porcentaje afirmativo");

            fmt = xls.GetCellVisibleFormatDef(124, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(124, 4, xls.AddFormat(fmt));
            xls.SetCellValue(124, 4, "Nota 3: Para más de un prestamo en el mismo año sumar los prestamos. Tomar el año"
            + " de devolucion como el año medio (en caso de ser distinto años)");

            fmt = xls.GetCellVisibleFormatDef(124, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(124, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(124, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(124, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(124, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(124, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(124, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(124, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(124, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(125, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(125, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 28;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 38;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.Automatic;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(125, 1, new TRichString("Recibio algún prestamo para  inversión por parte de un Banco u otro prestamista?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(125, 1, "Recibio alg&uacute;n prestamo para&nbsp;<font color = 'blue'>&nbsp;inversi&oacute;n</font>&nbsp;por"
            //+" parte de un Banco u otro prestamista?")


    fmt = xls.GetCellVisibleFormatDef(125, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(125, 2, xls.AddFormat(fmt));
            xls.SetCellValue(125, 2, new TFormula("='Inputs advanced'!F240"));

            fmt = xls.GetCellVisibleFormatDef(125, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(125, 3, xls.AddFormat(fmt));
            xls.SetCellValue(125, 3, "En que Año?:                                           ");

            fmt = xls.GetCellVisibleFormatDef(125, 4);
            fmt.WrapText = true;
            xls.SetCellFormat(125, 4, xls.AddFormat(fmt));
            xls.SetCellValue(125, 4, "Marque P (mayuscula) en el año en que recibio el prestamo");

            fmt = xls.GetCellVisibleFormatDef(125, 5);
            fmt.WrapText = true;
            xls.SetCellFormat(125, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(125, 6);
            fmt.WrapText = true;
            xls.SetCellFormat(125, 6, xls.AddFormat(fmt));
            xls.SetCellValue(125, 6, "Cantidad en moneda local");

            fmt = xls.GetCellVisibleFormatDef(125, 7);
            fmt.WrapText = true;
            xls.SetCellFormat(125, 7, xls.AddFormat(fmt));
            xls.SetCellValue(125, 7, "Cuando termina o terminó de pagar el prestamo? (Marque X mayuscula)");

            fmt = xls.GetCellVisibleFormatDef(125, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(125, 9, xls.AddFormat(fmt));
            xls.SetCellValue(125, 9, "Resumen Prestamos");

            fmt = xls.GetCellVisibleFormatDef(125, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(125, 10, xls.AddFormat(fmt));
            xls.SetCellValue(125, 10, "Resumen Pagos");

            fmt = xls.GetCellVisibleFormatDef(125, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(125, 11, xls.AddFormat(fmt));
            xls.SetCellValue(125, 11, "Valor prestamo");

            fmt = xls.GetCellVisibleFormatDef(125, 12);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.WrapText = true;
            xls.SetCellFormat(125, 12, xls.AddFormat(fmt));
            xls.SetCellValue(125, 12, "Cual es la tasa de interés annual de dicho prestamo");

            fmt = xls.GetCellVisibleFormatDef(125, 13);
            fmt.WrapText = true;
            xls.SetCellFormat(125, 13, xls.AddFormat(fmt));
            xls.SetCellValue(125, 13, "Cuanto pagó en total por ese prestamo?");

            fmt = xls.GetCellVisibleFormatDef(125, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 14, xls.AddFormat(fmt));
            xls.SetCellValue(125, 14, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(125, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 15, xls.AddFormat(fmt));
            xls.SetCellValue(125, 15, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(125, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 16, xls.AddFormat(fmt));
            xls.SetCellValue(125, 16, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(125, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 17, xls.AddFormat(fmt));
            xls.SetCellValue(125, 17, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(125, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 18, xls.AddFormat(fmt));
            xls.SetCellValue(125, 18, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(125, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 19, xls.AddFormat(fmt));
            xls.SetCellValue(125, 19, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(125, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 20, xls.AddFormat(fmt));
            xls.SetCellValue(125, 20, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(125, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 21, xls.AddFormat(fmt));
            xls.SetCellValue(125, 21, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(125, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(125, 22, xls.AddFormat(fmt));
            xls.SetCellValue(125, 22, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(126, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 1, xls.AddFormat(fmt));
            xls.SetCellValue(126, 1, "Tiempo del prestamo");

            fmt = xls.GetCellVisibleFormatDef(126, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 2, xls.AddFormat(fmt));
            xls.SetCellValue(126, 2, new TFormula("='Inputs advanced'!F245"));
            xls.SetCellValue(126, 3, "Preparacion terreno (Año 0) ");

            fmt = xls.GetCellVisibleFormatDef(126, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 4, xls.AddFormat(fmt));
            xls.SetCellValue(126, 4, "X");

            fmt = xls.GetCellVisibleFormatDef(126, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(126, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 6, xls.AddFormat(fmt));
            xls.SetCellValue(126, 6, new TFormula("=B127"));

            fmt = xls.GetCellVisibleFormatDef(126, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 7, xls.AddFormat(fmt));
            xls.SetCellValue(126, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(126, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 8, xls.AddFormat(fmt));
            xls.SetCellValue(126, 8, 0);

            fmt = xls.GetCellVisibleFormatDef(126, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 9, xls.AddFormat(fmt));
            xls.SetCellValue(126, 9, "X");

            fmt = xls.GetCellVisibleFormatDef(126, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 10, xls.AddFormat(fmt));
            xls.SetCellValue(126, 10, "Y");

            fmt = xls.GetCellVisibleFormatDef(126, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 11, xls.AddFormat(fmt));
            xls.SetCellValue(126, 11, new TFormula("=IF($B$131=1,VLOOKUP(I126,$D$126:$F$134,2,FALSE),0)"));

            fmt = xls.GetCellVisibleFormatDef(126, 12);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 12, xls.AddFormat(fmt));
            xls.SetCellValue(126, 12, new TFormula("=B128/100"));

            fmt = xls.GetCellVisibleFormatDef(126, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 13, xls.AddFormat(fmt));
            xls.SetCellValue(126, 13, new TFormula("=B127*(1+(L126*B126))"));

            fmt = xls.GetCellVisibleFormatDef(126, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 14, xls.AddFormat(fmt));
            xls.SetCellValue(126, 14, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 15, xls.AddFormat(fmt));
            xls.SetCellValue(126, 15, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 16, xls.AddFormat(fmt));
            xls.SetCellValue(126, 16, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 17, xls.AddFormat(fmt));
            xls.SetCellValue(126, 17, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 18, xls.AddFormat(fmt));
            xls.SetCellValue(126, 18, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 19);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 19, xls.AddFormat(fmt));
            xls.SetCellValue(126, 19, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 20);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 20, xls.AddFormat(fmt));
            xls.SetCellValue(126, 20, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 21);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 21, xls.AddFormat(fmt));
            xls.SetCellValue(126, 21, new TFormula("=$M126/9"));

            fmt = xls.GetCellVisibleFormatDef(126, 22);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(126, 22, xls.AddFormat(fmt));
            xls.SetCellValue(126, 22, new TFormula("=$M126/9"));
            xls.SetCellValue(126, 24, "Año Prestamo");

            fmt = xls.GetCellVisibleFormatDef(126, 25);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 25, xls.AddFormat(fmt));
            xls.SetCellValue(126, 25, "X");

            fmt = xls.GetCellVisibleFormatDef(126, 26);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 26, xls.AddFormat(fmt));
            xls.SetCellValue(126, 26, new TFormula("=IF(B131=1,VLOOKUP(Y126,D126:H134,4,FALSE),0)"));

            fmt = xls.GetCellVisibleFormatDef(126, 27);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(126, 27, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(127, 1, xls.AddFormat(fmt));
            xls.SetCellValue(127, 1, "Monto en moneda local");

            fmt = xls.GetCellVisibleFormatDef(127, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(127, 2, xls.AddFormat(fmt));
            xls.SetCellValue(127, 2, new TFormula("='Inputs advanced'!F243"));

            fmt = xls.GetCellVisibleFormatDef(127, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(127, 3, xls.AddFormat(fmt));
            xls.SetCellValue(127, 3, "Fertilización y control de plagas (Año 1)");

            fmt = xls.GetCellVisibleFormatDef(127, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 4, xls.AddFormat(fmt));
            xls.SetCellValue(127, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(127, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(127, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 7, xls.AddFormat(fmt));
            xls.SetCellValue(127, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(127, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 8, xls.AddFormat(fmt));
            xls.SetCellValue(127, 8, 1);

            fmt = xls.GetCellVisibleFormatDef(127, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 11);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(127, 13);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(127, 13, xls.AddFormat(fmt));
            xls.SetCellValue(127, 24, "Año Pago");

            fmt = xls.GetCellVisibleFormatDef(127, 25);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 25, xls.AddFormat(fmt));
            xls.SetCellValue(127, 25, "Y");

            fmt = xls.GetCellVisibleFormatDef(127, 26);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 26, xls.AddFormat(fmt));
            xls.SetCellValue(127, 26, new TFormula("=IF(B131=1,VLOOKUP(Y127,G126:H134,2,FALSE),0)"));

            fmt = xls.GetCellVisibleFormatDef(127, 27);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(127, 27, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(128, 1, xls.AddFormat(fmt));
            xls.SetCellValue(128, 1, "Tasa de interes Anual");

            fmt = xls.GetCellVisibleFormatDef(128, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(128, 2, xls.AddFormat(fmt));
            xls.SetCellValue(128, 2, new TFormula("='Inputs advanced'!F249"));

            fmt = xls.GetCellVisibleFormatDef(128, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(128, 3, xls.AddFormat(fmt));
            xls.SetCellValue(128, 3, "Cosecha y Postcosecha (Año 2)");

            fmt = xls.GetCellVisibleFormatDef(128, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(128, 4, xls.AddFormat(fmt));
            xls.SetCellValue(128, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(128, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(128, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(128, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(128, 7, xls.AddFormat(fmt));
            xls.SetCellValue(128, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(128, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(128, 8, xls.AddFormat(fmt));
            xls.SetCellValue(128, 8, 2);

            fmt = xls.GetCellVisibleFormatDef(128, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(128, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(128, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 11);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(128, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(128, 13);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(128, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(129, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(129, 3, xls.AddFormat(fmt));
            xls.SetCellValue(129, 3, "Cosecha y Postcosecha (Año 3)");

            fmt = xls.GetCellVisibleFormatDef(129, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(129, 4, xls.AddFormat(fmt));
            xls.SetCellValue(129, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(129, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(129, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(129, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(129, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(129, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(129, 7, xls.AddFormat(fmt));
            xls.SetCellValue(129, 7, "Y");

            fmt = xls.GetCellVisibleFormatDef(129, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(129, 8, xls.AddFormat(fmt));
            xls.SetCellValue(129, 8, 3);

            fmt = xls.GetCellVisibleFormatDef(129, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(129, 9, xls.AddFormat(fmt));
            xls.SetCellValue(129, 25, "Prestamo a cuantos años:");
            xls.SetCellValue(129, 26, new TFormula("=(Z127-Z126)+1"));

            fmt = xls.GetCellVisibleFormatDef(130, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(130, 2, xls.AddFormat(fmt));
            xls.SetCellValue(130, 2, "Si = 1");

            fmt = xls.GetCellVisibleFormatDef(130, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(130, 3, xls.AddFormat(fmt));
            xls.SetCellValue(130, 3, "Cosecha y Postcosecha (Año 4)");

            fmt = xls.GetCellVisibleFormatDef(130, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(130, 4, xls.AddFormat(fmt));
            xls.SetCellValue(130, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(130, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(130, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(130, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(130, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(130, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(130, 7, xls.AddFormat(fmt));
            xls.SetCellValue(130, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(130, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(130, 8, xls.AddFormat(fmt));
            xls.SetCellValue(130, 8, 4);

            fmt = xls.GetCellVisibleFormatDef(130, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(130, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(131, 2);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(131, 2, xls.AddFormat(fmt));
            xls.SetCellValue(131, 2, new TFormula("=IF(B125>=0.2,1,0)"));

            fmt = xls.GetCellVisibleFormatDef(131, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(131, 3, xls.AddFormat(fmt));
            xls.SetCellValue(131, 3, "Cosecha y Postcosecha (Año 5)");

            fmt = xls.GetCellVisibleFormatDef(131, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(131, 4, xls.AddFormat(fmt));
            xls.SetCellValue(131, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(131, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(131, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(131, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(131, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(131, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(131, 7, xls.AddFormat(fmt));
            xls.SetCellValue(131, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(131, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(131, 8, xls.AddFormat(fmt));
            xls.SetCellValue(131, 8, 5);

            fmt = xls.GetCellVisibleFormatDef(131, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(131, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(131, 13);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(131, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 3);
            fmt.WrapText = true;
            xls.SetCellFormat(132, 3, xls.AddFormat(fmt));
            xls.SetCellValue(132, 3, "Cosecha y Postcosecha (Año 6)");

            fmt = xls.GetCellVisibleFormatDef(132, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 4, xls.AddFormat(fmt));
            xls.SetCellValue(132, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(132, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(132, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 7, xls.AddFormat(fmt));
            xls.SetCellValue(132, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(132, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 8, xls.AddFormat(fmt));
            xls.SetCellValue(132, 8, 6);

            fmt = xls.GetCellVisibleFormatDef(132, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 9, xls.AddFormat(fmt));
            xls.SetCellValue(132, 25, "Tasa implicita anual");

            fmt = xls.GetCellVisibleFormatDef(132, 26);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(132, 26, xls.AddFormat(fmt));
            xls.SetCellValue(132, 26, new TFormula("=IF(B131=1,((M126-K126)/Z129)/K126,0)"));

            fmt = xls.GetCellVisibleFormatDef(133, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(133, 3, xls.AddFormat(fmt));
            xls.SetCellValue(133, 3, "Cosecha y Postcosecha (Año 7)");

            fmt = xls.GetCellVisibleFormatDef(133, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(133, 4, xls.AddFormat(fmt));
            xls.SetCellValue(133, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(133, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(133, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(133, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(133, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(133, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(133, 7, xls.AddFormat(fmt));
            xls.SetCellValue(133, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(133, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(133, 8, xls.AddFormat(fmt));
            xls.SetCellValue(133, 8, 7);

            fmt = xls.GetCellVisibleFormatDef(133, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(133, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(134, 3, xls.AddFormat(fmt));
            xls.SetCellValue(134, 3, "Cosecha y Postcosecha (Año 8)");

            fmt = xls.GetCellVisibleFormatDef(134, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(134, 4, xls.AddFormat(fmt));
            xls.SetCellValue(134, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(134, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(134, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(134, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(134, 7, xls.AddFormat(fmt));
            xls.SetCellValue(134, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(134, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(134, 8, xls.AddFormat(fmt));
            xls.SetCellValue(134, 8, 8);

            fmt = xls.GetCellVisibleFormatDef(134, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(134, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(135, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(135, 2, xls.AddFormat(fmt));
            xls.SetCellValue(135, 2, "Total Pago Prestamos Otros:");

            fmt = xls.GetCellVisibleFormatDef(135, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(135, 13, xls.AddFormat(fmt));
            xls.SetCellValue(135, 13, "Total Pago Prestamos Otros.");

            fmt = xls.GetCellVisibleFormatDef(135, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 14, xls.AddFormat(fmt));
            xls.SetCellValue(135, 14, new TFormula("=SUM(N126:N134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 15, xls.AddFormat(fmt));
            xls.SetCellValue(135, 15, new TFormula("=SUM(O126:O134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 16, xls.AddFormat(fmt));
            xls.SetCellValue(135, 16, new TFormula("=SUM(P126:P134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 17, xls.AddFormat(fmt));
            xls.SetCellValue(135, 17, new TFormula("=SUM(Q126:Q134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 18, xls.AddFormat(fmt));
            xls.SetCellValue(135, 18, new TFormula("=SUM(R126:R134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 19, xls.AddFormat(fmt));
            xls.SetCellValue(135, 19, new TFormula("=SUM(S126:S134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 20, xls.AddFormat(fmt));
            xls.SetCellValue(135, 20, new TFormula("=SUM(T126:T134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 21, xls.AddFormat(fmt));
            xls.SetCellValue(135, 21, new TFormula("=SUM(U126:U134)"));

            fmt = xls.GetCellVisibleFormatDef(135, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(135, 22, xls.AddFormat(fmt));
            xls.SetCellValue(135, 22, new TFormula("=SUM(V126:V134)"));

            fmt = xls.GetCellVisibleFormatDef(136, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(136, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(136, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(136, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(136, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(137, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 13);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(137, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 16);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 17);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 18);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 19);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 20);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 21);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(137, 22);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(137, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(139, 1);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(139, 1, xls.AddFormat(fmt));
            xls.SetCellValue(139, 1, "Productividad");

            fmt = xls.GetCellVisibleFormatDef(141, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(141, 8, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 5;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(141, 8, new TRichString("Nota: Para productividad esta es al final la variable objetivo", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(141, 8, "<b>Nota:</b>&nbsp;Para productividad esta es al final la variable objetivo")


            fmt = xls.GetCellVisibleFormatDef(141, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(141, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(141, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(141, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(142, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(142, 1, xls.AddFormat(fmt));
            xls.SetCellValue(142, 1, "Productividad - Café Pergamino Seco");

            fmt = xls.GetCellVisibleFormatDef(142, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(142, 8, xls.AddFormat(fmt));
            xls.SetCellValue(142, 8, "AJUSTAR EL PROMEDIO MANUALMENTE CONFORME A LOS DATOS DISPONIBLES");

            fmt = xls.GetCellVisibleFormatDef(142, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(142, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(142, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(142, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(143, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(143, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[9];
            Runs[0].FirstChar = 8;
            fnt = xls.GetDefaultFont;
            fnt.Color = TExcelColor.FromTheme(TThemeColor.Accent1);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 17;
            fnt = xls.GetDefaultFont;
            fnt.Style = TFlxFontStyles.Bold;
            Runs[1].FontIndex = xls.AddFont(fnt);
            Runs[2].FirstChar = 18;
            fnt = xls.GetDefaultFont;
            Runs[2].FontIndex = xls.AddFont(fnt);
            Runs[3].FirstChar = 26;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[3].FontIndex = xls.AddFont(fnt);
            Runs[4].FirstChar = 40;
            fnt = xls.GetDefaultFont;
            Runs[4].FontIndex = xls.AddFont(fnt);
            Runs[5].FirstChar = 52;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Underline = TFlxUnderline.Single;
            Runs[5].FontIndex = xls.AddFont(fnt);
            Runs[6].FirstChar = 61;
            fnt = xls.GetDefaultFont;
            Runs[6].FontIndex = xls.AddFont(fnt);
            Runs[7].FirstChar = 70;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[7].FontIndex = xls.AddFont(fnt);
            Runs[8].FirstChar = 73;
            fnt = xls.GetDefaultFont;
            Runs[8].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(143, 1, new TRichString("Cuantos quintales de café pergamino seco espera UD. por árbol en cada año?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(143, 1, "Cuantos&nbsp;<font color = '#4f81bd'>quintales</font><b>&nbsp;</b>de caf&eacute;&nbsp;<font"
            //+" color = 'blue'><b>pergamino seco</b></font>&nbsp;espera UD.&nbsp;<font color = 'blue'><b><u>por"
            //+ " &aacute;rbol</u></b></font>&nbsp;en cada&nbsp;<font color = 'blue'><b>a&ntilde;o</b></font>?")


    fmt = xls.GetCellVisibleFormatDef(143, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(143, 4, xls.AddFormat(fmt));
            xls.SetCellValue(143, 4, "Kilos por hectaria por año (conforme a proporcion reportada de la respectiva variedad"
            + " en una hectarea)");

            fmt = xls.GetCellVisibleFormatDef(143, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(143, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(143, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(143, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(143, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(143, 8, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 9;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 18;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Arial";
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(143, 8, new TRichString("¿Cuántas Quintales de PERGAMINO SECO recoge POR HECTÁREA?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(143, 8, "&iquest;Cu&aacute;ntas&nbsp;<font color = 'black'><b>Quintales</b></font>&nbsp;de"
            //+" PERGAMINO SECO recoge POR HECT&Aacute;REA?")


    fmt = xls.GetCellVisibleFormatDef(143, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(143, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(143, 13);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(143, 13, xls.AddFormat(fmt));
            xls.SetCellValue(143, 13, "EN QUINTALES");

            fmt = xls.GetCellVisibleFormatDef(144, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 1, xls.AddFormat(fmt));
            xls.SetCellValue(144, 1, "Año");

            fmt = xls.GetCellVisibleFormatDef(144, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 2, xls.AddFormat(fmt));
            xls.SetCellValue(144, 2, new TFormula("=A18"));

            fmt = xls.GetCellVisibleFormatDef(144, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 3, xls.AddFormat(fmt));
            xls.SetCellValue(144, 3, "Borbon");

            fmt = xls.GetCellVisibleFormatDef(144, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 4, xls.AddFormat(fmt));
            xls.SetCellValue(144, 4, "Caturra");

            fmt = xls.GetCellVisibleFormatDef(144, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 5, xls.AddFormat(fmt));
            xls.SetCellValue(144, 5, "Mondonovo");

            fmt = xls.GetCellVisibleFormatDef(144, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 6, xls.AddFormat(fmt));
            xls.SetCellValue(144, 6, "Marago");

            fmt = xls.GetCellVisibleFormatDef(144, 7);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(144, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(144, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 9);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(144, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 10);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(144, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 11);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            xls.SetCellFormat(144, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            xls.SetCellFormat(144, 12, xls.AddFormat(fmt));
            xls.SetCellValue(144, 12, "TOTAL PRODUCIDO");

            fmt = xls.GetCellVisibleFormatDef(144, 13);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            xls.SetCellFormat(144, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 14);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(144, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(144, 15);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(144, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(145, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(145, 1, xls.AddFormat(fmt));
            xls.SetCellValue(145, 1, 0);

            fmt = xls.GetCellVisibleFormatDef(145, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(145, 2, xls.AddFormat(fmt));
            xls.SetCellValue(145, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(145, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(145, 3, xls.AddFormat(fmt));
            xls.SetCellValue(145, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(145, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(145, 4, xls.AddFormat(fmt));
            xls.SetCellValue(145, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(145, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(145, 5, xls.AddFormat(fmt));
            xls.SetCellValue(145, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(145, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(145, 6, xls.AddFormat(fmt));
            xls.SetCellValue(145, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(145, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(145, 7, xls.AddFormat(fmt));
            xls.SetCellValue(145, 7, 0);

            fmt = xls.GetCellVisibleFormatDef(145, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(145, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(145, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(145, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(145, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(145, 12, xls.AddFormat(fmt));
            xls.SetCellValue(145, 12, new TFormula("=SUM(H145:K145)"));

            fmt = xls.GetCellVisibleFormatDef(145, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(145, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(145, 17);
            fmt.Format = "0";
            xls.SetCellFormat(145, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(146, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(146, 1, xls.AddFormat(fmt));
            xls.SetCellValue(146, 1, 1);

            fmt = xls.GetCellVisibleFormatDef(146, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(146, 2, xls.AddFormat(fmt));
            xls.SetCellValue(146, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(146, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(146, 3, xls.AddFormat(fmt));
            xls.SetCellValue(146, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(146, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(146, 4, xls.AddFormat(fmt));
            xls.SetCellValue(146, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(146, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(146, 5, xls.AddFormat(fmt));
            xls.SetCellValue(146, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(146, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(146, 6, xls.AddFormat(fmt));
            xls.SetCellValue(146, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(146, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(146, 7, xls.AddFormat(fmt));
            xls.SetCellValue(146, 7, 1);

            fmt = xls.GetCellVisibleFormatDef(146, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(146, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(146, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(146, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(146, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(146, 12, xls.AddFormat(fmt));
            xls.SetCellValue(146, 12, new TFormula("=SUM(H146:K146)"));

            fmt = xls.GetCellVisibleFormatDef(146, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(146, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(146, 17);
            fmt.Format = "0";
            xls.SetCellFormat(146, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(147, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(147, 1, xls.AddFormat(fmt));
            xls.SetCellValue(147, 1, 2);

            fmt = xls.GetCellVisibleFormatDef(147, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(147, 2, xls.AddFormat(fmt));
            xls.SetCellValue(147, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(147, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(147, 3, xls.AddFormat(fmt));
            xls.SetCellValue(147, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(147, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(147, 4, xls.AddFormat(fmt));
            xls.SetCellValue(147, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(147, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(147, 5, xls.AddFormat(fmt));
            xls.SetCellValue(147, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(147, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(147, 6, xls.AddFormat(fmt));
            xls.SetCellValue(147, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(147, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(147, 7, xls.AddFormat(fmt));
            xls.SetCellValue(147, 7, 2);

            fmt = xls.GetCellVisibleFormatDef(147, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(147, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(147, 9);
            fmt.Format = "0.00";
            xls.SetCellFormat(147, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(147, 10);
            fmt.Format = "0.00";
            xls.SetCellFormat(147, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(147, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(147, 12, xls.AddFormat(fmt));
            xls.SetCellValue(147, 12, new TFormula("='Inputs advanced'!F78"));

            fmt = xls.GetCellVisibleFormatDef(147, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(147, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(147, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(147, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(147, 17);
            fmt.Format = "0";
            xls.SetCellFormat(147, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(148, 1, xls.AddFormat(fmt));
            xls.SetCellValue(148, 1, 3);

            fmt = xls.GetCellVisibleFormatDef(148, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(148, 2, xls.AddFormat(fmt));
            xls.SetCellValue(148, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(148, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(148, 3, xls.AddFormat(fmt));
            xls.SetCellValue(148, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(148, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(148, 4, xls.AddFormat(fmt));
            xls.SetCellValue(148, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(148, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(148, 5, xls.AddFormat(fmt));
            xls.SetCellValue(148, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(148, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(148, 6, xls.AddFormat(fmt));
            xls.SetCellValue(148, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(148, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(148, 7, xls.AddFormat(fmt));
            xls.SetCellValue(148, 7, 3);

            fmt = xls.GetCellVisibleFormatDef(148, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(148, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 9);
            fmt.Format = "0.00";
            xls.SetCellFormat(148, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 10);
            fmt.Format = "0.00";
            xls.SetCellFormat(148, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(148, 12, xls.AddFormat(fmt));
            xls.SetCellValue(148, 12, new TFormula("='Inputs advanced'!F78"));

            fmt = xls.GetCellVisibleFormatDef(148, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(148, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(148, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(148, 17);
            fmt.Format = "0";
            xls.SetCellFormat(148, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(149, 1, xls.AddFormat(fmt));
            xls.SetCellValue(149, 1, 4);

            fmt = xls.GetCellVisibleFormatDef(149, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(149, 2, xls.AddFormat(fmt));
            xls.SetCellValue(149, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(149, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(149, 3, xls.AddFormat(fmt));
            xls.SetCellValue(149, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(149, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(149, 4, xls.AddFormat(fmt));
            xls.SetCellValue(149, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(149, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(149, 5, xls.AddFormat(fmt));
            xls.SetCellValue(149, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(149, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(149, 6, xls.AddFormat(fmt));
            xls.SetCellValue(149, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(149, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(149, 7, xls.AddFormat(fmt));
            xls.SetCellValue(149, 7, 4);

            fmt = xls.GetCellVisibleFormatDef(149, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(149, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 9);
            fmt.Format = "0.00";
            xls.SetCellFormat(149, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 10);
            fmt.Format = "0.00";
            xls.SetCellFormat(149, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(149, 12, xls.AddFormat(fmt));
            xls.SetCellValue(149, 12, new TFormula("='Inputs advanced'!F105"));

            fmt = xls.GetCellVisibleFormatDef(149, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(149, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(149, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(149, 17);
            fmt.Format = "0";
            xls.SetCellFormat(149, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(150, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(150, 1, xls.AddFormat(fmt));
            xls.SetCellValue(150, 1, 5);

            fmt = xls.GetCellVisibleFormatDef(150, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(150, 2, xls.AddFormat(fmt));
            xls.SetCellValue(150, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(150, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(150, 3, xls.AddFormat(fmt));
            xls.SetCellValue(150, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(150, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(150, 4, xls.AddFormat(fmt));
            xls.SetCellValue(150, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(150, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(150, 5, xls.AddFormat(fmt));
            xls.SetCellValue(150, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(150, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(150, 6, xls.AddFormat(fmt));
            xls.SetCellValue(150, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(150, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(150, 7, xls.AddFormat(fmt));
            xls.SetCellValue(150, 7, 5);

            fmt = xls.GetCellVisibleFormatDef(150, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(150, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(150, 9);
            fmt.Format = "0.00";
            xls.SetCellFormat(150, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(150, 10);
            fmt.Format = "0.00";
            xls.SetCellFormat(150, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(150, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(150, 12, xls.AddFormat(fmt));
            xls.SetCellValue(150, 12, new TFormula("='Inputs advanced'!F105"));

            fmt = xls.GetCellVisibleFormatDef(150, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(150, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(150, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(150, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(150, 17);
            fmt.Format = "0";
            xls.SetCellFormat(150, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(151, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(151, 1, xls.AddFormat(fmt));
            xls.SetCellValue(151, 1, 6);

            fmt = xls.GetCellVisibleFormatDef(151, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(151, 2, xls.AddFormat(fmt));
            xls.SetCellValue(151, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(151, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(151, 3, xls.AddFormat(fmt));
            xls.SetCellValue(151, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(151, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(151, 4, xls.AddFormat(fmt));
            xls.SetCellValue(151, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(151, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(151, 5, xls.AddFormat(fmt));
            xls.SetCellValue(151, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(151, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(151, 6, xls.AddFormat(fmt));
            xls.SetCellValue(151, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(151, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(151, 7, xls.AddFormat(fmt));
            xls.SetCellValue(151, 7, 6);

            fmt = xls.GetCellVisibleFormatDef(151, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(151, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(151, 9);
            fmt.Format = "0.00";
            xls.SetCellFormat(151, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(151, 10);
            fmt.Format = "0.00";
            xls.SetCellFormat(151, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(151, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(151, 12, xls.AddFormat(fmt));
            xls.SetCellValue(151, 12, new TFormula("='Inputs advanced'!F132"));

            fmt = xls.GetCellVisibleFormatDef(151, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(151, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(151, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(151, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(151, 17);
            fmt.Format = "0";
            xls.SetCellFormat(151, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(152, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(152, 1, xls.AddFormat(fmt));
            xls.SetCellValue(152, 1, 7);

            fmt = xls.GetCellVisibleFormatDef(152, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(152, 2, xls.AddFormat(fmt));
            xls.SetCellValue(152, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(152, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(152, 3, xls.AddFormat(fmt));
            xls.SetCellValue(152, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(152, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(152, 4, xls.AddFormat(fmt));
            xls.SetCellValue(152, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(152, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(152, 5, xls.AddFormat(fmt));
            xls.SetCellValue(152, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(152, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(152, 6, xls.AddFormat(fmt));
            xls.SetCellValue(152, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(152, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(152, 7, xls.AddFormat(fmt));
            xls.SetCellValue(152, 7, 7);

            fmt = xls.GetCellVisibleFormatDef(152, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(152, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(152, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(152, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(152, 10);
            fmt.Format = "0.00";
            xls.SetCellFormat(152, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(152, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(152, 12, xls.AddFormat(fmt));
            xls.SetCellValue(152, 12, new TFormula("='Inputs advanced'!F132"));

            fmt = xls.GetCellVisibleFormatDef(152, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(152, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(152, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(152, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(152, 17);
            fmt.Format = "0";
            xls.SetCellFormat(152, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(153, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(153, 1, xls.AddFormat(fmt));
            xls.SetCellValue(153, 1, 8);

            fmt = xls.GetCellVisibleFormatDef(153, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(153, 2, xls.AddFormat(fmt));
            xls.SetCellValue(153, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(153, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(153, 3, xls.AddFormat(fmt));
            xls.SetCellValue(153, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(153, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(153, 4, xls.AddFormat(fmt));
            xls.SetCellValue(153, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(153, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(153, 5, xls.AddFormat(fmt));
            xls.SetCellValue(153, 5, 0);

            fmt = xls.GetCellVisibleFormatDef(153, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(153, 6, xls.AddFormat(fmt));
            xls.SetCellValue(153, 6, 0);

            fmt = xls.GetCellVisibleFormatDef(153, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(153, 7, xls.AddFormat(fmt));
            xls.SetCellValue(153, 7, 8);

            fmt = xls.GetCellVisibleFormatDef(153, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(153, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(153, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0";
            xls.SetCellFormat(153, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(153, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(153, 12, xls.AddFormat(fmt));
            xls.SetCellValue(153, 12, new TFormula("='Inputs advanced'!F132"));

            fmt = xls.GetCellVisibleFormatDef(153, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(153, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(153, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(153, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(153, 17);
            fmt.Format = "0";
            xls.SetCellFormat(153, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(154, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(154, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(154, 6);
            fmt.Format = "0";
            xls.SetCellFormat(154, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(154, 7);
            fmt.Format = "0";
            xls.SetCellFormat(154, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(154, 8);
            fmt.Format = "0";
            xls.SetCellFormat(154, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(155, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 6);
            fmt.Format = "0";
            xls.SetCellFormat(155, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 7);
            fmt.Format = "0";
            xls.SetCellFormat(155, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 8);
            fmt.Format = "0";
            xls.SetCellFormat(155, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 11);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(155, 11, xls.AddFormat(fmt));
            xls.SetCellValue(155, 11, "Promedio");

            fmt = xls.GetCellVisibleFormatDef(155, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(155, 12, xls.AddFormat(fmt));
            xls.SetCellValue(155, 12, new TFormula("=AVERAGE(L147:L153)"));

            fmt = xls.GetCellVisibleFormatDef(155, 13);
            fmt.Format = "0.0";
            xls.SetCellFormat(155, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(155, 15);
            fmt.Format = "0";
            xls.SetCellFormat(155, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(156, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(156, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(156, 6);
            fmt.Format = "0";
            xls.SetCellFormat(156, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(156, 7);
            fmt.Format = "0";
            xls.SetCellFormat(156, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(156, 8);
            fmt.Format = "0";
            xls.SetCellFormat(156, 8, xls.AddFormat(fmt));
            xls.SetCellValue(156, 12, new TFormula("=L155*Conversiones!C14"));

            fmt = xls.GetCellVisibleFormatDef(157, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xE8, 0x1F, 0x19);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(157, 1, xls.AddFormat(fmt));
            xls.SetCellValue(157, 1, "ALTERNATIVAS");

            fmt = xls.GetCellVisibleFormatDef(157, 6);
            fmt.Format = "0";
            xls.SetCellFormat(157, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(157, 7);
            fmt.Format = "0";
            xls.SetCellFormat(157, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(157, 8);
            fmt.Format = "0";
            xls.SetCellFormat(157, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(157, 15);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(157, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(158, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(158, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[9];
            Runs[0].FirstChar = 8;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 14;
            fnt = xls.GetDefaultFont;
            fnt.Style = TFlxFontStyles.Bold;
            Runs[1].FontIndex = xls.AddFont(fnt);
            Runs[2].FirstChar = 15;
            fnt = xls.GetDefaultFont;
            Runs[2].FontIndex = xls.AddFont(fnt);
            Runs[3].FirstChar = 17;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            Runs[3].FontIndex = xls.AddFont(fnt);
            Runs[4].FirstChar = 35;
            fnt = xls.GetDefaultFont;
            Runs[4].FontIndex = xls.AddFont(fnt);
            Runs[5].FirstChar = 47;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Underline = TFlxUnderline.Single;
            Runs[5].FontIndex = xls.AddFont(fnt);
            Runs[6].FirstChar = 56;
            fnt = xls.GetDefaultFont;
            Runs[6].FontIndex = xls.AddFont(fnt);
            Runs[7].FirstChar = 65;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[7].FontIndex = xls.AddFont(fnt);
            Runs[8].FirstChar = 68;
            fnt = xls.GetDefaultFont;
            Runs[8].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(158, 1, new TRichString("Cuantos Libras de café cereza o uva espera UD. por arbol en cada año?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(158, 1, "Cuantos&nbsp;<font color = 'red'><b>Libras</b></font><b>&nbsp;</b>de<font color ="
            //+" 'blue'>&nbsp;caf&eacute; cereza o uva</font>&nbsp;espera UD.&nbsp;<font color = 'blue'><b><u>por"
            //+ " arbol</u></b></font>&nbsp;en cada&nbsp;<font color = 'blue'><b>a&ntilde;o</b></font>?")


    fmt = xls.GetCellVisibleFormatDef(158, 7);
            fmt.Format = "0";
            xls.SetCellFormat(158, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(158, 8);
            fmt.Format = "0";
            xls.SetCellFormat(158, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(158, 9);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(158, 9, xls.AddFormat(fmt));
            xls.SetCellValue(158, 9, "Quintales por manzana por año (conforme a proporcion reportada de la respectiva variedad"
            + " para un arbol)");

            fmt = xls.GetCellVisibleFormatDef(158, 10);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(158, 10, xls.AddFormat(fmt));
            xls.SetCellValue(158, 12, "Año mas representativo");

            fmt = xls.GetCellVisibleFormatDef(159, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 1, xls.AddFormat(fmt));
            xls.SetCellValue(159, 1, "Año");

            fmt = xls.GetCellVisibleFormatDef(159, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 2, xls.AddFormat(fmt));
            xls.SetCellValue(159, 2, new TFormula("=B144"));

            fmt = xls.GetCellVisibleFormatDef(159, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 3, xls.AddFormat(fmt));
            xls.SetCellValue(159, 3, new TFormula("=C144"));

            fmt = xls.GetCellVisibleFormatDef(159, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 4, xls.AddFormat(fmt));
            xls.SetCellValue(159, 4, new TFormula("=D144"));

            fmt = xls.GetCellVisibleFormatDef(159, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 5, xls.AddFormat(fmt));
            xls.SetCellValue(159, 5, new TFormula("=E144"));

            fmt = xls.GetCellVisibleFormatDef(159, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 6, xls.AddFormat(fmt));
            xls.SetCellValue(159, 6, new TFormula("=F144"));

            fmt = xls.GetCellVisibleFormatDef(159, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(159, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 8, xls.AddFormat(fmt));
            xls.SetCellValue(159, 8, "Media Otras Variedades");

            fmt = xls.GetCellVisibleFormatDef(159, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 9, xls.AddFormat(fmt));
            xls.SetCellValue(159, 9, "Catuai");

            fmt = xls.GetCellVisibleFormatDef(159, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 10, xls.AddFormat(fmt));
            xls.SetCellValue(159, 10, "Borbon");

            fmt = xls.GetCellVisibleFormatDef(159, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 11, xls.AddFormat(fmt));
            xls.SetCellValue(159, 11, "Icatu");

            fmt = xls.GetCellVisibleFormatDef(159, 12);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 12, xls.AddFormat(fmt));
            xls.SetCellValue(159, 12, "Lempira");

            fmt = xls.GetCellVisibleFormatDef(159, 13);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 13, xls.AddFormat(fmt));
            xls.SetCellValue(159, 13, "Icafe 90");

            fmt = xls.GetCellVisibleFormatDef(159, 14);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 14, xls.AddFormat(fmt));
            xls.SetCellValue(159, 14, "Media Otras Variedades");

            fmt = xls.GetCellVisibleFormatDef(159, 15);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(159, 15, xls.AddFormat(fmt));
            xls.SetCellValue(159, 15, "Media Total Ponderada");

            fmt = xls.GetCellVisibleFormatDef(160, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(160, 1, xls.AddFormat(fmt));
            xls.SetCellValue(160, 1, 0);

            fmt = xls.GetCellVisibleFormatDef(160, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(160, 2, xls.AddFormat(fmt));
            xls.SetCellValue(160, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(160, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(160, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(160, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(160, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(160, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(160, 7, xls.AddFormat(fmt));
            xls.SetCellValue(160, 7, 0);

            fmt = xls.GetCellVisibleFormatDef(160, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(160, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 9);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 10);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 11);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 12);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 13);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(160, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 16);
            fmt.Format = "0";
            xls.SetCellFormat(160, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(161, 1, xls.AddFormat(fmt));
            xls.SetCellValue(161, 1, 1);

            fmt = xls.GetCellVisibleFormatDef(161, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 2, xls.AddFormat(fmt));
            xls.SetCellValue(161, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(161, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 7, xls.AddFormat(fmt));
            xls.SetCellValue(161, 7, 0);

            fmt = xls.GetCellVisibleFormatDef(161, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(161, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(161, 16);
            fmt.Format = "0";
            xls.SetCellFormat(161, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(162, 1, xls.AddFormat(fmt));
            xls.SetCellValue(162, 1, 2);

            fmt = xls.GetCellVisibleFormatDef(162, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 2, xls.AddFormat(fmt));
            xls.SetCellValue(162, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 3, xls.AddFormat(fmt));
            xls.SetCellValue(162, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 4, xls.AddFormat(fmt));
            xls.SetCellValue(162, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 5, xls.AddFormat(fmt));
            xls.SetCellValue(162, 5, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 6, xls.AddFormat(fmt));
            xls.SetCellValue(162, 6, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 7, xls.AddFormat(fmt));
            xls.SetCellValue(162, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 8, xls.AddFormat(fmt));
            xls.SetCellValue(162, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 9, xls.AddFormat(fmt));
            xls.SetCellValue(162, 9, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 10, xls.AddFormat(fmt));
            xls.SetCellValue(162, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 11, xls.AddFormat(fmt));
            xls.SetCellValue(162, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 12, xls.AddFormat(fmt));
            xls.SetCellValue(162, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 13, xls.AddFormat(fmt));
            xls.SetCellValue(162, 13, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 14, xls.AddFormat(fmt));
            xls.SetCellValue(162, 14, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 15, xls.AddFormat(fmt));
            xls.SetCellValue(162, 15, ".");

            fmt = xls.GetCellVisibleFormatDef(162, 16);
            fmt.Format = "0";
            xls.SetCellFormat(162, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 20);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 20, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(162, 21);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(162, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(163, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(163, 1, xls.AddFormat(fmt));
            xls.SetCellValue(163, 1, 3);

            fmt = xls.GetCellVisibleFormatDef(163, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 2, xls.AddFormat(fmt));
            xls.SetCellValue(163, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 3, xls.AddFormat(fmt));
            xls.SetCellValue(163, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 4, xls.AddFormat(fmt));
            xls.SetCellValue(163, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 5, xls.AddFormat(fmt));
            xls.SetCellValue(163, 5, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 6, xls.AddFormat(fmt));
            xls.SetCellValue(163, 6, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 7, xls.AddFormat(fmt));
            xls.SetCellValue(163, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 8, xls.AddFormat(fmt));
            xls.SetCellValue(163, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 9, xls.AddFormat(fmt));
            xls.SetCellValue(163, 9, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 10, xls.AddFormat(fmt));
            xls.SetCellValue(163, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 11, xls.AddFormat(fmt));
            xls.SetCellValue(163, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 12, xls.AddFormat(fmt));
            xls.SetCellValue(163, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 13, xls.AddFormat(fmt));
            xls.SetCellValue(163, 13, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 14, xls.AddFormat(fmt));
            xls.SetCellValue(163, 14, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(163, 15, xls.AddFormat(fmt));
            xls.SetCellValue(163, 15, ".");

            fmt = xls.GetCellVisibleFormatDef(163, 16);
            fmt.Format = "0";
            xls.SetCellFormat(163, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(164, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(164, 1, xls.AddFormat(fmt));
            xls.SetCellValue(164, 1, 4);

            fmt = xls.GetCellVisibleFormatDef(164, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 2, xls.AddFormat(fmt));
            xls.SetCellValue(164, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 3, xls.AddFormat(fmt));
            xls.SetCellValue(164, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 4, xls.AddFormat(fmt));
            xls.SetCellValue(164, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 5, xls.AddFormat(fmt));
            xls.SetCellValue(164, 5, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 6, xls.AddFormat(fmt));
            xls.SetCellValue(164, 6, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 7, xls.AddFormat(fmt));
            xls.SetCellValue(164, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 8, xls.AddFormat(fmt));
            xls.SetCellValue(164, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 9, xls.AddFormat(fmt));
            xls.SetCellValue(164, 9, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 10, xls.AddFormat(fmt));
            xls.SetCellValue(164, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 11, xls.AddFormat(fmt));
            xls.SetCellValue(164, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 12, xls.AddFormat(fmt));
            xls.SetCellValue(164, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 13, xls.AddFormat(fmt));
            xls.SetCellValue(164, 13, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 14, xls.AddFormat(fmt));
            xls.SetCellValue(164, 14, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(164, 15, xls.AddFormat(fmt));
            xls.SetCellValue(164, 15, ".");

            fmt = xls.GetCellVisibleFormatDef(164, 16);
            fmt.Format = "0";
            xls.SetCellFormat(164, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(165, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(165, 1, xls.AddFormat(fmt));
            xls.SetCellValue(165, 1, 5);

            fmt = xls.GetCellVisibleFormatDef(165, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 2, xls.AddFormat(fmt));
            xls.SetCellValue(165, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 3, xls.AddFormat(fmt));
            xls.SetCellValue(165, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 4, xls.AddFormat(fmt));
            xls.SetCellValue(165, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 5, xls.AddFormat(fmt));
            xls.SetCellValue(165, 5, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 6, xls.AddFormat(fmt));
            xls.SetCellValue(165, 6, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 7, xls.AddFormat(fmt));
            xls.SetCellValue(165, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 8, xls.AddFormat(fmt));
            xls.SetCellValue(165, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 9, xls.AddFormat(fmt));
            xls.SetCellValue(165, 9, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 10, xls.AddFormat(fmt));
            xls.SetCellValue(165, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 11, xls.AddFormat(fmt));
            xls.SetCellValue(165, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 12, xls.AddFormat(fmt));
            xls.SetCellValue(165, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 13, xls.AddFormat(fmt));
            xls.SetCellValue(165, 13, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 14, xls.AddFormat(fmt));
            xls.SetCellValue(165, 14, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(165, 15, xls.AddFormat(fmt));
            xls.SetCellValue(165, 15, ".");

            fmt = xls.GetCellVisibleFormatDef(165, 16);
            fmt.Format = "0";
            xls.SetCellFormat(165, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(166, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(166, 1, xls.AddFormat(fmt));
            xls.SetCellValue(166, 1, 6);

            fmt = xls.GetCellVisibleFormatDef(166, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 2, xls.AddFormat(fmt));
            xls.SetCellValue(166, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 3, xls.AddFormat(fmt));
            xls.SetCellValue(166, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 4, xls.AddFormat(fmt));
            xls.SetCellValue(166, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 5, xls.AddFormat(fmt));
            xls.SetCellValue(166, 5, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 6, xls.AddFormat(fmt));
            xls.SetCellValue(166, 6, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 7, xls.AddFormat(fmt));
            xls.SetCellValue(166, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 8, xls.AddFormat(fmt));
            xls.SetCellValue(166, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 9, xls.AddFormat(fmt));
            xls.SetCellValue(166, 9, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 10, xls.AddFormat(fmt));
            xls.SetCellValue(166, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 11, xls.AddFormat(fmt));
            xls.SetCellValue(166, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 12, xls.AddFormat(fmt));
            xls.SetCellValue(166, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 13, xls.AddFormat(fmt));
            xls.SetCellValue(166, 13, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 14, xls.AddFormat(fmt));
            xls.SetCellValue(166, 14, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(166, 15, xls.AddFormat(fmt));
            xls.SetCellValue(166, 15, ".");

            fmt = xls.GetCellVisibleFormatDef(166, 16);
            fmt.Format = "0";
            xls.SetCellFormat(166, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(167, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(167, 1, xls.AddFormat(fmt));
            xls.SetCellValue(167, 1, 7);

            fmt = xls.GetCellVisibleFormatDef(167, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 2, xls.AddFormat(fmt));
            xls.SetCellValue(167, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 3, xls.AddFormat(fmt));
            xls.SetCellValue(167, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 4, xls.AddFormat(fmt));
            xls.SetCellValue(167, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 5, xls.AddFormat(fmt));
            xls.SetCellValue(167, 5, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 6, xls.AddFormat(fmt));
            xls.SetCellValue(167, 6, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 7, xls.AddFormat(fmt));
            xls.SetCellValue(167, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 8, xls.AddFormat(fmt));
            xls.SetCellValue(167, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 9, xls.AddFormat(fmt));
            xls.SetCellValue(167, 9, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 10, xls.AddFormat(fmt));
            xls.SetCellValue(167, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 11, xls.AddFormat(fmt));
            xls.SetCellValue(167, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 12, xls.AddFormat(fmt));
            xls.SetCellValue(167, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 13, xls.AddFormat(fmt));
            xls.SetCellValue(167, 13, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 14, xls.AddFormat(fmt));
            xls.SetCellValue(167, 14, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(167, 15, xls.AddFormat(fmt));
            xls.SetCellValue(167, 15, ".");

            fmt = xls.GetCellVisibleFormatDef(167, 16);
            fmt.Format = "0";
            xls.SetCellFormat(167, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 1, xls.AddFormat(fmt));
            xls.SetCellValue(168, 1, 8);

            fmt = xls.GetCellVisibleFormatDef(168, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 2, xls.AddFormat(fmt));
            xls.SetCellValue(168, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 3, xls.AddFormat(fmt));
            xls.SetCellValue(168, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 4, xls.AddFormat(fmt));
            xls.SetCellValue(168, 4, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 5, xls.AddFormat(fmt));
            xls.SetCellValue(168, 5, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 6, xls.AddFormat(fmt));
            xls.SetCellValue(168, 6, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 7, xls.AddFormat(fmt));
            xls.SetCellValue(168, 7, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 8, xls.AddFormat(fmt));
            xls.SetCellValue(168, 8, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 9, xls.AddFormat(fmt));
            xls.SetCellValue(168, 9, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 10, xls.AddFormat(fmt));
            xls.SetCellValue(168, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 11, xls.AddFormat(fmt));
            xls.SetCellValue(168, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 12, xls.AddFormat(fmt));
            xls.SetCellValue(168, 12, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 13, xls.AddFormat(fmt));
            xls.SetCellValue(168, 13, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 14, xls.AddFormat(fmt));
            xls.SetCellValue(168, 14, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 15);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(168, 15, xls.AddFormat(fmt));
            xls.SetCellValue(168, 15, ".");

            fmt = xls.GetCellVisibleFormatDef(168, 16);
            fmt.Format = "0";
            xls.SetCellFormat(168, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(169, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 6);
            fmt.Format = "0";
            xls.SetCellFormat(169, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 7);
            fmt.Format = "0";
            xls.SetCellFormat(169, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 8);
            fmt.Format = "0";
            xls.SetCellFormat(169, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(170, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[7];
            Runs[0].FirstChar = 7;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0xB0, 0x50);
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 17;
            fnt = xls.GetDefaultFont;
            fnt.Style = TFlxFontStyles.Bold;
            Runs[1].FontIndex = xls.AddFont(fnt);
            Runs[2].FirstChar = 18;
            fnt = xls.GetDefaultFont;
            Runs[2].FontIndex = xls.AddFont(fnt);
            Runs[3].FirstChar = 20;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            Runs[3].FontIndex = xls.AddFont(fnt);
            Runs[4].FirstChar = 35;
            fnt = xls.GetDefaultFont;
            Runs[4].FontIndex = xls.AddFont(fnt);
            Runs[5].FirstChar = 63;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fnt.Style = TFlxFontStyles.Bold;
            Runs[5].FontIndex = xls.AddFont(fnt);
            Runs[6].FirstChar = 66;
            fnt = xls.GetDefaultFont;
            Runs[6].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(170, 1, new TRichString("Cuantos Quintales de pergamino seco espera UD. manzana en cada año?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(170, 1, "Cuantos<font color = '#00b050'>&nbsp;Quintales</font><b>&nbsp;</b>de<font color ="
            //+" 'blue'>&nbsp;pergamino seco</font>&nbsp;espera UD. manzana en cada&nbsp;<font color"
            //+ " = 'blue'><b>a&ntilde;o</b></font>?")


    fmt = xls.GetCellVisibleFormatDef(170, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(170, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(170, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(170, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 7);
            fmt.Format = "0";
            xls.SetCellFormat(170, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 8);
            fmt.Format = "0";
            xls.SetCellFormat(170, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(171, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 1, xls.AddFormat(fmt));
            xls.SetCellValue(171, 1, "Año");

            fmt = xls.GetCellVisibleFormatDef(171, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 2, xls.AddFormat(fmt));
            xls.SetCellValue(171, 2, new TFormula("=B144"));

            fmt = xls.GetCellVisibleFormatDef(171, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 3, xls.AddFormat(fmt));
            xls.SetCellValue(171, 3, new TFormula("=C144"));

            fmt = xls.GetCellVisibleFormatDef(171, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 4, xls.AddFormat(fmt));
            xls.SetCellValue(171, 4, new TFormula("=D144"));

            fmt = xls.GetCellVisibleFormatDef(171, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 5, xls.AddFormat(fmt));
            xls.SetCellValue(171, 5, new TFormula("=E144"));

            fmt = xls.GetCellVisibleFormatDef(171, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 6, xls.AddFormat(fmt));
            xls.SetCellValue(171, 6, "Marago");

            fmt = xls.GetCellVisibleFormatDef(171, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(171, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 8, xls.AddFormat(fmt));
            xls.SetCellValue(171, 8, "Media Otras Variedades");

            fmt = xls.GetCellVisibleFormatDef(171, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(171, 9, xls.AddFormat(fmt));
            xls.SetCellValue(171, 9, "Media Total Ponderada");

            fmt = xls.GetCellVisibleFormatDef(171, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Format = "0";
            xls.SetCellFormat(171, 10, xls.AddFormat(fmt));
            xls.SetCellValue(171, 10, "Media En general Reportada I");

            fmt = xls.GetCellVisibleFormatDef(171, 11);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent3, -0.499984740745262);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Format = "0";
            xls.SetCellFormat(171, 11, xls.AddFormat(fmt));
            xls.SetCellValue(171, 11, "Media En general Reportada II");

            fmt = xls.GetCellVisibleFormatDef(172, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(172, 1, xls.AddFormat(fmt));
            xls.SetCellValue(172, 1, 0);

            fmt = xls.GetCellVisibleFormatDef(172, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 2, xls.AddFormat(fmt));
            xls.SetCellValue(172, 2, new TFormula("=B145"));

            fmt = xls.GetCellVisibleFormatDef(172, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 3, xls.AddFormat(fmt));
            xls.SetCellValue(172, 3, new TFormula("=C145"));

            fmt = xls.GetCellVisibleFormatDef(172, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 4, xls.AddFormat(fmt));
            xls.SetCellValue(172, 4, new TFormula("=D145"));

            fmt = xls.GetCellVisibleFormatDef(172, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 5, xls.AddFormat(fmt));
            xls.SetCellValue(172, 5, new TFormula("=E145"));

            fmt = xls.GetCellVisibleFormatDef(172, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 6, xls.AddFormat(fmt));
            xls.SetCellValue(172, 6, new TFormula("=F145"));

            fmt = xls.GetCellVisibleFormatDef(172, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(172, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 13);
            fmt.Format = "0";
            xls.SetCellFormat(172, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(172, 14);
            fmt.Format = "0";
            xls.SetCellFormat(172, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(173, 1, xls.AddFormat(fmt));
            xls.SetCellValue(173, 1, 1);

            fmt = xls.GetCellVisibleFormatDef(173, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 2, xls.AddFormat(fmt));
            xls.SetCellValue(173, 2, new TFormula("=B146"));

            fmt = xls.GetCellVisibleFormatDef(173, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 3, xls.AddFormat(fmt));
            xls.SetCellValue(173, 3, new TFormula("=C146"));

            fmt = xls.GetCellVisibleFormatDef(173, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 4, xls.AddFormat(fmt));
            xls.SetCellValue(173, 4, new TFormula("=D146"));

            fmt = xls.GetCellVisibleFormatDef(173, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 5, xls.AddFormat(fmt));
            xls.SetCellValue(173, 5, new TFormula("=E146"));

            fmt = xls.GetCellVisibleFormatDef(173, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 6, xls.AddFormat(fmt));
            xls.SetCellValue(173, 6, new TFormula("=F146"));

            fmt = xls.GetCellVisibleFormatDef(173, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(173, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 13);
            fmt.Format = "0";
            xls.SetCellFormat(173, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(173, 14);
            fmt.Format = "0";
            xls.SetCellFormat(173, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(174, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(174, 1, xls.AddFormat(fmt));
            xls.SetCellValue(174, 1, 2);

            fmt = xls.GetCellVisibleFormatDef(174, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 2, xls.AddFormat(fmt));
            xls.SetCellValue(174, 2, new TFormula("=B147"));

            fmt = xls.GetCellVisibleFormatDef(174, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 3, xls.AddFormat(fmt));
            xls.SetCellValue(174, 3, new TFormula("=C147"));

            fmt = xls.GetCellVisibleFormatDef(174, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 4, xls.AddFormat(fmt));
            xls.SetCellValue(174, 4, new TFormula("=D147"));

            fmt = xls.GetCellVisibleFormatDef(174, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 5, xls.AddFormat(fmt));
            xls.SetCellValue(174, 5, new TFormula("=E147"));

            fmt = xls.GetCellVisibleFormatDef(174, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 6, xls.AddFormat(fmt));
            xls.SetCellValue(174, 6, new TFormula("=F147"));

            fmt = xls.GetCellVisibleFormatDef(174, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(174, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(174, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 9, xls.AddFormat(fmt));
            xls.SetCellValue(174, 9, new TFormula("=(B174*$B$18+C174*$B$19+D174*$B$23+E174*$B$28+F174*$B$26)/0.973"));

            fmt = xls.GetCellVisibleFormatDef(174, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 10, xls.AddFormat(fmt));
            xls.SetCellValue(174, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(174, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(174, 11, xls.AddFormat(fmt));
            xls.SetCellValue(174, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(174, 13);
            fmt.Format = "0";
            xls.SetCellFormat(174, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(174, 14);
            fmt.Format = "0";
            xls.SetCellFormat(174, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(175, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(175, 1, xls.AddFormat(fmt));
            xls.SetCellValue(175, 1, 3);

            fmt = xls.GetCellVisibleFormatDef(175, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 2, xls.AddFormat(fmt));
            xls.SetCellValue(175, 2, new TFormula("=B148"));

            fmt = xls.GetCellVisibleFormatDef(175, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 3, xls.AddFormat(fmt));
            xls.SetCellValue(175, 3, new TFormula("=C148"));

            fmt = xls.GetCellVisibleFormatDef(175, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 4, xls.AddFormat(fmt));
            xls.SetCellValue(175, 4, new TFormula("=D148"));

            fmt = xls.GetCellVisibleFormatDef(175, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 5, xls.AddFormat(fmt));
            xls.SetCellValue(175, 5, new TFormula("=E148"));

            fmt = xls.GetCellVisibleFormatDef(175, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 6, xls.AddFormat(fmt));
            xls.SetCellValue(175, 6, new TFormula("=F148"));

            fmt = xls.GetCellVisibleFormatDef(175, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(175, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(175, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 9, xls.AddFormat(fmt));
            xls.SetCellValue(175, 9, new TFormula("=(B175*$B$18+C175*$B$19+D175*$B$23+E175*$B$28+F175*$B$26)/0.973"));

            fmt = xls.GetCellVisibleFormatDef(175, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 10, xls.AddFormat(fmt));
            xls.SetCellValue(175, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(175, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(175, 11, xls.AddFormat(fmt));
            xls.SetCellValue(175, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(175, 13);
            fmt.Format = "0";
            xls.SetCellFormat(175, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(175, 14);
            fmt.Format = "0";
            xls.SetCellFormat(175, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(176, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(176, 1, xls.AddFormat(fmt));
            xls.SetCellValue(176, 1, 4);

            fmt = xls.GetCellVisibleFormatDef(176, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 2, xls.AddFormat(fmt));
            xls.SetCellValue(176, 2, new TFormula("=B149"));

            fmt = xls.GetCellVisibleFormatDef(176, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 3, xls.AddFormat(fmt));
            xls.SetCellValue(176, 3, new TFormula("=C149"));

            fmt = xls.GetCellVisibleFormatDef(176, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 4, xls.AddFormat(fmt));
            xls.SetCellValue(176, 4, new TFormula("=D149"));

            fmt = xls.GetCellVisibleFormatDef(176, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 5, xls.AddFormat(fmt));
            xls.SetCellValue(176, 5, new TFormula("=E149"));

            fmt = xls.GetCellVisibleFormatDef(176, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 6, xls.AddFormat(fmt));
            xls.SetCellValue(176, 6, new TFormula("=F149"));

            fmt = xls.GetCellVisibleFormatDef(176, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(176, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(176, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 9, xls.AddFormat(fmt));
            xls.SetCellValue(176, 9, new TFormula("=(B176*$B$18+C176*$B$19+D176*$B$23+E176*$B$28+F176*$B$26)/0.973"));

            fmt = xls.GetCellVisibleFormatDef(176, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 10, xls.AddFormat(fmt));
            xls.SetCellValue(176, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(176, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(176, 11, xls.AddFormat(fmt));
            xls.SetCellValue(176, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(176, 13);
            fmt.Format = "0";
            xls.SetCellFormat(176, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(176, 14);
            fmt.Format = "0";
            xls.SetCellFormat(176, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(177, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(177, 1, xls.AddFormat(fmt));
            xls.SetCellValue(177, 1, 5);

            fmt = xls.GetCellVisibleFormatDef(177, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 2, xls.AddFormat(fmt));
            xls.SetCellValue(177, 2, new TFormula("=B150"));

            fmt = xls.GetCellVisibleFormatDef(177, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 3, xls.AddFormat(fmt));
            xls.SetCellValue(177, 3, new TFormula("=C150"));

            fmt = xls.GetCellVisibleFormatDef(177, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 4, xls.AddFormat(fmt));
            xls.SetCellValue(177, 4, new TFormula("=D150"));

            fmt = xls.GetCellVisibleFormatDef(177, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 5, xls.AddFormat(fmt));
            xls.SetCellValue(177, 5, new TFormula("=E150"));

            fmt = xls.GetCellVisibleFormatDef(177, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 6, xls.AddFormat(fmt));
            xls.SetCellValue(177, 6, new TFormula("=F150"));

            fmt = xls.GetCellVisibleFormatDef(177, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(177, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(177, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 9, xls.AddFormat(fmt));
            xls.SetCellValue(177, 9, new TFormula("=(B177*$B$18+C177*$B$19+D177*$B$23+E177*$B$28+F177*$B$26)/0.973"));

            fmt = xls.GetCellVisibleFormatDef(177, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 10, xls.AddFormat(fmt));
            xls.SetCellValue(177, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(177, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(177, 11, xls.AddFormat(fmt));
            xls.SetCellValue(177, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(177, 13);
            fmt.Format = "0";
            xls.SetCellFormat(177, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(177, 14);
            fmt.Format = "0";
            xls.SetCellFormat(177, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(178, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(178, 1, xls.AddFormat(fmt));
            xls.SetCellValue(178, 1, 6);

            fmt = xls.GetCellVisibleFormatDef(178, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 2, xls.AddFormat(fmt));
            xls.SetCellValue(178, 2, new TFormula("=B151"));

            fmt = xls.GetCellVisibleFormatDef(178, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 3, xls.AddFormat(fmt));
            xls.SetCellValue(178, 3, new TFormula("=C151"));

            fmt = xls.GetCellVisibleFormatDef(178, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 4, xls.AddFormat(fmt));
            xls.SetCellValue(178, 4, new TFormula("=D151"));

            fmt = xls.GetCellVisibleFormatDef(178, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 5, xls.AddFormat(fmt));
            xls.SetCellValue(178, 5, new TFormula("=E151"));

            fmt = xls.GetCellVisibleFormatDef(178, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 6, xls.AddFormat(fmt));
            xls.SetCellValue(178, 6, new TFormula("=F151"));

            fmt = xls.GetCellVisibleFormatDef(178, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(178, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(178, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 9, xls.AddFormat(fmt));
            xls.SetCellValue(178, 9, new TFormula("=(B178*$B$18+C178*$B$19+D178*$B$23+E178*$B$28+F178*$B$26)/0.973"));

            fmt = xls.GetCellVisibleFormatDef(178, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 10, xls.AddFormat(fmt));
            xls.SetCellValue(178, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(178, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(178, 11, xls.AddFormat(fmt));
            xls.SetCellValue(178, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(178, 13);
            fmt.Format = "0";
            xls.SetCellFormat(178, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(178, 14);
            fmt.Format = "0";
            xls.SetCellFormat(178, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(179, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(179, 1, xls.AddFormat(fmt));
            xls.SetCellValue(179, 1, 7);

            fmt = xls.GetCellVisibleFormatDef(179, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 2, xls.AddFormat(fmt));
            xls.SetCellValue(179, 2, new TFormula("=B152"));

            fmt = xls.GetCellVisibleFormatDef(179, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 3, xls.AddFormat(fmt));
            xls.SetCellValue(179, 3, new TFormula("=C152"));

            fmt = xls.GetCellVisibleFormatDef(179, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 4, xls.AddFormat(fmt));
            xls.SetCellValue(179, 4, new TFormula("=D152"));

            fmt = xls.GetCellVisibleFormatDef(179, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 5, xls.AddFormat(fmt));
            xls.SetCellValue(179, 5, new TFormula("=E152"));

            fmt = xls.GetCellVisibleFormatDef(179, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 6, xls.AddFormat(fmt));
            xls.SetCellValue(179, 6, new TFormula("=F152"));

            fmt = xls.GetCellVisibleFormatDef(179, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(179, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(179, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 9, xls.AddFormat(fmt));
            xls.SetCellValue(179, 9, new TFormula("=(B179*$B$18+C179*$B$19+D179*$B$23+E179*$B$28+F179*$B$26)/0.973"));

            fmt = xls.GetCellVisibleFormatDef(179, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 10, xls.AddFormat(fmt));
            xls.SetCellValue(179, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(179, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(179, 11, xls.AddFormat(fmt));
            xls.SetCellValue(179, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(179, 13);
            fmt.Format = "0";
            xls.SetCellFormat(179, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(179, 14);
            fmt.Format = "0";
            xls.SetCellFormat(179, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(180, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(180, 1, xls.AddFormat(fmt));
            xls.SetCellValue(180, 1, 8);

            fmt = xls.GetCellVisibleFormatDef(180, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 2, xls.AddFormat(fmt));
            xls.SetCellValue(180, 2, new TFormula("=B153"));

            fmt = xls.GetCellVisibleFormatDef(180, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 3, xls.AddFormat(fmt));
            xls.SetCellValue(180, 3, new TFormula("=C153"));

            fmt = xls.GetCellVisibleFormatDef(180, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 4, xls.AddFormat(fmt));
            xls.SetCellValue(180, 4, new TFormula("=D153"));

            fmt = xls.GetCellVisibleFormatDef(180, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 5, xls.AddFormat(fmt));
            xls.SetCellValue(180, 5, new TFormula("=E153"));

            fmt = xls.GetCellVisibleFormatDef(180, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 6, xls.AddFormat(fmt));
            xls.SetCellValue(180, 6, new TFormula("=F153"));

            fmt = xls.GetCellVisibleFormatDef(180, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(180, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(180, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 9, xls.AddFormat(fmt));
            xls.SetCellValue(180, 9, new TFormula("=(B180*$B$18+C180*$B$19+D180*$B$23+E180*$B$28+F180*$B$26)/0.973"));

            fmt = xls.GetCellVisibleFormatDef(180, 10);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 10, xls.AddFormat(fmt));
            xls.SetCellValue(180, 10, ".");

            fmt = xls.GetCellVisibleFormatDef(180, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(180, 11, xls.AddFormat(fmt));
            xls.SetCellValue(180, 11, ".");

            fmt = xls.GetCellVisibleFormatDef(180, 13);
            fmt.Format = "0";
            xls.SetCellFormat(180, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(180, 14);
            fmt.Format = "0";
            xls.SetCellFormat(180, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(181, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(181, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(181, 6);
            fmt.Format = "0";
            xls.SetCellFormat(181, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(181, 7);
            fmt.Format = "0";
            xls.SetCellFormat(181, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(181, 8);
            fmt.Format = "0";
            xls.SetCellFormat(181, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(182, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(182, 1, xls.AddFormat(fmt));
            xls.SetCellValue(182, 1, "Productividad - Café Pergamino Seco");

            fmt = xls.GetCellVisibleFormatDef(182, 6);
            fmt.Format = "0";
            xls.SetCellFormat(182, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(182, 7);
            fmt.Format = "0";
            xls.SetCellFormat(182, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(182, 8);
            fmt.Format = "0";
            xls.SetCellFormat(182, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(183, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(183, 1, xls.AddFormat(fmt));
            xls.SetCellValue(183, 1, "¿Cuántos KILOS de CEREZA recoge POR HECTÁREA?");

            fmt = xls.GetCellVisibleFormatDef(183, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(183, 4, xls.AddFormat(fmt));
            xls.SetCellValue(183, 4, "Kilos por hectaria por año (conforme a la hectarea de cada productor)");

            fmt = xls.GetCellVisibleFormatDef(183, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(183, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(183, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(183, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(183, 7);
            fmt.Format = "0";
            xls.SetCellFormat(183, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(183, 8);
            fmt.Format = "0";
            xls.SetCellFormat(183, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(184, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(184, 1, xls.AddFormat(fmt));
            xls.SetCellValue(184, 1, "Año");

            fmt = xls.GetCellVisibleFormatDef(184, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(184, 4, xls.AddFormat(fmt));
            xls.SetCellValue(184, 4, "TOTAL");

            fmt = xls.GetCellVisibleFormatDef(184, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(184, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(184, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(184, 6, xls.AddFormat(fmt));
            xls.SetCellValue(184, 6, new TFormula("=G171"));

            fmt = xls.GetCellVisibleFormatDef(184, 7);
            fmt.Format = "0";
            xls.SetCellFormat(184, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(184, 8);
            fmt.Format = "0";
            xls.SetCellFormat(184, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(185, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(185, 1, xls.AddFormat(fmt));
            xls.SetCellValue(185, 1, 0);

            fmt = xls.GetCellVisibleFormatDef(185, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(185, 2, xls.AddFormat(fmt));
            xls.SetCellValue(185, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(185, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(185, 3, xls.AddFormat(fmt));
            xls.SetCellValue(185, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(185, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(185, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(185, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(185, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(185, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(185, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(185, 7);
            fmt.Format = "0";
            xls.SetCellFormat(185, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(185, 8);
            fmt.Format = "0";
            xls.SetCellFormat(185, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(186, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(186, 1, xls.AddFormat(fmt));
            xls.SetCellValue(186, 1, 1);

            fmt = xls.GetCellVisibleFormatDef(186, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(186, 2, xls.AddFormat(fmt));
            xls.SetCellValue(186, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(186, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(186, 3, xls.AddFormat(fmt));
            xls.SetCellValue(186, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(186, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(186, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(186, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(186, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(186, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(186, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(186, 7);
            fmt.Format = "0";
            xls.SetCellFormat(186, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(186, 8);
            fmt.Format = "0";
            xls.SetCellFormat(186, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(187, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(187, 1, xls.AddFormat(fmt));
            xls.SetCellValue(187, 1, 2);

            fmt = xls.GetCellVisibleFormatDef(187, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(187, 2, xls.AddFormat(fmt));
            xls.SetCellValue(187, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(187, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(187, 3, xls.AddFormat(fmt));
            xls.SetCellValue(187, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(187, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(187, 4, xls.AddFormat(fmt));
            xls.SetCellValue(187, 4, new TFormula("='Inputs advanced'!F77"));

            fmt = xls.GetCellVisibleFormatDef(187, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(187, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(187, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(187, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(187, 7);
            fmt.Format = "0";
            xls.SetCellFormat(187, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(187, 8);
            fmt.Format = "0";
            xls.SetCellFormat(187, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(188, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(188, 1, xls.AddFormat(fmt));
            xls.SetCellValue(188, 1, 3);

            fmt = xls.GetCellVisibleFormatDef(188, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(188, 2, xls.AddFormat(fmt));
            xls.SetCellValue(188, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(188, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(188, 3, xls.AddFormat(fmt));
            xls.SetCellValue(188, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(188, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(188, 4, xls.AddFormat(fmt));
            xls.SetCellValue(188, 4, new TFormula("='Inputs advanced'!F77"));

            fmt = xls.GetCellVisibleFormatDef(188, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(188, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(188, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(188, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(188, 7);
            fmt.Format = "0";
            xls.SetCellFormat(188, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(188, 8);
            fmt.Format = "0";
            xls.SetCellFormat(188, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(189, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(189, 1, xls.AddFormat(fmt));
            xls.SetCellValue(189, 1, 4);

            fmt = xls.GetCellVisibleFormatDef(189, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(189, 2, xls.AddFormat(fmt));
            xls.SetCellValue(189, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(189, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(189, 3, xls.AddFormat(fmt));
            xls.SetCellValue(189, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(189, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(189, 4, xls.AddFormat(fmt));
            xls.SetCellValue(189, 4, new TFormula("='Inputs advanced'!F104"));

            fmt = xls.GetCellVisibleFormatDef(189, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(189, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(189, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(189, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(189, 7);
            fmt.Format = "0";
            xls.SetCellFormat(189, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(189, 8);
            fmt.Format = "0";
            xls.SetCellFormat(189, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(190, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(190, 1, xls.AddFormat(fmt));
            xls.SetCellValue(190, 1, 5);

            fmt = xls.GetCellVisibleFormatDef(190, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(190, 2, xls.AddFormat(fmt));
            xls.SetCellValue(190, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(190, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(190, 3, xls.AddFormat(fmt));
            xls.SetCellValue(190, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(190, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(190, 4, xls.AddFormat(fmt));
            xls.SetCellValue(190, 4, new TFormula("='Inputs advanced'!F104"));

            fmt = xls.GetCellVisibleFormatDef(190, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(190, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(190, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(190, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(190, 7);
            fmt.Format = "0";
            xls.SetCellFormat(190, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(190, 8);
            fmt.Format = "0";
            xls.SetCellFormat(190, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(191, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(191, 1, xls.AddFormat(fmt));
            xls.SetCellValue(191, 1, 6);

            fmt = xls.GetCellVisibleFormatDef(191, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(191, 2, xls.AddFormat(fmt));
            xls.SetCellValue(191, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(191, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(191, 3, xls.AddFormat(fmt));
            xls.SetCellValue(191, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(191, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(191, 4, xls.AddFormat(fmt));
            xls.SetCellValue(191, 4, new TFormula("='Inputs advanced'!F131"));

            fmt = xls.GetCellVisibleFormatDef(191, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(191, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(191, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(191, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(191, 7);
            fmt.Format = "0";
            xls.SetCellFormat(191, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(191, 8);
            fmt.Format = "0";
            xls.SetCellFormat(191, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(192, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(192, 1, xls.AddFormat(fmt));
            xls.SetCellValue(192, 1, 7);

            fmt = xls.GetCellVisibleFormatDef(192, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(192, 2, xls.AddFormat(fmt));
            xls.SetCellValue(192, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(192, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(192, 3, xls.AddFormat(fmt));
            xls.SetCellValue(192, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(192, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(192, 4, xls.AddFormat(fmt));
            xls.SetCellValue(192, 4, new TFormula("='Inputs advanced'!F131"));

            fmt = xls.GetCellVisibleFormatDef(192, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(192, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(192, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(192, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(192, 7);
            fmt.Format = "0";
            xls.SetCellFormat(192, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(192, 8);
            fmt.Format = "0";
            xls.SetCellFormat(192, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(193, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(193, 1, xls.AddFormat(fmt));
            xls.SetCellValue(193, 1, 8);

            fmt = xls.GetCellVisibleFormatDef(193, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(193, 2, xls.AddFormat(fmt));
            xls.SetCellValue(193, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(193, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(193, 3, xls.AddFormat(fmt));
            xls.SetCellValue(193, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(193, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(193, 4, xls.AddFormat(fmt));
            xls.SetCellValue(193, 4, new TFormula("='Inputs advanced'!F131"));

            fmt = xls.GetCellVisibleFormatDef(193, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(193, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(193, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(193, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(193, 7);
            fmt.Format = "0";
            xls.SetCellFormat(193, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(193, 8);
            fmt.Format = "0";
            xls.SetCellFormat(193, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(194, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(194, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 4);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(194, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 5);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(194, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 6);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(194, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 7);
            fmt.Format = "0";
            xls.SetCellFormat(194, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 8);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(194, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(194, 10);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(194, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(195, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(195, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(195, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 5);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 6);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 7);
            fmt.Format = "0";
            xls.SetCellFormat(195, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 8);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 9);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 10);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 11);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(195, 12);
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            xls.SetCellFormat(195, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(196, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(196, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(196, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(196, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(196, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(196, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 7);
            fmt.Format = "0";
            xls.SetCellFormat(196, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 8);
            fmt.Format = "0";
            xls.SetCellFormat(196, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 9);
            fmt.Format = "0";
            xls.SetCellFormat(196, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 10);
            fmt.Format = "0";
            xls.SetCellFormat(196, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 11);
            fmt.Format = "0";
            xls.SetCellFormat(196, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(196, 12);
            fmt.Format = "0";
            xls.SetCellFormat(196, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(197, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(197, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(197, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(197, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(197, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(197, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 7);
            fmt.Format = "0";
            xls.SetCellFormat(197, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 8);
            fmt.Format = "0";
            xls.SetCellFormat(197, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 9);
            fmt.Format = "0";
            xls.SetCellFormat(197, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 10);
            fmt.Format = "0";
            xls.SetCellFormat(197, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 11);
            fmt.Format = "0";
            xls.SetCellFormat(197, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(197, 12);
            fmt.Format = "0";
            xls.SetCellFormat(197, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(198, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(198, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(198, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(198, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(198, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(198, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 7);
            fmt.Format = "0";
            xls.SetCellFormat(198, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 8);
            fmt.Format = "0";
            xls.SetCellFormat(198, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 9);
            fmt.Format = "0";
            xls.SetCellFormat(198, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 10);
            fmt.Format = "0";
            xls.SetCellFormat(198, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 11);
            fmt.Format = "0";
            xls.SetCellFormat(198, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(198, 12);
            fmt.Format = "0";
            xls.SetCellFormat(198, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(199, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(199, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(199, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(199, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(199, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(199, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 7);
            fmt.Format = "0";
            xls.SetCellFormat(199, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 8);
            fmt.Format = "0";
            xls.SetCellFormat(199, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 9);
            fmt.Format = "0";
            xls.SetCellFormat(199, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 10);
            fmt.Format = "0";
            xls.SetCellFormat(199, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 11);
            fmt.Format = "0";
            xls.SetCellFormat(199, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(199, 12);
            fmt.Format = "0";
            xls.SetCellFormat(199, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(200, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(200, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(200, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(200, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(200, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(200, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 7);
            fmt.Format = "0";
            xls.SetCellFormat(200, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 8);
            fmt.Format = "0";
            xls.SetCellFormat(200, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 9);
            fmt.Format = "0";
            xls.SetCellFormat(200, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 10);
            fmt.Format = "0";
            xls.SetCellFormat(200, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 11);
            fmt.Format = "0";
            xls.SetCellFormat(200, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(200, 12);
            fmt.Format = "0";
            xls.SetCellFormat(200, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(201, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(201, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(201, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(201, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(201, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 7);
            fmt.Format = "0";
            xls.SetCellFormat(201, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 8);
            fmt.Format = "0";
            xls.SetCellFormat(201, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 9);
            fmt.Format = "0";
            xls.SetCellFormat(201, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 10);
            fmt.Format = "0";
            xls.SetCellFormat(201, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 11);
            fmt.Format = "0";
            xls.SetCellFormat(201, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 12);
            fmt.Format = "0";
            xls.SetCellFormat(201, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(202, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(202, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(202, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(202, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(202, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(202, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 7);
            fmt.Format = "0";
            xls.SetCellFormat(202, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 8);
            fmt.Format = "0";
            xls.SetCellFormat(202, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 9);
            fmt.Format = "0";
            xls.SetCellFormat(202, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 10);
            fmt.Format = "0";
            xls.SetCellFormat(202, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 11);
            fmt.Format = "0";
            xls.SetCellFormat(202, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 12);
            fmt.Format = "0";
            xls.SetCellFormat(202, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(203, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(203, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(203, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(203, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(203, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(203, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 7);
            fmt.Format = "0";
            xls.SetCellFormat(203, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 8);
            fmt.Format = "0";
            xls.SetCellFormat(203, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 9);
            fmt.Format = "0";
            xls.SetCellFormat(203, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 10);
            fmt.Format = "0";
            xls.SetCellFormat(203, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 11);
            fmt.Format = "0";
            xls.SetCellFormat(203, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(203, 12);
            fmt.Format = "0";
            xls.SetCellFormat(203, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(204, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(204, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(204, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(204, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(204, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(204, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 7);
            fmt.Format = "0";
            xls.SetCellFormat(204, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 8);
            fmt.Format = "0";
            xls.SetCellFormat(204, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 9);
            fmt.Format = "0";
            xls.SetCellFormat(204, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 10);
            fmt.Format = "0";
            xls.SetCellFormat(204, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 11);
            fmt.Format = "0";
            xls.SetCellFormat(204, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(204, 12);
            fmt.Format = "0";
            xls.SetCellFormat(204, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(205, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(205, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(205, 6);
            fmt.Format = "0";
            xls.SetCellFormat(205, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(205, 7);
            fmt.Format = "0";
            xls.SetCellFormat(205, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(205, 8);
            fmt.Format = "0";
            xls.SetCellFormat(205, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(206, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(206, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(206, 6);
            fmt.Format = "0";
            xls.SetCellFormat(206, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(206, 7);
            fmt.Format = "0";
            xls.SetCellFormat(206, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(206, 8);
            fmt.Format = "0";
            xls.SetCellFormat(206, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(207, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.0499893185216834);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(207, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(207, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(207, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(207, 6);
            fmt.Format = "0";
            xls.SetCellFormat(207, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(207, 7);
            fmt.Format = "0";
            xls.SetCellFormat(207, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(207, 8);
            fmt.Format = "0";
            xls.SetCellFormat(207, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.0499893185216834);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(208, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(208, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 6);
            fmt.Format = "0";
            xls.SetCellFormat(208, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 7);
            fmt.Format = "0";
            xls.SetCellFormat(208, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(208, 8);
            fmt.Format = "0";
            xls.SetCellFormat(208, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(209, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(209, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(209, 6);
            fmt.Format = "0";
            xls.SetCellFormat(209, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(209, 7);
            fmt.Format = "0";
            xls.SetCellFormat(209, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(209, 8);
            fmt.Format = "0";
            xls.SetCellFormat(209, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(210, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(210, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(210, 6);
            fmt.Format = "0";
            xls.SetCellFormat(210, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(210, 7);
            fmt.Format = "0";
            xls.SetCellFormat(210, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(210, 8);
            fmt.Format = "0";
            xls.SetCellFormat(210, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(211, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(211, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(211, 6);
            fmt.Format = "0";
            xls.SetCellFormat(211, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(211, 7);
            fmt.Format = "0";
            xls.SetCellFormat(211, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(211, 8);
            fmt.Format = "0";
            xls.SetCellFormat(211, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(212, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(212, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(212, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 6);
            fmt.Format = "0";
            xls.SetCellFormat(212, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 7);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(212, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(212, 8);
            fmt.Format = "0";
            xls.SetCellFormat(212, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(213, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(213, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(213, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(213, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(213, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(213, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(213, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(213, 8);
            fmt.Format = "0";
            xls.SetCellFormat(213, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(214, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(214, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(214, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(214, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(214, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(214, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(214, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(214, 8);
            fmt.Format = "0";
            xls.SetCellFormat(214, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(215, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(215, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(215, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(215, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(215, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(215, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(215, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(215, 8);
            fmt.Format = "0";
            xls.SetCellFormat(215, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(216, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(216, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(216, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(216, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(216, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(216, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(216, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(216, 8);
            fmt.Format = "0";
            xls.SetCellFormat(216, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(217, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(217, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(217, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(217, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(217, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(217, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(217, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(217, 8);
            fmt.Format = "0";
            xls.SetCellFormat(217, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(218, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(218, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(218, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(218, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(218, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(218, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(218, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(218, 8);
            fmt.Format = "0";
            xls.SetCellFormat(218, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(219, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(219, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(219, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(219, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(219, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(219, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(219, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(219, 8);
            fmt.Format = "0";
            xls.SetCellFormat(219, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(220, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(220, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(220, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(220, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(220, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(220, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(220, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(220, 8);
            fmt.Format = "0";
            xls.SetCellFormat(220, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(221, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(221, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(221, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(221, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(221, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(221, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(221, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(221, 8);
            fmt.Format = "0";
            xls.SetCellFormat(221, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(226, 1);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(226, 1, xls.AddFormat(fmt));
            xls.SetCellValue(226, 1, "Costos");

            fmt = xls.GetCellVisibleFormatDef(227, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(227, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(228, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(228, 1, xls.AddFormat(fmt));
            xls.SetCellValue(228, 1, "Materiales");

            fmt = xls.GetCellVisibleFormatDef(228, 2);
            fmt.WrapText = true;
            xls.SetCellFormat(228, 2, xls.AddFormat(fmt));
            xls.SetCellValue(228, 2, "Cantidad confirmada con el productor necesarias para una manzana");

            fmt = xls.GetCellVisibleFormatDef(228, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(228, 3, xls.AddFormat(fmt));
            xls.SetCellValue(228, 3, "Unidad de medida de venta");

            fmt = xls.GetCellVisibleFormatDef(228, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(228, 4, xls.AddFormat(fmt));
            xls.SetCellValue(228, 4, "Precio POR UNIDAD  en Moneda local ");

            fmt = xls.GetCellVisibleFormatDef(228, 5);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(228, 5, xls.AddFormat(fmt));
            xls.SetCellValue(228, 5, "Costo Total en Moneda local ");

            fmt = xls.GetCellVisibleFormatDef(229, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(229, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(229, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(229, 2, xls.AddFormat(fmt));
            xls.SetCellValue(229, 2, "Nota 1: Si no usa 0");

            fmt = xls.GetCellVisibleFormatDef(229, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(229, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(229, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(229, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(229, 5);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(229, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(229, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(229, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(229, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(229, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(230, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(230, 2, xls.AddFormat(fmt));
            xls.SetCellValue(230, 2, "Nota 2: Algunas cosas se necesitan independiente si es una hectarea o 3. Registrarlas");

            fmt = xls.GetCellVisibleFormatDef(230, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(230, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(230, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 5);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(230, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(230, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(230, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(230, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(231, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(231, 1, xls.AddFormat(fmt));
            xls.SetCellValue(231, 1, "Materiales para el Germinador ");

            fmt = xls.GetCellVisibleFormatDef(231, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(231, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(231, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(231, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(231, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(231, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(231, 5);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(231, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(231, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(231, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(231, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(231, 7, xls.AddFormat(fmt));
            xls.SetCellValue(232, 1, "Semilla");

            fmt = xls.GetCellVisibleFormatDef(232, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(232, 2, xls.AddFormat(fmt));
            xls.SetCellValue(232, 2, new TFormula("='Inputs advanced'!F252"));

            fmt = xls.GetCellVisibleFormatDef(232, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(232, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(232, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(232, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(232, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(232, 7, xls.AddFormat(fmt));
            xls.SetCellValue(233, 1, "Germinador/Marco semillero");

            fmt = xls.GetCellVisibleFormatDef(233, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(233, 2, xls.AddFormat(fmt));
            xls.SetCellValue(233, 2, new TFormula("='Inputs advanced'!F253"));

            fmt = xls.GetCellVisibleFormatDef(233, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(233, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(233, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(233, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(233, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(233, 7, xls.AddFormat(fmt));
            xls.SetCellValue(234, 1, "Sustrato de arena");

            fmt = xls.GetCellVisibleFormatDef(234, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(234, 2, xls.AddFormat(fmt));
            xls.SetCellValue(234, 2, new TFormula("='Inputs advanced'!F254"));

            fmt = xls.GetCellVisibleFormatDef(234, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(234, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(234, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(234, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(234, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(234, 7, xls.AddFormat(fmt));
            xls.SetCellValue(235, 1, "Sulfocalcio");

            fmt = xls.GetCellVisibleFormatDef(235, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(235, 2, xls.AddFormat(fmt));
            xls.SetCellValue(235, 2, new TFormula("='Inputs advanced'!F255"));

            fmt = xls.GetCellVisibleFormatDef(235, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(235, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(235, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(235, 4, xls.AddFormat(fmt));
            xls.SetCellValue(236, 1, "Cal");

            fmt = xls.GetCellVisibleFormatDef(236, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(236, 2, xls.AddFormat(fmt));
            xls.SetCellValue(236, 2, new TFormula("='Inputs advanced'!F256"));

            fmt = xls.GetCellVisibleFormatDef(236, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(236, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(236, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(236, 4, xls.AddFormat(fmt));
            xls.SetCellValue(237, 1, "Plastico");

            fmt = xls.GetCellVisibleFormatDef(237, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(237, 2, xls.AddFormat(fmt));
            xls.SetCellValue(237, 2, new TFormula("='Inputs advanced'!F257"));

            fmt = xls.GetCellVisibleFormatDef(237, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(237, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(237, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(237, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(238, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(238, 1, xls.AddFormat(fmt));
            xls.SetCellValue(238, 1, "Otros");

            fmt = xls.GetCellVisibleFormatDef(238, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(238, 2, xls.AddFormat(fmt));
            xls.SetCellValue(238, 2, new TFormula("='Inputs advanced'!F258"));

            fmt = xls.GetCellVisibleFormatDef(238, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(238, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(238, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(238, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(239, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(239, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(240, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(240, 1, xls.AddFormat(fmt));
            xls.SetCellValue(240, 1, "Materiales para vivero o ramada incluyendo almacigos");

            fmt = xls.GetCellVisibleFormatDef(240, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(240, 3, xls.AddFormat(fmt));
            xls.SetCellValue(241, 1, "Abono orgánico (Ej: Bocachi, otros)");

            fmt = xls.GetCellVisibleFormatDef(241, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(241, 2, xls.AddFormat(fmt));
            xls.SetCellValue(241, 2, new TFormula("='Inputs advanced'!F260"));

            fmt = xls.GetCellVisibleFormatDef(241, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(241, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(241, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(241, 4, xls.AddFormat(fmt));
            xls.SetCellValue(242, 1, "Bolsitas de plastico");

            fmt = xls.GetCellVisibleFormatDef(242, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(242, 2, xls.AddFormat(fmt));
            xls.SetCellValue(242, 2, new TFormula("='Inputs advanced'!F261"));

            fmt = xls.GetCellVisibleFormatDef(242, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(242, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(242, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            xls.SetCellFormat(242, 4, xls.AddFormat(fmt));
            xls.SetCellValue(243, 1, "Saran - Polisombra - Malla rache");

            fmt = xls.GetCellVisibleFormatDef(243, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(243, 2, xls.AddFormat(fmt));
            xls.SetCellValue(243, 2, new TFormula("='Inputs advanced'!F262"));

            fmt = xls.GetCellVisibleFormatDef(243, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(243, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(243, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(243, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(243, 7);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(243, 7, xls.AddFormat(fmt));
            xls.SetCellValue(244, 1, "Postes de madera");

            fmt = xls.GetCellVisibleFormatDef(244, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(244, 2, xls.AddFormat(fmt));
            xls.SetCellValue(244, 2, new TFormula("='Inputs advanced'!F263"));

            fmt = xls.GetCellVisibleFormatDef(244, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(244, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(244, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(244, 4, xls.AddFormat(fmt));
            xls.SetCellValue(245, 1, "Alambre de amarre");

            fmt = xls.GetCellVisibleFormatDef(245, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(245, 2, xls.AddFormat(fmt));
            xls.SetCellValue(245, 2, new TFormula("='Inputs advanced'!F264"));

            fmt = xls.GetCellVisibleFormatDef(245, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(245, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(245, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(245, 4, xls.AddFormat(fmt));
            xls.SetCellValue(246, 1, "Malla Ciclonica");

            fmt = xls.GetCellVisibleFormatDef(246, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(246, 2, xls.AddFormat(fmt));
            xls.SetCellValue(246, 2, new TFormula("='Inputs advanced'!F265"));

            fmt = xls.GetCellVisibleFormatDef(246, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(246, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(246, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(246, 4, xls.AddFormat(fmt));
            xls.SetCellValue(247, 1, "Grapas");

            fmt = xls.GetCellVisibleFormatDef(247, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(247, 2, xls.AddFormat(fmt));
            xls.SetCellValue(247, 2, new TFormula("='Inputs advanced'!F266"));

            fmt = xls.GetCellVisibleFormatDef(247, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(247, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(247, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(247, 4, xls.AddFormat(fmt));
            xls.SetCellValue(248, 1, "Tierra para almacigos");

            fmt = xls.GetCellVisibleFormatDef(248, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(248, 2, xls.AddFormat(fmt));
            xls.SetCellValue(248, 2, new TFormula("='Inputs advanced'!F267"));

            fmt = xls.GetCellVisibleFormatDef(248, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(248, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(248, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(248, 4, xls.AddFormat(fmt));
            xls.SetCellValue(249, 1, "Biofertilizantes líquidos (para foliar en el vivero)");

            fmt = xls.GetCellVisibleFormatDef(249, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(249, 2, xls.AddFormat(fmt));
            xls.SetCellValue(249, 2, new TFormula("='Inputs advanced'!F268"));

            fmt = xls.GetCellVisibleFormatDef(249, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(249, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(249, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(249, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(249, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(249, 6, xls.AddFormat(fmt));
            xls.SetCellValue(250, 1, "Agroquímicos (en el vivero)");

            fmt = xls.GetCellVisibleFormatDef(250, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(250, 2, xls.AddFormat(fmt));
            xls.SetCellValue(250, 2, new TFormula("='Inputs advanced'!F269"));

            fmt = xls.GetCellVisibleFormatDef(250, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(250, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(250, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(250, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(250, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(250, 6, xls.AddFormat(fmt));
            xls.SetCellValue(251, 1, "Fungicida");

            fmt = xls.GetCellVisibleFormatDef(251, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(251, 2, xls.AddFormat(fmt));
            xls.SetCellValue(251, 2, new TFormula("='Inputs advanced'!F270"));

            fmt = xls.GetCellVisibleFormatDef(251, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(251, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(251, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(251, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(251, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(251, 6, xls.AddFormat(fmt));
            xls.SetCellValue(252, 1, "Roca fosfórica");

            fmt = xls.GetCellVisibleFormatDef(252, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(252, 2, xls.AddFormat(fmt));
            xls.SetCellValue(252, 2, new TFormula("='Inputs advanced'!F271"));

            fmt = xls.GetCellVisibleFormatDef(252, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(252, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(252, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(252, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(252, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(252, 6, xls.AddFormat(fmt));
            xls.SetCellValue(253, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(253, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(253, 2, xls.AddFormat(fmt));
            xls.SetCellValue(253, 2, new TFormula("='Inputs advanced'!F272"));

            fmt = xls.GetCellVisibleFormatDef(253, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(253, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(253, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(253, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(253, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(253, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(254, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(254, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(255, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(255, 1, xls.AddFormat(fmt));
            xls.SetCellValue(255, 1, "Materiales para Preparación del terreno y siembra");

            fmt = xls.GetCellVisibleFormatDef(255, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(255, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(255, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(255, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(256, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(256, 1, xls.AddFormat(fmt));
            xls.SetCellValue(256, 1, "Abono organico para los Hoyos");

            fmt = xls.GetCellVisibleFormatDef(256, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(256, 2, xls.AddFormat(fmt));
            xls.SetCellValue(256, 2, new TFormula("='Inputs advanced'!F274"));

            fmt = xls.GetCellVisibleFormatDef(256, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(256, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(256, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(256, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(257, 1, xls.AddFormat(fmt));
            xls.SetCellValue(257, 1, "Especificos:");

            fmt = xls.GetCellVisibleFormatDef(257, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(257, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(258, 1, xls.AddFormat(fmt));
            xls.SetCellValue(258, 1, "Harina de Roca");

            fmt = xls.GetCellVisibleFormatDef(258, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(258, 2, xls.AddFormat(fmt));
            xls.SetCellValue(258, 2, new TFormula("='Inputs advanced'!F275"));

            fmt = xls.GetCellVisibleFormatDef(258, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(258, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(258, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(259, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(259, 1, xls.AddFormat(fmt));
            xls.SetCellValue(259, 1, "Cascarilla de Café");

            fmt = xls.GetCellVisibleFormatDef(259, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(259, 2, xls.AddFormat(fmt));
            xls.SetCellValue(259, 2, new TFormula("='Inputs advanced'!F276"));

            fmt = xls.GetCellVisibleFormatDef(259, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(259, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(259, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(259, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(260, 1, xls.AddFormat(fmt));
            xls.SetCellValue(260, 1, "Gallinaza");

            fmt = xls.GetCellVisibleFormatDef(260, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(260, 2, xls.AddFormat(fmt));
            xls.SetCellValue(260, 2, new TFormula("='Inputs advanced'!F277"));

            fmt = xls.GetCellVisibleFormatDef(260, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(260, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(260, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(261, 1, xls.AddFormat(fmt));
            xls.SetCellValue(261, 1, "Abono químico para los hoyos");

            fmt = xls.GetCellVisibleFormatDef(261, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(261, 2, xls.AddFormat(fmt));
            xls.SetCellValue(261, 2, new TFormula("=IF('Inputs advanced'!F179=1,0,'Inputs advanced'!F278)"));

            fmt = xls.GetCellVisibleFormatDef(261, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(261, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(261, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(262, 1, xls.AddFormat(fmt));
            xls.SetCellValue(262, 1, "Cal");

            fmt = xls.GetCellVisibleFormatDef(262, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(262, 2, xls.AddFormat(fmt));
            xls.SetCellValue(262, 2, new TFormula("='Inputs advanced'!F279"));

            fmt = xls.GetCellVisibleFormatDef(262, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(262, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(262, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(263, 1, xls.AddFormat(fmt));
            xls.SetCellValue(263, 1, "Otros elementos para los hoyos: ");

            fmt = xls.GetCellVisibleFormatDef(263, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(263, 2, xls.AddFormat(fmt));
            xls.SetCellValue(263, 2, new TFormula("='Inputs advanced'!F280"));

            fmt = xls.GetCellVisibleFormatDef(263, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(263, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(263, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(264, 1, xls.AddFormat(fmt));
            xls.SetCellValue(264, 1, "Total Abonos");

            fmt = xls.GetCellVisibleFormatDef(264, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(264, 2, xls.AddFormat(fmt));
            xls.SetCellValue(264, 2, new TFormula("=SUM(B256:B263)"));

            fmt = xls.GetCellVisibleFormatDef(264, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(264, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(264, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(265, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(265, 1, xls.AddFormat(fmt));
            xls.SetCellValue(265, 1, "Materiales para Levante");

            fmt = xls.GetCellVisibleFormatDef(265, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(265, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(265, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(265, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(265, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(265, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(265, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(265, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(266, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(266, 1, xls.AddFormat(fmt));
            xls.SetCellValue(266, 1, "Abono organico para levante (alrededor de la planta)");

            fmt = xls.GetCellVisibleFormatDef(266, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(266, 2, xls.AddFormat(fmt));
            xls.SetCellValue(266, 2, new TFormula("='Inputs advanced'!F282"));

            fmt = xls.GetCellVisibleFormatDef(266, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(266, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(266, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(266, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(266, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(266, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(267, 1, xls.AddFormat(fmt));
            xls.SetCellValue(267, 1, "Especificos:");

            fmt = xls.GetCellVisibleFormatDef(267, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(267, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(268, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(268, 1, xls.AddFormat(fmt));
            xls.SetCellValue(268, 1, "Harina de Roca");

            fmt = xls.GetCellVisibleFormatDef(268, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(268, 2, xls.AddFormat(fmt));
            xls.SetCellValue(268, 2, new TFormula("='Inputs advanced'!F283"));

            fmt = xls.GetCellVisibleFormatDef(268, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(268, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(268, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(268, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(269, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(269, 1, xls.AddFormat(fmt));
            xls.SetCellValue(269, 1, "Cascarilla de Café");

            fmt = xls.GetCellVisibleFormatDef(269, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(269, 2, xls.AddFormat(fmt));
            xls.SetCellValue(269, 2, new TFormula("='Inputs advanced'!F284"));

            fmt = xls.GetCellVisibleFormatDef(269, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(269, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(269, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(269, 4, xls.AddFormat(fmt));
            xls.SetCellValue(269, 4, "pasar estos insumos uno por uno");

            fmt = xls.GetCellVisibleFormatDef(270, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(270, 1, xls.AddFormat(fmt));
            xls.SetCellValue(270, 1, "Gallinaza");

            fmt = xls.GetCellVisibleFormatDef(270, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(270, 2, xls.AddFormat(fmt));
            xls.SetCellValue(270, 2, new TFormula("='Inputs advanced'!F285"));

            fmt = xls.GetCellVisibleFormatDef(270, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(270, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(270, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(270, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(270, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(270, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(271, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(271, 1, xls.AddFormat(fmt));
            xls.SetCellValue(271, 1, "Abono químico para levante (alrededor de la planta)");

            fmt = xls.GetCellVisibleFormatDef(271, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(271, 2, xls.AddFormat(fmt));
            xls.SetCellValue(271, 2, new TFormula("=IF('Inputs advanced'!F179=1,0,'Inputs advanced'!F286)"));

            fmt = xls.GetCellVisibleFormatDef(271, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(271, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(271, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(271, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(271, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(271, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(272, 1, xls.AddFormat(fmt));
            xls.SetCellValue(272, 1, "Insumos para la foliación en la plantilla");

            fmt = xls.GetCellVisibleFormatDef(272, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(272, 2, xls.AddFormat(fmt));
            xls.SetCellValue(272, 2, new TFormula("='Inputs advanced'!F287"));

            fmt = xls.GetCellVisibleFormatDef(272, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(272, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(272, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(272, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(273, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(273, 1, xls.AddFormat(fmt));
            xls.SetCellValue(273, 1, "Otros elementos para siembra y levante:");

            fmt = xls.GetCellVisibleFormatDef(273, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(273, 2, xls.AddFormat(fmt));
            xls.SetCellValue(273, 2, new TFormula("='Inputs advanced'!F288"));

            fmt = xls.GetCellVisibleFormatDef(273, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(273, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(273, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(273, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(273, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(273, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(274, 1, xls.AddFormat(fmt));
            xls.SetCellValue(274, 1, "Total Abonos");

            fmt = xls.GetCellVisibleFormatDef(274, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(274, 2, xls.AddFormat(fmt));
            xls.SetCellValue(274, 2, new TFormula("=SUM(B265:B273)"));

            fmt = xls.GetCellVisibleFormatDef(274, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(274, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 4);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(274, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(275, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(275, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(275, 4);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(275, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(275, 5);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(275, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(276, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(276, 1, xls.AddFormat(fmt));
            xls.SetCellValue(276, 1, "Materiales para Sostenimiento o Mantenimiento de la Finca");

            fmt = xls.GetCellVisibleFormatDef(276, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.WrapText = true;
            xls.SetCellFormat(276, 2, xls.AddFormat(fmt));
            xls.SetCellValue(276, 2, "Costo Total en Moneda local ");

            fmt = xls.GetCellVisibleFormatDef(276, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(276, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(276, 4);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(276, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(276, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(276, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(276, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(276, 7, xls.AddFormat(fmt));
            xls.SetCellValue(276, 7, "Notes");

            fmt = xls.GetCellVisibleFormatDef(277, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(277, 1, xls.AddFormat(fmt));
            xls.SetCellValue(277, 1, "Fertilización y control de plagas");

            fmt = xls.GetCellVisibleFormatDef(277, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(277, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(278, 1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(278, 1, xls.AddFormat(fmt));
            xls.SetCellValue(278, 1, "Abonos");

            fmt = xls.GetCellVisibleFormatDef(278, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(278, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 1);
            fmt.WrapText = true;
            xls.SetCellFormat(279, 1, xls.AddFormat(fmt));
            xls.SetCellValue(279, 1, ".");

            fmt = xls.GetCellVisibleFormatDef(279, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(279, 2, xls.AddFormat(fmt));
            xls.SetCellValue(279, 2, new TFormula("='Inputs advanced'!F290"));

            fmt = xls.GetCellVisibleFormatDef(279, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(279, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(279, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(280, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(280, 1, xls.AddFormat(fmt));
            xls.SetCellValue(280, 1, "Especificos:");

            fmt = xls.GetCellVisibleFormatDef(280, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(280, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(281, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(281, 1, xls.AddFormat(fmt));
            xls.SetCellValue(281, 1, "Harina de Roca");

            fmt = xls.GetCellVisibleFormatDef(281, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(281, 2, xls.AddFormat(fmt));
            xls.SetCellValue(281, 2, new TFormula("='Inputs advanced'!F291"));

            fmt = xls.GetCellVisibleFormatDef(281, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(281, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(281, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(281, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(281, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(281, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(282, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(282, 1, xls.AddFormat(fmt));
            xls.SetCellValue(282, 1, "Cascarilla de Café");

            fmt = xls.GetCellVisibleFormatDef(282, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(282, 2, xls.AddFormat(fmt));
            xls.SetCellValue(282, 2, new TFormula("='Inputs advanced'!F292"));

            fmt = xls.GetCellVisibleFormatDef(282, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(282, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(282, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(282, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(282, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(282, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(283, 1, xls.AddFormat(fmt));
            xls.SetCellValue(283, 1, "Gallinaza");

            fmt = xls.GetCellVisibleFormatDef(283, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(283, 2, xls.AddFormat(fmt));
            xls.SetCellValue(283, 2, new TFormula("='Inputs advanced'!F293"));

            fmt = xls.GetCellVisibleFormatDef(283, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(283, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(283, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(283, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(284, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(284, 1, xls.AddFormat(fmt));
            xls.SetCellValue(284, 1, "Roca fosfórica");

            fmt = xls.GetCellVisibleFormatDef(284, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(284, 2, xls.AddFormat(fmt));
            xls.SetCellValue(284, 2, new TFormula("='Inputs advanced'!F294"));

            fmt = xls.GetCellVisibleFormatDef(284, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(284, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(284, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(284, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(284, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(284, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(285, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(285, 1, xls.AddFormat(fmt));
            xls.SetCellValue(285, 1, "Abono químico para mantenimiento del cultivo ");

            fmt = xls.GetCellVisibleFormatDef(285, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(285, 2, xls.AddFormat(fmt));
            xls.SetCellValue(285, 2, new TFormula("=IF('Inputs advanced'!F179=1,0,'Inputs advanced'!F295)"));

            fmt = xls.GetCellVisibleFormatDef(285, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(285, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(285, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(285, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(285, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(285, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(286, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(286, 1, xls.AddFormat(fmt));
            xls.SetCellValue(286, 1, "Otro(s) abono (s):");

            fmt = xls.GetCellVisibleFormatDef(286, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(286, 2, xls.AddFormat(fmt));
            xls.SetCellValue(286, 2, new TFormula("='Inputs advanced'!F296"));

            fmt = xls.GetCellVisibleFormatDef(286, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(286, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(286, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(286, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(286, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(286, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(287, 1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(287, 1, xls.AddFormat(fmt));
            xls.SetCellValue(287, 1, "Fertilizantes");

            fmt = xls.GetCellVisibleFormatDef(287, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(287, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(288, 1, xls.AddFormat(fmt));
            xls.SetCellValue(288, 1, "Fertilizante organico para foliación:");

            fmt = xls.GetCellVisibleFormatDef(288, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(288, 2, xls.AddFormat(fmt));
            xls.SetCellValue(288, 2, new TFormula("='Inputs advanced'!F297"));

            fmt = xls.GetCellVisibleFormatDef(288, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(288, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(288, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(288, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(289, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(289, 1, xls.AddFormat(fmt));
            xls.SetCellValue(289, 1, "Especificos:");

            fmt = xls.GetCellVisibleFormatDef(289, 2);
            fmt.Format = "#,##0";
            xls.SetCellFormat(289, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(290, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(290, 1, xls.AddFormat(fmt));
            xls.SetCellValue(290, 1, "Caldos bordeles");

            fmt = xls.GetCellVisibleFormatDef(290, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(290, 2, xls.AddFormat(fmt));
            xls.SetCellValue(290, 2, new TFormula("='Inputs advanced'!F298"));

            fmt = xls.GetCellVisibleFormatDef(290, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(290, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(290, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(290, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(290, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(290, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(291, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(291, 1, xls.AddFormat(fmt));
            xls.SetCellValue(291, 1, "Sulfocalcio");

            fmt = xls.GetCellVisibleFormatDef(291, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(291, 2, xls.AddFormat(fmt));
            xls.SetCellValue(291, 2, new TFormula("='Inputs advanced'!F299"));

            fmt = xls.GetCellVisibleFormatDef(291, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(291, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(291, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(291, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(291, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(291, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(292, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(292, 1, xls.AddFormat(fmt));
            xls.SetCellValue(292, 1, "Biofertilizante - multiminerales");

            fmt = xls.GetCellVisibleFormatDef(292, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(292, 2, xls.AddFormat(fmt));
            xls.SetCellValue(292, 2, new TFormula("='Inputs advanced'!F300"));

            fmt = xls.GetCellVisibleFormatDef(292, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(292, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(292, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(292, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(292, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(292, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(293, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(293, 1, xls.AddFormat(fmt));
            xls.SetCellValue(293, 1, "Químicos para foliación");

            fmt = xls.GetCellVisibleFormatDef(293, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent4, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(293, 2, xls.AddFormat(fmt));
            xls.SetCellValue(293, 2, new TFormula("=IF('Inputs advanced'!F179=1,0,'Inputs advanced'!F301)"));

            fmt = xls.GetCellVisibleFormatDef(293, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(293, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(293, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(293, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(293, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(293, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(294, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(294, 1, xls.AddFormat(fmt));
            xls.SetCellValue(294, 1, "Otro(s) fertilizantes (s):");

            fmt = xls.GetCellVisibleFormatDef(294, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(294, 2, xls.AddFormat(fmt));
            xls.SetCellValue(294, 2, new TFormula("='Inputs advanced'!F302"));

            fmt = xls.GetCellVisibleFormatDef(294, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(294, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(294, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(294, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(294, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(294, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(295, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(295, 1, xls.AddFormat(fmt));
            xls.SetCellValue(295, 1, "Combustible:");

            fmt = xls.GetCellVisibleFormatDef(295, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(295, 2, xls.AddFormat(fmt));
            xls.SetCellValue(295, 2, new TFormula("='Inputs advanced'!F303"));

            fmt = xls.GetCellVisibleFormatDef(295, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(295, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(295, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(295, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(295, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(295, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(296, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(296, 1, xls.AddFormat(fmt));
            xls.SetCellValue(296, 1, "Total Fertilizaciones");

            fmt = xls.GetCellVisibleFormatDef(296, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(296, 2, xls.AddFormat(fmt));
            xls.SetCellValue(296, 2, new TFormula("=SUM(B279:B295)"));

            fmt = xls.GetCellVisibleFormatDef(296, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(296, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(296, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(296, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(296, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(296, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(297, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(297, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(297, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(297, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(298, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(298, 1, xls.AddFormat(fmt));
            xls.SetCellValue(298, 1, "Equipo  y materiales reutilizables");

            fmt = xls.GetCellVisibleFormatDef(298, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            fmt.WrapText = true;
            xls.SetCellFormat(298, 2, xls.AddFormat(fmt));
            xls.SetCellValue(298, 2, "Costo Total en Moneda local ");

            fmt = xls.GetCellVisibleFormatDef(298, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(298, 3, xls.AddFormat(fmt));
            xls.SetCellValue(298, 3, "Años de vida del Equipo (Vida util)");

            fmt = xls.GetCellVisibleFormatDef(298, 4);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(298, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(299, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(299, 1, xls.AddFormat(fmt));
            xls.SetCellValue(299, 1, "Herramientas generales");

            fmt = xls.GetCellVisibleFormatDef(299, 2);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(299, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(299, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(299, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(299, 4);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(299, 4, xls.AddFormat(fmt));
            xls.SetCellValue(300, 1, "Bomba manual ");

            fmt = xls.GetCellVisibleFormatDef(300, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(300, 2, xls.AddFormat(fmt));
            xls.SetCellValue(300, 2, new TFormula("='Inputs advanced'!F306"));

            fmt = xls.GetCellVisibleFormatDef(300, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(300, 3, xls.AddFormat(fmt));
            xls.SetCellValue(300, 3, new TFormula("='Inputs advanced'!F307"));

            fmt = xls.GetCellVisibleFormatDef(300, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(300, 4, xls.AddFormat(fmt));
            xls.SetCellValue(301, 1, "Machete");

            fmt = xls.GetCellVisibleFormatDef(301, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(301, 2, xls.AddFormat(fmt));
            xls.SetCellValue(301, 2, new TFormula("='Inputs advanced'!F308"));

            fmt = xls.GetCellVisibleFormatDef(301, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(301, 3, xls.AddFormat(fmt));
            xls.SetCellValue(301, 3, new TFormula("='Inputs advanced'!F309"));

            fmt = xls.GetCellVisibleFormatDef(301, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(301, 4, xls.AddFormat(fmt));
            xls.SetCellValue(302, 1, "Pala");

            fmt = xls.GetCellVisibleFormatDef(302, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(302, 2, xls.AddFormat(fmt));
            xls.SetCellValue(302, 2, new TFormula("='Inputs advanced'!F310"));

            fmt = xls.GetCellVisibleFormatDef(302, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(302, 3, xls.AddFormat(fmt));
            xls.SetCellValue(302, 3, new TFormula("='Inputs advanced'!F311"));

            fmt = xls.GetCellVisibleFormatDef(302, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(302, 4, xls.AddFormat(fmt));
            xls.SetCellValue(303, 1, "Azadón");

            fmt = xls.GetCellVisibleFormatDef(303, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(303, 2, xls.AddFormat(fmt));
            xls.SetCellValue(303, 2, new TFormula("='Inputs advanced'!F312"));

            fmt = xls.GetCellVisibleFormatDef(303, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(303, 3, xls.AddFormat(fmt));
            xls.SetCellValue(303, 3, new TFormula("='Inputs advanced'!F313"));

            fmt = xls.GetCellVisibleFormatDef(303, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(303, 4, xls.AddFormat(fmt));
            xls.SetCellValue(304, 1, "Carretilla");

            fmt = xls.GetCellVisibleFormatDef(304, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(304, 2, xls.AddFormat(fmt));
            xls.SetCellValue(304, 2, new TFormula("='Inputs advanced'!F314"));

            fmt = xls.GetCellVisibleFormatDef(304, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(304, 3, xls.AddFormat(fmt));
            xls.SetCellValue(304, 3, new TFormula("='Inputs advanced'!F315"));

            fmt = xls.GetCellVisibleFormatDef(304, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(304, 4, xls.AddFormat(fmt));
            xls.SetCellValue(305, 1, "Lima");

            fmt = xls.GetCellVisibleFormatDef(305, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(305, 2, xls.AddFormat(fmt));
            xls.SetCellValue(305, 2, new TFormula("='Inputs advanced'!F316"));

            fmt = xls.GetCellVisibleFormatDef(305, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(305, 3, xls.AddFormat(fmt));
            xls.SetCellValue(305, 3, new TFormula("='Inputs advanced'!F317"));

            fmt = xls.GetCellVisibleFormatDef(305, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(305, 4, xls.AddFormat(fmt));
            xls.SetCellValue(306, 1, "Chancha o ahoyador");

            fmt = xls.GetCellVisibleFormatDef(306, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(306, 2, xls.AddFormat(fmt));
            xls.SetCellValue(306, 2, new TFormula("='Inputs advanced'!F318"));

            fmt = xls.GetCellVisibleFormatDef(306, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(306, 3, xls.AddFormat(fmt));
            xls.SetCellValue(306, 3, new TFormula("='Inputs advanced'!F319"));

            fmt = xls.GetCellVisibleFormatDef(306, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(306, 4, xls.AddFormat(fmt));
            xls.SetCellValue(307, 1, "Barretón");

            fmt = xls.GetCellVisibleFormatDef(307, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(307, 2, xls.AddFormat(fmt));
            xls.SetCellValue(307, 2, new TFormula("='Inputs advanced'!F320"));

            fmt = xls.GetCellVisibleFormatDef(307, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(307, 3, xls.AddFormat(fmt));
            xls.SetCellValue(307, 3, new TFormula("='Inputs advanced'!F321"));

            fmt = xls.GetCellVisibleFormatDef(307, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(307, 4, xls.AddFormat(fmt));
            xls.SetCellValue(308, 1, "Mangueras");

            fmt = xls.GetCellVisibleFormatDef(308, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(308, 2, xls.AddFormat(fmt));
            xls.SetCellValue(308, 2, new TFormula("='Inputs advanced'!F322"));

            fmt = xls.GetCellVisibleFormatDef(308, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(308, 3, xls.AddFormat(fmt));
            xls.SetCellValue(308, 3, new TFormula("='Inputs advanced'!F323"));

            fmt = xls.GetCellVisibleFormatDef(308, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(308, 4, xls.AddFormat(fmt));
            xls.SetCellValue(309, 1, "Sistema de riego");

            fmt = xls.GetCellVisibleFormatDef(309, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(309, 2, xls.AddFormat(fmt));
            xls.SetCellValue(309, 2, new TFormula("='Inputs advanced'!F324"));

            fmt = xls.GetCellVisibleFormatDef(309, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(309, 3, xls.AddFormat(fmt));
            xls.SetCellValue(309, 3, new TFormula("='Inputs advanced'!F325"));

            fmt = xls.GetCellVisibleFormatDef(309, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(309, 4, xls.AddFormat(fmt));
            xls.SetCellValue(310, 1, "Motosierra");

            fmt = xls.GetCellVisibleFormatDef(310, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(310, 2, xls.AddFormat(fmt));
            xls.SetCellValue(310, 2, new TFormula("='Inputs advanced'!F326"));

            fmt = xls.GetCellVisibleFormatDef(310, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(310, 3, xls.AddFormat(fmt));
            xls.SetCellValue(310, 3, new TFormula("='Inputs advanced'!F327"));

            fmt = xls.GetCellVisibleFormatDef(310, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(310, 4, xls.AddFormat(fmt));
            xls.SetCellValue(311, 1, "Serrucho");

            fmt = xls.GetCellVisibleFormatDef(311, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(311, 2, xls.AddFormat(fmt));
            xls.SetCellValue(311, 2, new TFormula("='Inputs advanced'!F328"));

            fmt = xls.GetCellVisibleFormatDef(311, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(311, 3, xls.AddFormat(fmt));
            xls.SetCellValue(311, 3, new TFormula("='Inputs advanced'!F329"));

            fmt = xls.GetCellVisibleFormatDef(311, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(311, 4, xls.AddFormat(fmt));
            xls.SetCellValue(312, 1, "Bomba motor");

            fmt = xls.GetCellVisibleFormatDef(312, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(312, 2, xls.AddFormat(fmt));
            xls.SetCellValue(312, 2, new TFormula("='Inputs advanced'!F330"));

            fmt = xls.GetCellVisibleFormatDef(312, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(312, 3, xls.AddFormat(fmt));
            xls.SetCellValue(312, 3, new TFormula("='Inputs advanced'!F331"));

            fmt = xls.GetCellVisibleFormatDef(312, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(312, 4, xls.AddFormat(fmt));
            xls.SetCellValue(313, 1, "Tijeras Podar");

            fmt = xls.GetCellVisibleFormatDef(313, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(313, 2, xls.AddFormat(fmt));
            xls.SetCellValue(313, 2, new TFormula("='Inputs advanced'!F332"));

            fmt = xls.GetCellVisibleFormatDef(313, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(313, 3, xls.AddFormat(fmt));
            xls.SetCellValue(313, 3, new TFormula("='Inputs advanced'!F333"));

            fmt = xls.GetCellVisibleFormatDef(313, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(313, 4, xls.AddFormat(fmt));
            xls.SetCellValue(314, 1, "Hacha");

            fmt = xls.GetCellVisibleFormatDef(314, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(314, 2, xls.AddFormat(fmt));
            xls.SetCellValue(314, 2, new TFormula("='Inputs advanced'!F334"));

            fmt = xls.GetCellVisibleFormatDef(314, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.0";
            xls.SetCellFormat(314, 3, xls.AddFormat(fmt));
            xls.SetCellValue(314, 3, new TFormula("='Inputs advanced'!F335"));

            fmt = xls.GetCellVisibleFormatDef(314, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(314, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(315, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(315, 1, xls.AddFormat(fmt));
            xls.SetCellValue(315, 1, "Equipos para el beneficio");

            fmt = xls.GetCellVisibleFormatDef(315, 3);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(315, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(315, 4);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(315, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(315, 5);
            fmt.Format = "#,##0";
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(315, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(316, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(316, 1, xls.AddFormat(fmt));
            xls.SetCellValue(316, 1, "Beneficio humedo");

            fmt = xls.GetCellVisibleFormatDef(316, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(316, 5, xls.AddFormat(fmt));
            xls.SetCellValue(317, 1, "Despulpadora");

            fmt = xls.GetCellVisibleFormatDef(317, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(317, 2, xls.AddFormat(fmt));
            xls.SetCellValue(317, 2, new TFormula("='Inputs advanced'!F358"));

            fmt = xls.GetCellVisibleFormatDef(317, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(317, 3, xls.AddFormat(fmt));
            xls.SetCellValue(317, 3, new TFormula("='Inputs advanced'!F359"));

            fmt = xls.GetCellVisibleFormatDef(317, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(317, 4, xls.AddFormat(fmt));
            xls.SetCellValue(318, 1, "Sifon-Tolba");

            fmt = xls.GetCellVisibleFormatDef(318, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(318, 2, xls.AddFormat(fmt));
            xls.SetCellValue(318, 2, new TFormula("='Inputs advanced'!F360"));

            fmt = xls.GetCellVisibleFormatDef(318, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(318, 3, xls.AddFormat(fmt));
            xls.SetCellValue(318, 3, new TFormula("='Inputs advanced'!F361"));

            fmt = xls.GetCellVisibleFormatDef(318, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(318, 4, xls.AddFormat(fmt));
            xls.SetCellValue(319, 1, "Motor");

            fmt = xls.GetCellVisibleFormatDef(319, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(319, 2, xls.AddFormat(fmt));
            xls.SetCellValue(319, 2, new TFormula("='Inputs advanced'!F362"));

            fmt = xls.GetCellVisibleFormatDef(319, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(319, 3, xls.AddFormat(fmt));
            xls.SetCellValue(319, 3, new TFormula("='Inputs advanced'!F363"));

            fmt = xls.GetCellVisibleFormatDef(319, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(319, 4, xls.AddFormat(fmt));
            xls.SetCellValue(320, 1, "Tanques o pilas de fermentacion");

            fmt = xls.GetCellVisibleFormatDef(320, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(320, 2, xls.AddFormat(fmt));
            xls.SetCellValue(320, 2, new TFormula("='Inputs advanced'!F364"));

            fmt = xls.GetCellVisibleFormatDef(320, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(320, 3, xls.AddFormat(fmt));
            xls.SetCellValue(320, 3, new TFormula("='Inputs advanced'!F365"));

            fmt = xls.GetCellVisibleFormatDef(320, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(320, 4, xls.AddFormat(fmt));
            xls.SetCellValue(321, 1, "Canal de correo para lavar café");

            fmt = xls.GetCellVisibleFormatDef(321, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(321, 2, xls.AddFormat(fmt));
            xls.SetCellValue(321, 2, new TFormula("='Inputs advanced'!F366"));

            fmt = xls.GetCellVisibleFormatDef(321, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(321, 3, xls.AddFormat(fmt));
            xls.SetCellValue(321, 3, new TFormula("='Inputs advanced'!F367"));

            fmt = xls.GetCellVisibleFormatDef(321, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(321, 4, xls.AddFormat(fmt));
            xls.SetCellValue(322, 1, "Tubos PVC");

            fmt = xls.GetCellVisibleFormatDef(322, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(322, 2, xls.AddFormat(fmt));
            xls.SetCellValue(322, 2, new TFormula("='Inputs advanced'!F368"));

            fmt = xls.GetCellVisibleFormatDef(322, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(322, 3, xls.AddFormat(fmt));
            xls.SetCellValue(322, 3, new TFormula("='Inputs advanced'!F369"));

            fmt = xls.GetCellVisibleFormatDef(322, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(322, 4, xls.AddFormat(fmt));
            xls.SetCellValue(323, 1, "Sistema de filtración de agua (finca orgánica)");

            fmt = xls.GetCellVisibleFormatDef(323, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(323, 2, xls.AddFormat(fmt));
            xls.SetCellValue(323, 2, new TFormula("='Inputs advanced'!F370"));

            fmt = xls.GetCellVisibleFormatDef(323, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(323, 3, xls.AddFormat(fmt));
            xls.SetCellValue(323, 3, new TFormula("='Inputs advanced'!F371"));

            fmt = xls.GetCellVisibleFormatDef(323, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(323, 4, xls.AddFormat(fmt));
            xls.SetCellValue(324, 1, "Criba - Zaranda");

            fmt = xls.GetCellVisibleFormatDef(324, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(324, 2, xls.AddFormat(fmt));
            xls.SetCellValue(324, 2, new TFormula("='Inputs advanced'!F372"));

            fmt = xls.GetCellVisibleFormatDef(324, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(324, 3, xls.AddFormat(fmt));
            xls.SetCellValue(324, 3, new TFormula("='Inputs advanced'!F373"));

            fmt = xls.GetCellVisibleFormatDef(324, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(324, 4, xls.AddFormat(fmt));
            xls.SetCellValue(325, 1, "Desmucilagador");

            fmt = xls.GetCellVisibleFormatDef(325, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(325, 2, xls.AddFormat(fmt));
            xls.SetCellValue(325, 2, new TFormula("='Inputs advanced'!F374"));

            fmt = xls.GetCellVisibleFormatDef(325, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(325, 3, xls.AddFormat(fmt));
            xls.SetCellValue(325, 3, new TFormula("='Inputs advanced'!F375"));

            fmt = xls.GetCellVisibleFormatDef(325, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(325, 4, xls.AddFormat(fmt));
            xls.SetCellValue(326, 1, "Pozo");

            fmt = xls.GetCellVisibleFormatDef(326, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(326, 2, xls.AddFormat(fmt));
            xls.SetCellValue(326, 2, new TFormula("='Inputs advanced'!F376"));

            fmt = xls.GetCellVisibleFormatDef(326, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(326, 3, xls.AddFormat(fmt));
            xls.SetCellValue(326, 3, new TFormula("='Inputs advanced'!F377"));

            fmt = xls.GetCellVisibleFormatDef(326, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(326, 4, xls.AddFormat(fmt));
            xls.SetCellValue(327, 1, "Otro componente del beneficio húmedo");

            fmt = xls.GetCellVisibleFormatDef(327, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(327, 2, xls.AddFormat(fmt));
            xls.SetCellValue(327, 2, new TFormula("='Inputs advanced'!F378"));

            fmt = xls.GetCellVisibleFormatDef(327, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(327, 3, xls.AddFormat(fmt));
            xls.SetCellValue(327, 3, new TFormula("='Inputs advanced'!F379"));

            fmt = xls.GetCellVisibleFormatDef(327, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(327, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(328, 1);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(328, 1, xls.AddFormat(fmt));
            xls.SetCellValue(328, 1, "Beneficio seco");

            fmt = xls.GetCellVisibleFormatDef(328, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0";
            xls.SetCellFormat(328, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(328, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(328, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(328, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(328, 4, xls.AddFormat(fmt));
            xls.SetCellValue(329, 1, "Secador solar - Plancha concreto");

            fmt = xls.GetCellVisibleFormatDef(329, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(329, 2, xls.AddFormat(fmt));
            xls.SetCellValue(329, 2, new TFormula("='Inputs advanced'!F380"));

            fmt = xls.GetCellVisibleFormatDef(329, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(329, 3, xls.AddFormat(fmt));
            xls.SetCellValue(329, 3, new TFormula("='Inputs advanced'!F381"));

            fmt = xls.GetCellVisibleFormatDef(329, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(329, 4, xls.AddFormat(fmt));
            xls.SetCellValue(330, 1, "Plastico");

            fmt = xls.GetCellVisibleFormatDef(330, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(330, 2, xls.AddFormat(fmt));
            xls.SetCellValue(330, 2, new TFormula("='Inputs advanced'!F382"));

            fmt = xls.GetCellVisibleFormatDef(330, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(330, 3, xls.AddFormat(fmt));
            xls.SetCellValue(330, 3, new TFormula("='Inputs advanced'!F383"));

            fmt = xls.GetCellVisibleFormatDef(330, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(330, 4, xls.AddFormat(fmt));
            xls.SetCellValue(331, 1, "Rastrillo");

            fmt = xls.GetCellVisibleFormatDef(331, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(331, 2, xls.AddFormat(fmt));
            xls.SetCellValue(331, 2, new TFormula("='Inputs advanced'!F384"));

            fmt = xls.GetCellVisibleFormatDef(331, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(331, 3, xls.AddFormat(fmt));
            xls.SetCellValue(331, 3, new TFormula("='Inputs advanced'!F385"));
            xls.SetCellValue(332, 1, "Escoba");

            fmt = xls.GetCellVisibleFormatDef(332, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(332, 2, xls.AddFormat(fmt));
            xls.SetCellValue(332, 2, new TFormula("='Inputs advanced'!F386"));

            fmt = xls.GetCellVisibleFormatDef(332, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(332, 3, xls.AddFormat(fmt));
            xls.SetCellValue(332, 3, new TFormula("='Inputs advanced'!F387"));
            xls.SetCellValue(333, 1, "Bodega");

            fmt = xls.GetCellVisibleFormatDef(333, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(333, 2, xls.AddFormat(fmt));
            xls.SetCellValue(333, 2, new TFormula("='Inputs advanced'!F388"));

            fmt = xls.GetCellVisibleFormatDef(333, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(333, 3, xls.AddFormat(fmt));
            xls.SetCellValue(333, 3, new TFormula("='Inputs advanced'!F389"));

            fmt = xls.GetCellVisibleFormatDef(333, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(333, 4, xls.AddFormat(fmt));
            xls.SetCellValue(334, 1, "Otro componente del beneficio seco");

            fmt = xls.GetCellVisibleFormatDef(334, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0";
            xls.SetCellFormat(334, 2, xls.AddFormat(fmt));
            xls.SetCellValue(334, 2, new TFormula("='Inputs advanced'!F390"));

            fmt = xls.GetCellVisibleFormatDef(334, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(334, 3, xls.AddFormat(fmt));
            xls.SetCellValue(334, 3, new TFormula("='Inputs advanced'!F391"));

            fmt = xls.GetCellVisibleFormatDef(334, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(334, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(335, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(335, 1, xls.AddFormat(fmt));
            xls.SetCellValue(335, 1, "Otros equipos y/o materiales reutilizables");

            fmt = xls.GetCellVisibleFormatDef(335, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(335, 5, xls.AddFormat(fmt));
            xls.SetCellValue(336, 1, "Bascula o balanza");

            fmt = xls.GetCellVisibleFormatDef(336, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(336, 2, xls.AddFormat(fmt));
            xls.SetCellValue(336, 2, new TFormula("='Inputs advanced'!F337"));

            fmt = xls.GetCellVisibleFormatDef(336, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(336, 3, xls.AddFormat(fmt));
            xls.SetCellValue(336, 3, new TFormula("='Inputs advanced'!F338"));

            fmt = xls.GetCellVisibleFormatDef(336, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(336, 4, xls.AddFormat(fmt));
            xls.SetCellValue(337, 1, "Vehiculo o automovil para trabajo");

            fmt = xls.GetCellVisibleFormatDef(337, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(337, 2, xls.AddFormat(fmt));
            xls.SetCellValue(337, 2, new TFormula("='Inputs advanced'!F339"));

            fmt = xls.GetCellVisibleFormatDef(337, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(337, 3, xls.AddFormat(fmt));
            xls.SetCellValue(337, 3, new TFormula("='Inputs advanced'!F340"));

            fmt = xls.GetCellVisibleFormatDef(337, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(337, 4, xls.AddFormat(fmt));
            xls.SetCellValue(338, 1, "Animal de trabajo");

            fmt = xls.GetCellVisibleFormatDef(338, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(338, 2, xls.AddFormat(fmt));
            xls.SetCellValue(338, 2, new TFormula("='Inputs advanced'!F341"));

            fmt = xls.GetCellVisibleFormatDef(338, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(338, 3, xls.AddFormat(fmt));
            xls.SetCellValue(338, 3, new TFormula("='Inputs advanced'!F342"));

            fmt = xls.GetCellVisibleFormatDef(338, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(338, 4, xls.AddFormat(fmt));
            xls.SetCellValue(339, 1, "Sacos para la recoleccion");

            fmt = xls.GetCellVisibleFormatDef(339, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(339, 2, xls.AddFormat(fmt));
            xls.SetCellValue(339, 2, new TFormula("='Inputs advanced'!F345"));

            fmt = xls.GetCellVisibleFormatDef(339, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(339, 3, xls.AddFormat(fmt));
            xls.SetCellValue(339, 3, new TFormula("='Inputs advanced'!F346"));

            fmt = xls.GetCellVisibleFormatDef(339, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(339, 4, xls.AddFormat(fmt));
            xls.SetCellValue(340, 1, "Sacos Pergamino");

            fmt = xls.GetCellVisibleFormatDef(340, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(340, 2, xls.AddFormat(fmt));
            xls.SetCellValue(340, 2, new TFormula("='Inputs advanced'!F347"));

            fmt = xls.GetCellVisibleFormatDef(340, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(340, 3, xls.AddFormat(fmt));
            xls.SetCellValue(340, 3, new TFormula("='Inputs advanced'!F348"));

            fmt = xls.GetCellVisibleFormatDef(340, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(340, 4, xls.AddFormat(fmt));
            xls.SetCellValue(341, 1, "Cabuya:");

            fmt = xls.GetCellVisibleFormatDef(341, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(341, 2, xls.AddFormat(fmt));
            xls.SetCellValue(341, 2, new TFormula("='Inputs advanced'!F349"));

            fmt = xls.GetCellVisibleFormatDef(341, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(341, 3, xls.AddFormat(fmt));
            xls.SetCellValue(341, 3, new TFormula("='Inputs advanced'!F350"));

            fmt = xls.GetCellVisibleFormatDef(341, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(341, 4, xls.AddFormat(fmt));
            xls.SetCellValue(342, 1, "Canastas");

            fmt = xls.GetCellVisibleFormatDef(342, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(342, 2, xls.AddFormat(fmt));
            xls.SetCellValue(342, 2, new TFormula("='Inputs advanced'!F351"));

            fmt = xls.GetCellVisibleFormatDef(342, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(342, 3, xls.AddFormat(fmt));
            xls.SetCellValue(342, 3, new TFormula("='Inputs advanced'!F352"));

            fmt = xls.GetCellVisibleFormatDef(342, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(342, 4, xls.AddFormat(fmt));
            xls.SetCellValue(343, 1, "Cajas");

            fmt = xls.GetCellVisibleFormatDef(343, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(343, 2, xls.AddFormat(fmt));
            xls.SetCellValue(343, 2, new TFormula("='Inputs advanced'!F353"));

            fmt = xls.GetCellVisibleFormatDef(343, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(343, 3, xls.AddFormat(fmt));
            xls.SetCellValue(343, 3, new TFormula("='Inputs advanced'!F354"));

            fmt = xls.GetCellVisibleFormatDef(343, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(343, 4, xls.AddFormat(fmt));
            xls.SetCellValue(344, 1, "Motocicleta");

            fmt = xls.GetCellVisibleFormatDef(344, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(344, 2, xls.AddFormat(fmt));
            xls.SetCellValue(344, 2, new TFormula("='Inputs advanced'!F343"));

            fmt = xls.GetCellVisibleFormatDef(344, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(344, 3, xls.AddFormat(fmt));
            xls.SetCellValue(344, 3, new TFormula("='Inputs advanced'!F344"));

            fmt = xls.GetCellVisibleFormatDef(344, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(344, 4, xls.AddFormat(fmt));
            xls.SetCellValue(345, 1, "Otros");

            fmt = xls.GetCellVisibleFormatDef(345, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(345, 2, xls.AddFormat(fmt));
            xls.SetCellValue(345, 2, new TFormula("='Inputs advanced'!F355"));

            fmt = xls.GetCellVisibleFormatDef(345, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(345, 3, xls.AddFormat(fmt));
            xls.SetCellValue(345, 3, new TFormula("='Inputs advanced'!F356"));

            fmt = xls.GetCellVisibleFormatDef(345, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(345, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(346, 5);
            fmt.Format = "#,##0";
            xls.SetCellFormat(346, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(347, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(347, 1, xls.AddFormat(fmt));
            xls.SetCellValue(347, 1, "\"Servicios\" en el beneficio humedo y seco");

            fmt = xls.GetCellVisibleFormatDef(347, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(347, 2, xls.AddFormat(fmt));
            xls.SetCellValue(347, 2, "Cantidad necesaria");

            fmt = xls.GetCellVisibleFormatDef(347, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(347, 3, xls.AddFormat(fmt));
            xls.SetCellValue(347, 3, "Unidad de medida");

            fmt = xls.GetCellVisibleFormatDef(347, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(347, 4, xls.AddFormat(fmt));
            xls.SetCellValue(347, 4, "Costo por Unidad");

            fmt = xls.GetCellVisibleFormatDef(347, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(347, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(348, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(348, 1, xls.AddFormat(fmt));

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 71;
            fnt = xls.GetDefaultFont;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 77;
            fnt = xls.GetDefaultFont;
            fnt.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetCellValue(348, 1, new TRichString("Que tantos litros de agua se pueden gastar en el beneficio húmedo de 1 arroba de café"+ " pergamino seco?", Runs, xls));
            //We could also have used: xls.SetCellFromHtml(348, 1, "Que tantos litros de agua se pueden gastar en el beneficio h&uacute;medo de 1&nbsp;<font"
            //+" color = 'black'>arroba</font>&nbsp;de caf&eacute; pergamino seco?")


    fmt = xls.GetCellVisibleFormatDef(348, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(348, 2, xls.AddFormat(fmt));
            xls.SetCellValue(348, 2, new TFormula("='Inputs advanced'!F393"));

            fmt = xls.GetCellVisibleFormatDef(348, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(348, 3, xls.AddFormat(fmt));
            xls.SetCellValue(348, 3, "litros/quintal");

            fmt = xls.GetCellVisibleFormatDef(348, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(348, 4, xls.AddFormat(fmt));
            xls.SetCellValue(348, 4, new TFormula("='Inputs advanced'!F394"));

            fmt = xls.GetCellVisibleFormatDef(348, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(348, 5, xls.AddFormat(fmt));
            xls.SetCellValue(348, 5, new TFormula("=B348*D348"));
            xls.SetCellValue(348, 6, "litros por quintal de pergamino seco. 1 quintal 45.99 kilos  ");

            fmt = xls.GetCellVisibleFormatDef(349, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(349, 1, xls.AddFormat(fmt));
            xls.SetCellValue(349, 1, "Cuantos Kw consume en el secado de un quintal de pergaminos seco?");

            fmt = xls.GetCellVisibleFormatDef(349, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(349, 2, xls.AddFormat(fmt));
            xls.SetCellValue(349, 2, 0);

            fmt = xls.GetCellVisibleFormatDef(349, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(349, 3, xls.AddFormat(fmt));
            xls.SetCellValue(349, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(349, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(349, 4, xls.AddFormat(fmt));
            xls.SetCellValue(349, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(349, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(349, 5, xls.AddFormat(fmt));
            xls.SetCellValue(349, 5, new TFormula("='Inputs advanced'!F395"));
            xls.SetCellValue(349, 6, "kw ");

            fmt = xls.GetCellVisibleFormatDef(350, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(350, 1, xls.AddFormat(fmt));
            xls.SetCellValue(350, 1, "Que cantidad de combustible se gasta para secar un quintal de café pergamino?");

            fmt = xls.GetCellVisibleFormatDef(350, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(350, 2, xls.AddFormat(fmt));
            xls.SetCellValue(350, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(350, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(350, 3, xls.AddFormat(fmt));
            xls.SetCellValue(350, 3, ".");

            fmt = xls.GetCellVisibleFormatDef(350, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(350, 4, xls.AddFormat(fmt));
            xls.SetCellValue(350, 4, 0);

            fmt = xls.GetCellVisibleFormatDef(350, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(350, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(350, 6);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(350, 6, xls.AddFormat(fmt));
            xls.SetCellValue(350, 6, "galones/quintal");

            fmt = xls.GetCellVisibleFormatDef(351, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(351, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(352, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(352, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(353, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(353, 1, xls.AddFormat(fmt));
            xls.SetCellValue(353, 1, "Costos de Transporte");

            fmt = xls.GetCellVisibleFormatDef(353, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(353, 2, xls.AddFormat(fmt));
            xls.SetCellValue(353, 2, "Costo en transporte");

            fmt = xls.GetCellVisibleFormatDef(354, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(354, 1, xls.AddFormat(fmt));
            xls.SetCellValue(354, 1, "Semillero:");
            xls.SetCellValue(355, 1, "ir a comprar la semilla");

            fmt = xls.GetCellVisibleFormatDef(355, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(355, 2, xls.AddFormat(fmt));
            xls.SetCellValue(355, 2, new TFormula("='Inputs advanced'!F43"));

            fmt = xls.GetCellVisibleFormatDef(355, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(355, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(355, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(355, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(355, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(355, 5, xls.AddFormat(fmt));
            xls.SetCellValue(356, 1, "Llevada madera");

            fmt = xls.GetCellVisibleFormatDef(356, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "# ??/??";
            xls.SetCellFormat(356, 2, xls.AddFormat(fmt));
            xls.SetCellValue(356, 2, new TFormula("='Inputs advanced'!F44"));

            fmt = xls.GetCellVisibleFormatDef(356, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "# ??/??";
            xls.SetCellFormat(356, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(356, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(356, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(356, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(356, 5, xls.AddFormat(fmt));
            xls.SetCellValue(357, 1, "Llevada arena");

            fmt = xls.GetCellVisibleFormatDef(357, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(357, 2, xls.AddFormat(fmt));
            xls.SetCellValue(357, 2, new TFormula("='Inputs advanced'!F45"));

            fmt = xls.GetCellVisibleFormatDef(357, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(357, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(357, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(357, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(357, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(357, 5, xls.AddFormat(fmt));
            xls.SetCellValue(358, 1, "Otro(s):");

            fmt = xls.GetCellVisibleFormatDef(358, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(358, 2, xls.AddFormat(fmt));
            xls.SetCellValue(358, 2, new TFormula("='Inputs advanced'!F46"));

            fmt = xls.GetCellVisibleFormatDef(358, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(358, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(358, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(358, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(358, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(358, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(359, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(359, 1, xls.AddFormat(fmt));
            xls.SetCellValue(359, 1, "Vivero:");
            xls.SetCellValue(360, 1, "Jalada de tierra");

            fmt = xls.GetCellVisibleFormatDef(360, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(360, 2, xls.AddFormat(fmt));
            xls.SetCellValue(360, 2, new TFormula("='Inputs advanced'!F48"));

            fmt = xls.GetCellVisibleFormatDef(360, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(360, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(360, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(360, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(360, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(360, 5, xls.AddFormat(fmt));
            xls.SetCellValue(361, 1, "Ir a comprar bolsas y otros insumos para el vivero");

            fmt = xls.GetCellVisibleFormatDef(361, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(361, 2, xls.AddFormat(fmt));
            xls.SetCellValue(361, 2, new TFormula("='Inputs advanced'!F49"));

            fmt = xls.GetCellVisibleFormatDef(361, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(361, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(361, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(361, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(361, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(361, 5, xls.AddFormat(fmt));
            xls.SetCellValue(362, 1, "Otro(s)");

            fmt = xls.GetCellVisibleFormatDef(362, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(362, 2, xls.AddFormat(fmt));
            xls.SetCellValue(362, 2, new TFormula("='Inputs advanced'!F50"));

            fmt = xls.GetCellVisibleFormatDef(362, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(362, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(362, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(362, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(362, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(362, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(363, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(363, 1, xls.AddFormat(fmt));
            xls.SetCellValue(363, 1, "Preparación terreno y siembra:");
            xls.SetCellValue(364, 1, "Llevada de leña");

            fmt = xls.GetCellVisibleFormatDef(364, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(364, 2, xls.AddFormat(fmt));
            xls.SetCellValue(364, 2, new TFormula("='Inputs advanced'!F52"));

            fmt = xls.GetCellVisibleFormatDef(364, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(364, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(364, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(364, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(364, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(364, 5, xls.AddFormat(fmt));
            xls.SetCellValue(365, 1, "Lleva del abono");

            fmt = xls.GetCellVisibleFormatDef(365, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(365, 2, xls.AddFormat(fmt));
            xls.SetCellValue(365, 2, new TFormula("='Inputs advanced'!F53"));

            fmt = xls.GetCellVisibleFormatDef(365, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(365, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(365, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(365, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(365, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(365, 5, xls.AddFormat(fmt));
            xls.SetCellValue(366, 1, "Llevar plantas del vivero al campo");

            fmt = xls.GetCellVisibleFormatDef(366, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(366, 2, xls.AddFormat(fmt));
            xls.SetCellValue(366, 2, new TFormula("='Inputs advanced'!F54"));

            fmt = xls.GetCellVisibleFormatDef(366, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(366, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(366, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(366, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(366, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(366, 5, xls.AddFormat(fmt));
            xls.SetCellValue(367, 1, "Otro(s)");

            fmt = xls.GetCellVisibleFormatDef(367, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(367, 2, xls.AddFormat(fmt));
            xls.SetCellValue(367, 2, new TFormula("='Inputs advanced'!F55"));

            fmt = xls.GetCellVisibleFormatDef(367, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(367, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(367, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(367, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(367, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(367, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(368, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(368, 1, xls.AddFormat(fmt));
            xls.SetCellValue(368, 1, "Otros gastos en transporte en terminos anuales");
            xls.SetCellValue(369, 1, "Transporte equipo y herramientas");

            fmt = xls.GetCellVisibleFormatDef(369, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(369, 2, xls.AddFormat(fmt));
            xls.SetCellValue(369, 2, new TFormula("='Inputs advanced'!F57"));

            fmt = xls.GetCellVisibleFormatDef(369, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(369, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(369, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(369, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(369, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(369, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(370, 1);
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
            xls.SetCellFormat(370, 1, xls.AddFormat(fmt));
            xls.SetCellValue(370, 1, "Transporte mano de obra (no pagada en el jornal)");

            fmt = xls.GetCellVisibleFormatDef(370, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(370, 2, xls.AddFormat(fmt));
            xls.SetCellValue(370, 2, new TFormula("='Inputs advanced'!F58"));

            fmt = xls.GetCellVisibleFormatDef(370, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(370, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(370, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(370, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(370, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(370, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(371, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
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
            xls.SetCellFormat(371, 1, xls.AddFormat(fmt));
            xls.SetCellValue(371, 1, "Transporte de la cosecha al centro de acopio o asociación  ");

            fmt = xls.GetCellVisibleFormatDef(371, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(371, 2, xls.AddFormat(fmt));
            xls.SetCellValue(371, 2, new TFormula("='Inputs advanced'!F59"));

            fmt = xls.GetCellVisibleFormatDef(371, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(371, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(371, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(371, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(371, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(371, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(372, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(372, 1, xls.AddFormat(fmt));
            xls.SetCellValue(372, 1, "Transporte para ir a supervisas actividades (limpias, manejos, podas, obras conservación)");

            fmt = xls.GetCellVisibleFormatDef(372, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(372, 2, xls.AddFormat(fmt));
            xls.SetCellValue(372, 2, new TFormula("='Inputs advanced'!F60"));

            fmt = xls.GetCellVisibleFormatDef(372, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(372, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(372, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(372, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(372, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(372, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(373, 1);
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
            xls.SetCellFormat(373, 1, xls.AddFormat(fmt));
            xls.SetCellValue(373, 1, "Otro(s) transportes no considerados:");

            fmt = xls.GetCellVisibleFormatDef(373, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0.00";
            xls.SetCellFormat(373, 2, xls.AddFormat(fmt));
            xls.SetCellValue(373, 2, new TFormula("='Inputs advanced'!F61"));

            fmt = xls.GetCellVisibleFormatDef(373, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(373, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(373, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(373, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(373, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(373, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(374, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(374, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(375, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(375, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(377, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(377, 1, xls.AddFormat(fmt));
            xls.SetCellValue(377, 1, "Membresía a la Cooperativa");

            fmt = xls.GetCellVisibleFormatDef(378, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(378, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(379, 1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(379, 1, xls.AddFormat(fmt));
            xls.SetCellValue(379, 1, "Item");

            fmt = xls.GetCellVisibleFormatDef(379, 2);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(379, 2, xls.AddFormat(fmt));
            xls.SetCellValue(379, 2, "Al principio");

            fmt = xls.GetCellVisibleFormatDef(379, 3);
            fmt.Font.Style = TFlxFontStyles.Italic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(379, 3, xls.AddFormat(fmt));
            xls.SetCellValue(379, 3, "Anualmente");
            xls.SetCellValue(380, 1, "Costo de entrada");

            fmt = xls.GetCellVisibleFormatDef(380, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(380, 2, xls.AddFormat(fmt));
            xls.SetCellValue(380, 2, new TFormula("='Inputs advanced'!F398"));
            xls.SetCellValue(381, 1, "Membresía annual");

            fmt = xls.GetCellVisibleFormatDef(381, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(381, 3, xls.AddFormat(fmt));
            xls.SetCellValue(381, 3, new TFormula("='Inputs advanced'!F399"));
            xls.SetCellValue(382, 1, "Seguro de vida");

            fmt = xls.GetCellVisibleFormatDef(382, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(382, 3, xls.AddFormat(fmt));
            xls.SetCellValue(382, 3, new TFormula("='Inputs advanced'!F400"));

            fmt = xls.GetCellVisibleFormatDef(383, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(383, 1, xls.AddFormat(fmt));
            xls.SetCellValue(383, 1, "Pago por certificación Fair Trade");

            fmt = xls.GetCellVisibleFormatDef(383, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(383, 3, xls.AddFormat(fmt));
            xls.SetCellValue(383, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(384, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(384, 1, xls.AddFormat(fmt));
            xls.SetCellValue(384, 1, "Pago annual por certificación Orgánico");

            fmt = xls.GetCellVisibleFormatDef(384, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(384, 3, xls.AddFormat(fmt));
            xls.SetCellValue(384, 3, 0);

            fmt = xls.GetCellVisibleFormatDef(386, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x33, 0x66, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(386, 1, xls.AddFormat(fmt));
            xls.SetCellValue(386, 1, "Valor de la Tierra");

            fmt = xls.GetCellVisibleFormatDef(386, 2);
            fmt.Format = "0";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(386, 2, xls.AddFormat(fmt));
            xls.SetCellValue(386, 2, "1=yes");

            fmt = xls.GetCellVisibleFormatDef(387, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(387, 1, xls.AddFormat(fmt));
            xls.SetCellValue(387, 1, "Es usted propietario de su finca?");

            fmt = xls.GetCellVisibleFormatDef(387, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(387, 2, xls.AddFormat(fmt));
            xls.SetCellValue(387, 2, new TFormula("='Inputs advanced'!F402"));

            fmt = xls.GetCellVisibleFormatDef(388, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(388, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(388, 2);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(388, 2, xls.AddFormat(fmt));
            xls.SetCellValue(388, 2, "Total Value");

            fmt = xls.GetCellVisibleFormatDef(388, 3);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(388, 3, xls.AddFormat(fmt));
            xls.SetCellValue(388, 3, "Annual Opportunity Cost (4% interest)");
            xls.SetCellValue(389, 1, "Valor Por hectarea");

            fmt = xls.GetStyle(xls.GetBuiltInStyleName(TBuiltInStyle.Comma, 0), true);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Foreground2, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(389, 2, xls.AddFormat(fmt));
            xls.SetCellValue(389, 2, new TFormula("='Inputs advanced'!F403"));

            fmt = xls.GetCellVisibleFormatDef(389, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(389, 3, xls.AddFormat(fmt));
            xls.SetCellValue(389, 3, new TFormula("=B389*0.04"));
            xls.SetCellValue(389, 6, new TFormula("=B389/B67"));
            xls.SetCellValue(390, 1, "En total hectareas finca");

            fmt = xls.GetCellVisibleFormatDef(390, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(390, 2, xls.AddFormat(fmt));
            xls.SetCellValue(390, 2, new TFormula("=B389*B7"));

            fmt = xls.GetCellVisibleFormatDef(390, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(390, 3, xls.AddFormat(fmt));
            xls.SetCellValue(390, 3, new TFormula("=B390*0.04"));
            xls.SetCellValue(392, 6, new TFormula("=70000*H6"));

            fmt = xls.GetCellVisibleFormatDef(393, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(393, 1, xls.AddFormat(fmt));
            xls.SetCellValue(393, 1, "Cuanto paga anualmente por impuestos a la propiedad en  su finca?");

            fmt = xls.GetCellVisibleFormatDef(393, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(393, 2, xls.AddFormat(fmt));
            xls.SetCellValue(393, 2, 147.28);

            fmt = xls.GetCellVisibleFormatDef(395, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.WrapText = true;
            xls.SetCellFormat(395, 1, xls.AddFormat(fmt));
            xls.SetCellValue(395, 1, "En caso de no ser propietario,  paga alguna renta, cual es el valor ANUAL?");

            fmt = xls.GetCellVisibleFormatDef(395, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent5, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(395, 2, xls.AddFormat(fmt));
            xls.SetCellValue(395, 2, new TFormula("='Inputs advanced'!F404"));

            fmt = xls.GetCellVisibleFormatDef(398, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(398, 1, xls.AddFormat(fmt));
            xls.SetCellValue(398, 1, "Caracterización del productor");
            xls.SetCellValue(398, 2, "(Esta información es relevante para la narrativa)");

            fmt = xls.GetCellVisibleFormatDef(399, 1);
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
            fmt.WrapText = true;
            xls.SetCellFormat(399, 1, xls.AddFormat(fmt));
            xls.SetCellValue(399, 1, "¿Cuanto tiempo lleva usted en la actividad cafetera?");

            fmt = xls.GetCellVisibleFormatDef(399, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(399, 2, xls.AddFormat(fmt));
            xls.SetCellValue(399, 2, new TFormula("='Inputs advanced'!F146"));

            fmt = xls.GetCellVisibleFormatDef(400, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.WrapText = true;
            xls.SetCellFormat(400, 1, xls.AddFormat(fmt));
            xls.SetCellValue(400, 1, "¿Cuantos arboles de café por hectarea estima en su finca?");

            fmt = xls.GetCellVisibleFormatDef(400, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(400, 2, xls.AddFormat(fmt));
            xls.SetCellValue(400, 2, new TFormula("='Inputs advanced'!F149"));

            fmt = xls.GetCellVisibleFormatDef(403, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            xls.SetCellFormat(403, 1, xls.AddFormat(fmt));
            xls.SetCellValue(403, 1, "Metodos de Producción 1=si, 0=no");

            fmt = xls.GetCellVisibleFormatDef(404, 1);
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
            xls.SetCellFormat(404, 1, xls.AddFormat(fmt));
            xls.SetCellValue(404, 1, "Finca Quimica");

            fmt = xls.GetCellVisibleFormatDef(404, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(404, 2, xls.AddFormat(fmt));
            xls.SetCellValue(404, 2, new TFormula("='Inputs advanced'!F178"));

            fmt = xls.GetCellVisibleFormatDef(405, 1);
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
            xls.SetCellFormat(405, 1, xls.AddFormat(fmt));
            xls.SetCellValue(405, 1, "Finca Organica");

            fmt = xls.GetCellVisibleFormatDef(405, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(405, 2, xls.AddFormat(fmt));
            xls.SetCellValue(405, 2, new TFormula("='Inputs advanced'!F179"));

            fmt = xls.GetCellVisibleFormatDef(406, 1);
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
            xls.SetCellFormat(406, 1, xls.AddFormat(fmt));
            xls.SetCellValue(406, 1, "Transición");

            fmt = xls.GetCellVisibleFormatDef(406, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(406, 2, xls.AddFormat(fmt));
            xls.SetCellValue(406, 2, new TFormula("='Inputs advanced'!F180"));

            fmt = xls.GetCellVisibleFormatDef(408, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
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
            xls.SetCellFormat(408, 1, xls.AddFormat(fmt));
            xls.SetCellValue(408, 1, "Tipo de café producido y llevado a la cooperativa (Porcentaje):");

            fmt = xls.GetCellVisibleFormatDef(409, 1);
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
            xls.SetCellFormat(409, 1, xls.AddFormat(fmt));
            xls.SetCellValue(409, 1, "Uva / Cereza");

            fmt = xls.GetCellVisibleFormatDef(409, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(409, 2, xls.AddFormat(fmt));
            xls.SetCellValue(409, 2, new TFormula("='Inputs advanced'!F182"));

            fmt = xls.GetCellVisibleFormatDef(410, 1);
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
            xls.SetCellFormat(410, 1, xls.AddFormat(fmt));
            xls.SetCellValue(410, 1, "Pergamino húmedo");

            fmt = xls.GetCellVisibleFormatDef(410, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(410, 2, xls.AddFormat(fmt));
            xls.SetCellValue(410, 2, new TFormula("='Inputs advanced'!F183"));

            fmt = xls.GetCellVisibleFormatDef(411, 1);
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
            xls.SetCellFormat(411, 1, xls.AddFormat(fmt));
            xls.SetCellValue(411, 1, "Pergamino seco");

            fmt = xls.GetCellVisibleFormatDef(411, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(411, 2, xls.AddFormat(fmt));
            xls.SetCellValue(411, 2, new TFormula("='Inputs advanced'!F184"));

            fmt = xls.GetCellVisibleFormatDef(412, 1);
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
            xls.SetCellFormat(412, 1, xls.AddFormat(fmt));
            xls.SetCellValue(412, 1, "Trillado");

            fmt = xls.GetCellVisibleFormatDef(412, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(412, 2, xls.AddFormat(fmt));
            xls.SetCellValue(412, 2, new TFormula("='Inputs advanced'!F185"));

            fmt = xls.GetCellVisibleFormatDef(413, 1);
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
            xls.SetCellFormat(413, 1, xls.AddFormat(fmt));
            xls.SetCellValue(413, 1, "Otro:");

            fmt = xls.GetCellVisibleFormatDef(413, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(413, 2, xls.AddFormat(fmt));
            xls.SetCellValue(413, 2, ".");

            fmt = xls.GetCellVisibleFormatDef(416, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(416, 1, xls.AddFormat(fmt));
            xls.SetCellValue(416, 1, "Sabe si al momento de llevar su café a la cooperativa recibió algun premio ? (Ejemplo:"
            + " Fair Trade o comercio justo, orgánico, priemio de calidad). Especifique.");

            fmt = xls.GetCellVisibleFormatDef(416, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(416, 2, xls.AddFormat(fmt));
            xls.SetCellValue(416, 2, new TFormula("='Inputs advanced'!F197"));

            fmt = xls.GetCellVisibleFormatDef(418, 1);
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
            fmt.WrapText = true;
            xls.SetCellFormat(418, 1, xls.AddFormat(fmt));
            xls.SetCellValue(418, 1, "Es usted propietario de su finca?                                          ((1=si,"
            + " 0=no, \".\" =  No sabe, no responde)");

            fmt = xls.GetCellVisibleFormatDef(418, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(418, 2, xls.AddFormat(fmt));
            xls.SetCellValue(418, 2, new TFormula("=$B$387"));

            fmt = xls.GetCellVisibleFormatDef(421, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(421, 1, xls.AddFormat(fmt));
            xls.SetCellValue(421, 1, "¿Se ha visto su finca fuertmente afectada por roya en algún año en particular? ¿Cuál"
            + " ANO?");

            fmt = xls.GetCellVisibleFormatDef(421, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(421, 2, xls.AddFormat(fmt));
            xls.SetCellValue(421, 2, new TFormula("='Inputs advanced'!F206"));

            fmt = xls.GetCellVisibleFormatDef(422, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(422, 1, xls.AddFormat(fmt));
            xls.SetCellValue(422, 1, "¿En cuanto se redujo su producción como consecuencia de los efectos de la roya?. Porcentaje");

            fmt = xls.GetCellVisibleFormatDef(422, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(422, 2, xls.AddFormat(fmt));
            xls.SetCellValue(422, 2, new TFormula("='Inputs advanced'!F207"));

            fmt = xls.GetCellVisibleFormatDef(423, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(423, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(423, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(423, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(424, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(424, 1, xls.AddFormat(fmt));
            xls.SetCellValue(424, 1, "¿Que alternativas utilizó UD. Para sobrepasar el choque en los ingresos que la roya"
            + " representó?                                       Escriba ingreso aproximado ( 0"
            + " = No)");

            fmt = xls.GetCellVisibleFormatDef(424, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(424, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(425, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(425, 1, xls.AddFormat(fmt));
            xls.SetCellValue(425, 1, "Préstamos");

            fmt = xls.GetCellVisibleFormatDef(425, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(425, 2, xls.AddFormat(fmt));
            xls.SetCellValue(425, 2, new TFormula("='Inputs advanced'!F209"));

            fmt = xls.GetCellVisibleFormatDef(425, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(425, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(426, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(426, 1, xls.AddFormat(fmt));
            xls.SetCellValue(426, 1, "Venta de activos (Lotes, tierra, acciones en la asociación)");

            fmt = xls.GetCellVisibleFormatDef(426, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(426, 2, xls.AddFormat(fmt));
            xls.SetCellValue(426, 2, new TFormula("='Inputs advanced'!F210"));

            fmt = xls.GetCellVisibleFormatDef(426, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(426, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(427, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
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
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(427, 1, xls.AddFormat(fmt));
            xls.SetCellValue(427, 1, "Trabajo particular");

            fmt = xls.GetCellVisibleFormatDef(427, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(427, 2, xls.AddFormat(fmt));
            xls.SetCellValue(427, 2, new TFormula("='Inputs advanced'!F211"));

            fmt = xls.GetCellVisibleFormatDef(427, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(427, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(428, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
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
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(428, 1, xls.AddFormat(fmt));
            xls.SetCellValue(428, 1, "Uso de ahorros");

            fmt = xls.GetCellVisibleFormatDef(428, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(428, 2, xls.AddFormat(fmt));
            xls.SetCellValue(428, 2, new TFormula("='Inputs advanced'!F212"));

            fmt = xls.GetCellVisibleFormatDef(428, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(428, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(429, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(429, 1, xls.AddFormat(fmt));
            xls.SetCellValue(429, 1, "Renovación a otras variedades. ¿Cuál?");

            fmt = xls.GetCellVisibleFormatDef(429, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(429, 2, xls.AddFormat(fmt));
            xls.SetCellValue(429, 2, new TFormula("='Inputs advanced'!F213"));

            fmt = xls.GetCellVisibleFormatDef(429, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(429, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(430, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Thin;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(430, 1, xls.AddFormat(fmt));
            xls.SetCellValue(430, 1, "Transición químico a orgánico");

            fmt = xls.GetCellVisibleFormatDef(430, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(430, 2, xls.AddFormat(fmt));
            xls.SetCellValue(430, 2, new TFormula("='Inputs advanced'!F214"));

            fmt = xls.GetCellVisibleFormatDef(430, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(430, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(431, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
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
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(431, 1, xls.AddFormat(fmt));
            xls.SetCellValue(431, 1, "Transición orgánico a químico");

            fmt = xls.GetCellVisibleFormatDef(431, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(431, 2, xls.AddFormat(fmt));
            xls.SetCellValue(431, 2, new TFormula("='Inputs advanced'!F215"));

            fmt = xls.GetCellVisibleFormatDef(431, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(431, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(432, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
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
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(432, 1, xls.AddFormat(fmt));
            xls.SetCellValue(432, 1, "Paso a otro cultivo");

            fmt = xls.GetCellVisibleFormatDef(432, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(432, 2, xls.AddFormat(fmt));
            xls.SetCellValue(432, 2, new TFormula("='Inputs advanced'!F216"));

            fmt = xls.GetCellVisibleFormatDef(433, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Size20 = 280;
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
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(433, 1, xls.AddFormat(fmt));
            xls.SetCellValue(433, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(433, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0";
            xls.SetCellFormat(433, 2, xls.AddFormat(fmt));
            xls.SetCellValue(433, 2, new TFormula("='Inputs advanced'!F217"));

            fmt = xls.GetCellVisibleFormatDef(444, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(444, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(444, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(444, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(444, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(444, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(445, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(445, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(445, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(445, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(445, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(445, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(446, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.WrapText = true;
            xls.SetCellFormat(446, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(446, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(446, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(446, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(446, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(447, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.top;
            fmt.WrapText = true;
            xls.SetCellFormat(447, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(447, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(447, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(447, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(447, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(448, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(448, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(448, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(448, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(449, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(449, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(449, 3);
            fmt.Format = "0%";
            xls.SetCellFormat(449, 3, xls.AddFormat(fmt));

            //Comments

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Tahoma";
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 23;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Tahoma";
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetComment(148, 3, new TRichString("Adriana Ramírez Flores:\naquí no funciona pero se lleva a como 15 celdas abajo", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(148, 3, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 143, 206, 4, 63, 147, 121, 4, 719);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Adriana Ramírez Flores:\naquí no funciona pero se lleva a como 15 celdas abajo", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(148, 3, CommentProps);

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
            fnt = xls.GetDefaultFont;
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 14;
            fnt = xls.GetDefaultFont;
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetComment(183, 4, new TRichString("Luisa Escobar:\n En la pregunta queda implicito que se refiere a la productividad"
            + " de la variedad por hectarea de cada productor, lo cual asumiria que este esta asumiendo"
            + " las proporciones respectivas de cada variedad)", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            CommentProps = TCommentProperties.CreateStandard(183, 4, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 194, 24, 5, 647, 194, 24, 6, 199);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Luisa Escobar:\n En la pregunta queda implicito que se refiere a la productividad"
            //+" de la variedad por hectarea de cada productor, lo cual asumiria que este esta asumiendo"
            //+ " las proporciones respectivas de cada variedad)", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(183, 4, CommentProps);

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Tahoma";
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Style = TFlxFontStyles.Bold;
            fnt.Scheme = TFontScheme.None;
            Runs[0].FontIndex = xls.AddFont(fnt);
            Runs[1].FirstChar = 23;
            fnt = xls.GetDefaultFont;
            fnt.Name = "Tahoma";
            fnt.Size20 = 180;
            fnt.Color = TExcelColor.Automatic;
            fnt.Scheme = TFontScheme.None;
            Runs[1].FontIndex = xls.AddFont(fnt);
            xls.SetComment(349, 1, new TRichString("Adriana Ramírez Flores:\nesta puede ser el recibo de luz\n", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            CommentProps = TCommentProperties.CreateStandard(349, 1, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 349, 73, 2, 72, 350, 109, 2, 823);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Adriana Ramírez Flores:\nesta puede ser el recibo de luz\n", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(349, 1, CommentProps);

            //Cell selection and scroll position.
            xls.SelectCell(19, 25, false);

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
