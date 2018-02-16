using CoffeeCore.DTO;
using CoffeeCore.Interfaces;
using FlexCel.Core;
using FlexCel.XlsAdapter;
using System;
using System.Collections.Generic;
using System.IO;


namespace CoffeeInfrastructure.Flexcel
{
    public class Output
    {
        public void Outcome(ExcelFile xls, TWorkspace workspace)
        {
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

            xls.ActiveSheet = 2;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsCheckCompatibility = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Outcome 1.0";

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
            xls.DefaultColWidth = 2272;

            xls.SetColWidth(1, 2, 2272);    //(8.13 + 0.75) * 256

            xls.SetColWidth(3, 3, 3968);    //(14.75 + 0.75) * 256

            xls.SetColWidth(4, 5, 2272);    //(8.13 + 0.75) * 256

            xls.SetColWidth(6, 6, 3072);    //(11.25 + 0.75) * 256

            xls.SetColWidth(7, 8, 2272);    //(8.13 + 0.75) * 256

            xls.SetColWidth(9, 9, 1504);    //(5.13 + 0.75) * 256

            xls.SetColWidth(10, 10, 1760);    //(6.13 + 0.75) * 256

            xls.SetColWidth(11, 11, 2400);    //(8.63 + 0.75) * 256

            xls.SetColWidth(12, 13, 1504);    //(5.13 + 0.75) * 256

            xls.SetColWidth(14, 16384, 2272);    //(8.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(1, 330);    //16.50 * 20
            xls.SetRowHeight(7, 420);    //21.00 * 20
            xls.SetRowHeight(19, 330);    //16.50 * 20
            xls.SetRowHeight(23, 330);    //16.50 * 20
            xls.SetRowHeight(30, 375);    //18.75 * 20
            xls.SetRowHeight(32, 465);    //23.25 * 20
            xls.SetRowHeight(33, 465);    //23.25 * 20
            xls.SetRowHeight(35, 330);    //16.50 * 20

            //Merged Cells
            xls.MergeCells(20, 21, 23, 24);
            xls.MergeCells(7, 21, 7, 29);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(2, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 4);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 5);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 6);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 7);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 8);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 9);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 10);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 11);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 12);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 13);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 14);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 15);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 16);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 17);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 18);
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(2, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(3, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(4, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 21);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 21, xls.AddFormat(fmt));
            xls.SetCellValue(7, 21, "(Note to Programmer)");

            fmt = xls.GetCellVisibleFormatDef(7, 22);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 23);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 24);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 25);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 25, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 26);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 26, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 27);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 27, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 28);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 28, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 29);
            fmt.Font.Size20 = 320;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(7, 29, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 19, xls.AddFormat(fmt));
            xls.SetCellValue(9, 21, "Please add in the graph the blue line according to the following linked value");
            xls.SetCellValue(9, 29, new TFormula("=DATABASE_Schema!$J$15"));

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 19, xls.AddFormat(fmt));
            xls.SetCellValue(11, 21, "Please add to the graph the red line according to the price of coffee per pound in"
            + " the ");

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 19, xls.AddFormat(fmt));
            xls.SetCellValue(12, 21, "stock market");
            xls.SetCellValue(12, 29, 1.34);

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(13, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(14, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 19, xls.AddFormat(fmt));
            xls.SetCellValue(15, 21, "Y axis: US/POUND");

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(16, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(17, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(20, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 21);
            fmt.Font.Size20 = 480;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(20, 21, xls.AddFormat(fmt));
            xls.SetCellValue(20, 21, "SCREEN");

            fmt = xls.GetCellVisibleFormatDef(20, 22);
            fmt.Font.Size20 = 480;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(20, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 23);
            fmt.Font.Size20 = 480;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(20, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 24);
            fmt.Font.Size20 = 480;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Top.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Top.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(20, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(21, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 21);
            fmt.Font.Size20 = 480;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(21, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 22);
            fmt.Font.Size20 = 480;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(21, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 23);
            fmt.Font.Size20 = 480;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(21, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 24);
            fmt.Font.Size20 = 480;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(21, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 21);
            fmt.Font.Size20 = 480;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(22, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 22);
            fmt.Font.Size20 = 480;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(22, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 23);
            fmt.Font.Size20 = 480;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(22, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 24);
            fmt.Font.Size20 = 480;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(22, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 21);
            fmt.Font.Size20 = 480;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(23, 21, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 22);
            fmt.Font.Size20 = 480;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(23, 22, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 23);
            fmt.Font.Size20 = 480;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(23, 23, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 24);
            fmt.Font.Size20 = 480;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(23, 24, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(25, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(28, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(29, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 10);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(30, 10, xls.AddFormat(fmt));
            xls.SetCellValue(30, 10, "US/ht");

            fmt = xls.GetCellVisibleFormatDef(30, 11);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(30, 11, xls.AddFormat(fmt));
            xls.SetCellValue(30, 11, "Soles/ht");

            fmt = xls.GetCellVisibleFormatDef(30, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 3);
            fmt.Font.Size20 = 360;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 3, xls.AddFormat(fmt));
            xls.SetCellValue(32, 3, "Your variable cost of production is: ");

            fmt = xls.GetCellVisibleFormatDef(32, 4);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 5);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 6);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 7);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 8);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 9);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 10);
            fmt.Font.Size20 = 280;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(32, 10, xls.AddFormat(fmt));
            xls.SetCellValue(32, 10, new TFormula("=DATABASE_Schema!$F$35"));

            fmt = xls.GetCellVisibleFormatDef(32, 11);
            fmt.Font.Size20 = 280;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(32, 11, xls.AddFormat(fmt));
            xls.SetCellValue(32, 11, new TFormula("=DATABASE_Schema!$G$35"));

            fmt = xls.GetCellVisibleFormatDef(32, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(32, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 3);
            fmt.Font.Size20 = 360;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));
            xls.SetCellValue(33, 3, "Your total cost of production is: ");

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 5);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 6);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 7);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 8);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 9);
            fmt.Font.Size20 = 200;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 10);
            fmt.Font.Size20 = 280;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(33, 10, xls.AddFormat(fmt));
            xls.SetCellValue(33, 10, new TFormula("=DATABASE_Schema!$H$35"));

            fmt = xls.GetCellVisibleFormatDef(33, 11);
            fmt.Font.Size20 = 280;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "0";
            xls.SetCellFormat(33, 11, xls.AddFormat(fmt));
            xls.SetCellValue(33, 11, new TFormula("=DATABASE_Schema!$I$35"));

            fmt = xls.GetCellVisibleFormatDef(33, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(33, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 15);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 16);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 17);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 18);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(34, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 8);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 9);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 10);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 11);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 12);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 13);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 14);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 15);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 16);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 17);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 18);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 19);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(36, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(36, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(36, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(37, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(37, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(37, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(38, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(38, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(39, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 7);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(39, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 5);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.000";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(41, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(41, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(41, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(42, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(42, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(42, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(43, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(43, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(43, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(44, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(44, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(44, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(45, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(45, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(45, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(45, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 4);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(46, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 5);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(46, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 6);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(46, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 7);
            fmt.Font.Size20 = 200;
            xls.SetCellFormat(46, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.Font.Size20 = 200;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 7);
            fmt.Font.Size20 = 200;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(47, 7, xls.AddFormat(fmt));

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
            xls.SetComment(9, 29, new TRichString("Juan Hernandez:\nIt will change according to user input", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            TCommentProperties CommentProps = TCommentProperties.CreateStandard(9, 29, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 7, 164, 29, 577, 12, 24, 31, 548);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nIt will change according to user input", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(9, 29, CommentProps);

            Runs = new TRTFRun[2];
            Runs[0].FirstChar = 0;
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
            xls.SetComment(12, 29, new TRichString("Juan Hernandez:\nFeel free to attach this number to a reliable source that update"
            + " each day.\nFrom now I am taking from:\nhttp://markets.businessinsider.com/commodities/coffee-price", Runs, xls));

            //You probably don't need to call the lines below. This code is needed only if you want to change the comment box properties like color or default location
            CommentProps = TCommentProperties.CreateStandard(12, 29, xls);
            CommentProps.Anchor = new TClientAnchor(TFlxAnchorType.DontMoveAndDontResize, 10, 170, 29, 577, 17, 97, 34, 635);

            //Excel by doesn't autofit the comment box so it can hold all text.
            //There is an option in TCommentProperties, but if you use it Excel will show the text in a single line.
            //To have FlexCel autofit the comment for you, you can do it with the following code:

            //    CommentProps.Anchor = xls.AutofitComment(new TRichString("Juan Hernandez:\nFeel free to attach this number to a reliable source that update"
            //+" each day.\nFrom now I am taking from:\nhttp://markets.businessinsider.com/commodities/coffee-price", Runs, xls), 1.5, true, 1.1, 0, CommentProps.Anchor);

            xls.SetCommentProperties(12, 29, CommentProps);

            //Cell selection and scroll position.
            xls.SelectCell(30, 22, false);
            xls.ScrollWindow(1, 2);

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
