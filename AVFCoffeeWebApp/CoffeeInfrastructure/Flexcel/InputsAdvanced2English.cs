using FlexCel.Core;
using System;
using System.Collections.Generic;
using System.Text;

namespace CoffeeInfrastructure.Flexcel
{
    public class InputsAdvanced2English
    {
        public void InputAdvanced2English(ExcelFile xls)
        {
            //xls.NewFile(38, TExcelFileFormat.v2016);    //Create a new Excel file with 38 sheets.

            //Set the names of the sheets
            xls.ActiveSheet = 1;
            xls.SheetName = "Language";
            xls.ActiveSheet = 2;
            xls.SheetName = "Metrics Spanish";
            xls.ActiveSheet = 3;
            xls.SheetName = "Metrics English";
            xls.ActiveSheet = 4;
            xls.SheetName = "Inputs 1.0_Spa";
            xls.ActiveSheet = 5;
            xls.SheetName = "Inputs 1.0_Eng";
            xls.ActiveSheet = 6;
            xls.SheetName = "Inputs advance 2.0_Spa";
            xls.ActiveSheet = 7;
            xls.SheetName = "Inputs advance 2.0_Eng";
            xls.ActiveSheet = 8;
            xls.SheetName = "Outcome 1.0";
            xls.ActiveSheet = 9;
            xls.SheetName = "Additional 2.0";
            xls.ActiveSheet = 10;
            xls.SheetName = "Fixed 2.0";
            xls.ActiveSheet = 11;
            xls.SheetName = "Variable 2.0";
            xls.ActiveSheet = 12;
            xls.SheetName = "General Budget 2.0";
            xls.ActiveSheet = 13;
            xls.SheetName = "DATABASE_Schema";
            xls.ActiveSheet = 14;
            xls.SheetName = "Metrics";
            xls.ActiveSheet = 15;
            xls.SheetName = "Inputs 1.0";
            xls.ActiveSheet = 16;
            xls.SheetName = "Inputs advance 2.0";
            xls.ActiveSheet = 17;
            xls.SheetName = "Inputs 2.0 Conv. default values";
            xls.ActiveSheet = 18;
            xls.SheetName = "Inputs 2.0 Conv. new inputs";
            xls.ActiveSheet = 19;
            xls.SheetName = "Inputs TOT advanced";
            xls.ActiveSheet = 20;
            xls.SheetName = "Gral Conf. Summary_Spa";
            xls.ActiveSheet = 21;
            xls.SheetName = "Gral Conf. Summary";
            xls.ActiveSheet = 22;
            xls.SheetName = "Inputs 1.0 default values";
            xls.ActiveSheet = 23;
            xls.SheetName = "Inputs 1.0 Conv. new values";
            xls.ActiveSheet = 24;
            xls.SheetName = "Outcome TOTAL_Adj";
            xls.ActiveSheet = 25;
            xls.SheetName = "Outcome_Y_Adjustment";
            xls.ActiveSheet = 26;
            xls.SheetName = "Outcome_L Adjustment";
            xls.ActiveSheet = 27;
            xls.SheetName = "Proportions";
            xls.ActiveSheet = 28;
            xls.SheetName = "Budget_Supuestos";
            xls.ActiveSheet = 29;
            xls.SheetName = "Budget_Equipo";
            xls.ActiveSheet = 30;
            xls.SheetName = "Budget_M Obra";
            xls.ActiveSheet = 31;
            xls.SheetName = "Budget_Presupuesto";
            xls.ActiveSheet = 32;
            xls.SheetName = "Budget_Valor de M Obra";
            xls.ActiveSheet = 33;
            xls.SheetName = "Budget_Establecimiento";
            xls.ActiveSheet = 34;
            xls.SheetName = "Budget_Sostenemiento";
            xls.ActiveSheet = 35;
            xls.SheetName = "Outcome 1.0 pre_metric_currency";
            xls.ActiveSheet = 36;
            xls.SheetName = "Conversiones";
            xls.ActiveSheet = 37;
            xls.SheetName = "Proporción de productividad";
            xls.ActiveSheet = 38;
            xls.SheetName = "Inputs 1.0 (Ref)";

            xls.ActiveSheet = 7;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Inputs advance 2.0_Eng";

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
            Range = new TXlsNamedRange(RangeName, 33, 32, "=Budget_Establecimiento!$A$3:$C$53");
            //You could also use: Range = new TXlsNamedRange(RangeName, 33, 33, 3, 1, 53, 3, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 30, 32, "='Budget_M Obra'!$A$1:$K$86");
            //You could also use: Range = new TXlsNamedRange(RangeName, 30, 30, 1, 1, 86, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 31, 32, "=Budget_Presupuesto!$A$34:$J$46");
            //You could also use: Range = new TXlsNamedRange(RangeName, 31, 31, 34, 1, 46, 10, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 34, 32, "=Budget_Sostenemiento!$A$1:$K$44");
            //You could also use: Range = new TXlsNamedRange(RangeName, 34, 34, 1, 1, 44, 11, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 28, 32, "=Budget_Supuestos!$A$276:$G$297");
            //You could also use: Range = new TXlsNamedRange(RangeName, 28, 28, 276, 1, 297, 7, 32);
            xls.SetNamedRange(Range);

            RangeName = TXlsNamedRange.GetInternalName(InternalNameRange.Print_Area);
            Range = new TXlsNamedRange(RangeName, 32, 32, "='Budget_Valor de M Obra'!$A$2:$J$85");
            //You could also use: Range = new TXlsNamedRange(RangeName, 32, 32, 2, 1, 85, 10, 32);
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

            xls.SetColWidth(1, 1, 2272);    //(8.13 + 0.75) * 256

            xls.SetColWidth(2, 2, 21408);    //(82.88 + 0.75) * 256

            TFlxFormat ColFmt;
            ColFmt = xls.GetFormat(xls.GetColFormat(2));
            ColFmt.VAlignment = TVFlxAlignment.center;
            xls.SetColFormat(2, 2, xls.AddFormat(ColFmt));

            xls.SetColWidth(3, 3, 4640);    //(17.38 + 0.75) * 256

            xls.SetColWidth(4, 4, 5472);    //(20.63 + 0.75) * 256

            xls.SetColWidth(5, 7, 9728);    //(37.25 + 0.75) * 256

            xls.SetColWidth(8, 10, 2272);    //(8.13 + 0.75) * 256

            xls.SetColWidth(11, 16384, 2272);    //(8.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(4, 390);    //19.50 * 20
            xls.SetRowHeight(5, 360);    //18.00 * 20

            TFlxFormat RowFmt;
            RowFmt = xls.GetFormat(xls.GetRowFormat(5));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(5, xls.AddFormat(RowFmt));
            xls.SetRowHeight(9, 619);    //30.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(11));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(11, xls.AddFormat(RowFmt));
            xls.SetRowHeight(12, 300);    //15.00 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(12));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(12, xls.AddFormat(RowFmt));
            xls.SetRowHeight(13, 402);    //20.10 * 20
            xls.SetRowHeight(17, 522);    //26.10 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(18));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(18, xls.AddFormat(RowFmt));
            xls.SetRowHeight(29, 582);    //29.10 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(29));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(29, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(41));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(41, xls.AddFormat(RowFmt));
            xls.SetRowHeight(46, 499);    //24.95 * 20
            xls.SetRowHeight(47, 402);    //20.10 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(47));
            RowFmt.Font.Color = TExcelColor.Automatic;
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(47, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(48));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(48, xls.AddFormat(RowFmt));
            xls.SetRowHeight(58, 499);    //24.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(59));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(59, xls.AddFormat(RowFmt));
            xls.SetRowHeight(60, 630);    //31.50 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(62));
            RowFmt.Font.Color = TExcelColor.Automatic;
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(62, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(72));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(72, xls.AddFormat(RowFmt));
            xls.SetRowHeight(76, 480);    //24.00 * 20
            xls.SetRowHeight(82, 522);    //26.10 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(83));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(83, xls.AddFormat(RowFmt));
            xls.SetRowHeight(84, 630);    //31.50 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(86));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(86, xls.AddFormat(RowFmt));
            xls.SetRowHeight(89, 559);    //27.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(96));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(96, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(107));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(107, xls.AddFormat(RowFmt));
            xls.SetRowHeight(108, 630);    //31.50 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(110));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(110, xls.AddFormat(RowFmt));
            xls.SetRowHeight(120, 420);    //21.00 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(120));
            RowFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            RowFmt.Font.Style = TFlxFontStyles.Bold;
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(120, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(121));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(121, xls.AddFormat(RowFmt));
            xls.SetRowHeight(122, 630);    //31.50 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(125));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(125, xls.AddFormat(RowFmt));
            xls.SetRowHeight(126, 630);    //31.50 * 20
            xls.SetRowHeight(129, 630);    //31.50 * 20
            xls.SetRowHeight(132, 439);    //21.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(132));
            RowFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(132, xls.AddFormat(RowFmt));
            xls.SetRowHeight(133, 882);    //44.10 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(133));
            RowFmt.Font.Color = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(133, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(134));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(134, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(142));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(142, xls.AddFormat(RowFmt));
            xls.SetRowHeight(149, 739);    //36.95 * 20
            xls.SetRowHeight(156, 379);    //18.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(156));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(156, xls.AddFormat(RowFmt));
            xls.SetRowHeight(157, 522);    //26.10 * 20
            xls.SetRowHeight(158, 522);    //26.10 * 20
            xls.SetRowHeight(159, 379);    //18.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(159));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetRowFormat(159, xls.AddFormat(RowFmt));
            xls.SetRowHeight(160, 402);    //20.10 * 20
            xls.SetRowHeight(161, 522);    //26.10 * 20
            xls.SetRowHeight(162, 522);    //26.10 * 20
            xls.SetRowHeight(163, 522);    //26.10 * 20
            xls.SetRowHeight(165, 600);    //30.00 * 20
            xls.SetRowHeight(166, 600);    //30.00 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(168));
            RowFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(168, xls.AddFormat(RowFmt));
            xls.SetRowHeight(169, 859);    //42.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(169));
            RowFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(169, xls.AddFormat(RowFmt));
            xls.SetRowHeight(170, 499);    //24.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(170));
            RowFmt.Font.Color = TExcelColor.Automatic;
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(170, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(201));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(201, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(222));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(222, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(257));
            RowFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(257, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(258));
            RowFmt.Font.Color = TExcelColor.Automatic;
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(258, xls.AddFormat(RowFmt));
            xls.SetRowHeight(259, 499);    //24.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(259));
            RowFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            RowFmt.HAlignment = THFlxAlignment.center;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(259, xls.AddFormat(RowFmt));
            xls.SetRowHeight(260, 499);    //24.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(260));
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(260, xls.AddFormat(RowFmt));
            xls.SetRowHeight(261, 499);    //24.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(261));
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(261, xls.AddFormat(RowFmt));
            xls.SetRowHeight(262, 499);    //24.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(262));
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(262, xls.AddFormat(RowFmt));
            xls.SetRowHeight(263, 499);    //24.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(263));
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(263, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(264));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(264, xls.AddFormat(RowFmt));
            xls.SetRowHeight(266, 439);    //21.95 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(267));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(267, xls.AddFormat(RowFmt));
            xls.SetRowHeight(268, 630);    //31.50 * 20
            xls.SetRowHeight(269, 945);    //47.25 * 20
            xls.SetRowHeight(270, 630);    //31.50 * 20
            xls.SetRowHeight(271, 630);    //31.50 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(272));
            RowFmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(272, xls.AddFormat(RowFmt));
            xls.SetRowHeight(273, 630);    //31.50 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(273));
            RowFmt.Font.Color = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.center;
            xls.SetRowFormat(273, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(274));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.left;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(274, xls.AddFormat(RowFmt));
            xls.SetRowHeight(277, 510);    //25.50 * 20

            RowFmt = xls.GetFormat(xls.GetRowFormat(279));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.left;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(279, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(283));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(283, xls.AddFormat(RowFmt));

            RowFmt = xls.GetFormat(xls.GetRowFormat(288));
            RowFmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            RowFmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            RowFmt.FillPattern.BgColor = TExcelColor.Automatic;
            RowFmt.HAlignment = THFlxAlignment.left;
            RowFmt.VAlignment = TVFlxAlignment.center;
            xls.SetRowFormat(288, xls.AddFormat(RowFmt));
            xls.SetRowHeight(290, 379);    //18.95 * 20
            xls.SetRowHeight(291, 480);    //24.00 * 20
            xls.SetRowHeight(293, 330);    //16.50 * 20

            //Merged Cells
            xls.MergeCells(4, 2, 4, 3);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(4, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 2, xls.AddFormat(fmt));
            xls.SetCellValue(4, 2, "INPUTS ADVANCE");

            fmt = xls.GetCellVisibleFormatDef(4, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 4);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            xls.SetCellFormat(4, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 5);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            xls.SetCellFormat(4, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(4, 6);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0xFF, 0xFF);
            xls.SetCellFormat(4, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
            xls.SetCellValue(5, 2, "Labor");

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, new TFormula("=+\"Please, describe in days how much time is invested in the next activities for"
            + " ONE \"&'Gral Conf. Summary'!$I$23&\" of coffee\""));

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, "Each working day is represents six hours of effective work  (Ex: 3 hours = 0.5 days"
            + " ;  12 hours = 2 days)");

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));
            xls.SetCellValue(8, 2, "In addition, the total number of days is equal to:  Number of people * Days * Number"
            + " of times per year");

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));
            xls.SetCellValue(9, 2, "Ex: If one activity requires 2 people, working 1 day and this activity is performed"
            + " 3 times per year,  then total days = 2*1*3 =6");

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));
            xls.SetCellValue(9, 5, "Note for Rishi & Eric");

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));
            xls.SetCellValue(10, 2, "Write 0 if the activity is not done.");

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFC, 0xD5, 0xB4);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));
            xls.SetCellValue(10, 5, "Fist level");

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));
            xls.SetCellValue(11, 2, "Labor during establishment and vegetative growth years");

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));
            xls.SetCellValue(11, 5, "Second level");

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));
            xls.SetCellValue(12, 2, "Germinator Labor ");

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));
            xls.SetCellValue(12, 5, "Third level");

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));
            xls.SetCellValue(13, 2, "Seed collection");

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));
            xls.SetCellValue(13, 3, new TFormula("='Inputs 2.0 Conv. default values'!I13"));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x92, 0xCD, 0xDC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));
            xls.SetCellValue(13, 5, "Fourth level");

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));
            xls.SetCellValue(14, 2, "Seed selection");

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));
            xls.SetCellValue(14, 3, new TFormula("='Inputs 2.0 Conv. default values'!I14"));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xC0, 0xDA);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));
            xls.SetCellValue(14, 5, "Fifth level");

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));
            xls.SetCellValue(15, 2, "Germinator construction");

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));
            xls.SetCellValue(15, 3, new TFormula("='Inputs 2.0 Conv. default values'!I15"));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));
            xls.SetCellValue(15, 5, "Sixth level");

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));
            xls.SetCellValue(16, 2, "Germinator maintenance - Irrigation");

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));
            xls.SetCellValue(16, 3, new TFormula("='Inputs 2.0 Conv. default values'!I16"));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, "Other");

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));
            xls.SetCellValue(17, 3, new TFormula("='Inputs 2.0 Conv. default values'!I17"));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));
            xls.SetCellValue(18, 2, "Nursery labor ");

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));
            xls.SetCellValue(19, 2, "Nursery construction");

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));
            xls.SetCellValue(19, 3, new TFormula("='Inputs 2.0 Conv. default values'!I19"));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));
            xls.SetCellValue(20, 2, "Nursery soil transport");

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));
            xls.SetCellValue(20, 3, new TFormula("='Inputs 2.0 Conv. default values'!I20"));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, "Nursery weeding");

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));
            xls.SetCellValue(21, 3, new TFormula("='Inputs 2.0 Conv. default values'!I21"));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));
            xls.SetCellValue(22, 2, "Compost mix for bags");

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));
            xls.SetCellValue(22, 3, new TFormula("='Inputs 2.0 Conv. default values'!I22"));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));
            xls.SetCellValue(23, 2, "Seedling bags filling");

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));
            xls.SetCellValue(23, 3, new TFormula("='Inputs 2.0 Conv. default values'!I23"));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));
            xls.SetCellValue(24, 2, "Seedling sowing");

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));
            xls.SetCellValue(24, 3, new TFormula("='Inputs 2.0 Conv. default values'!I24"));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));
            xls.SetCellValue(25, 2, "Irrigation");

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));
            xls.SetCellValue(25, 3, new TFormula("='Inputs 2.0 Conv. default values'!I25"));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));
            xls.SetCellValue(26, 2, "Organic foliar spraying");

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));
            xls.SetCellValue(26, 3, new TFormula("='Inputs 2.0 Conv. default values'!I26"));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));
            xls.SetCellValue(27, 2, "Seedling replanting");

            fmt = xls.GetCellVisibleFormatDef(27, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));
            xls.SetCellValue(27, 3, new TFormula("='Inputs 2.0 Conv. default values'!I27"));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));
            xls.SetCellValue(28, 2, "Other");

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));
            xls.SetCellValue(28, 3, new TFormula("='Inputs 2.0 Conv. default values'!I28"));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, "Land preparation and sowing labor");

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));
            xls.SetCellValue(30, 2, "Field cleaning");

            fmt = xls.GetCellVisibleFormatDef(30, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));
            xls.SetCellValue(30, 3, new TFormula("='Inputs 2.0 Conv. default values'!I30"));

            fmt = xls.GetCellVisibleFormatDef(31, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));
            xls.SetCellValue(31, 2, "Old coffee trees cutting or other timber");

            fmt = xls.GetCellVisibleFormatDef(31, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));
            xls.SetCellValue(31, 3, new TFormula("='Inputs 2.0 Conv. default values'!I31"));

            fmt = xls.GetCellVisibleFormatDef(32, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(32, 2, xls.AddFormat(fmt));
            xls.SetCellValue(32, 2, "Wood collection");

            fmt = xls.GetCellVisibleFormatDef(32, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(32, 3, xls.AddFormat(fmt));
            xls.SetCellValue(32, 3, new TFormula("='Inputs 2.0 Conv. default values'!I32"));

            fmt = xls.GetCellVisibleFormatDef(33, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, "Wood chopping");

            fmt = xls.GetCellVisibleFormatDef(33, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));
            xls.SetCellValue(33, 3, new TFormula("='Inputs 2.0 Conv. default values'!I33"));

            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));
            xls.SetCellValue(34, 2, "Coffee and shade layout");

            fmt = xls.GetCellVisibleFormatDef(34, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));
            xls.SetCellValue(34, 3, new TFormula("='Inputs 2.0 Conv. default values'!I34"));

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));
            xls.SetCellValue(35, 2, "Hole digging");

            fmt = xls.GetCellVisibleFormatDef(35, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));
            xls.SetCellValue(35, 3, new TFormula("='Inputs 2.0 Conv. default values'!I35"));

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));
            xls.SetCellValue(36, 2, "Seedling transportation to the plot");

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));
            xls.SetCellValue(36, 3, new TFormula("='Inputs 2.0 Conv. default values'!I36"));

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));
            xls.SetCellValue(37, 2, "Seedling transplant");

            fmt = xls.GetCellVisibleFormatDef(37, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));
            xls.SetCellValue(37, 3, new TFormula("='Inputs 2.0 Conv. default values'!I37"));

            fmt = xls.GetCellVisibleFormatDef(38, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));
            xls.SetCellValue(38, 2, "Shade adjustment");

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));
            xls.SetCellValue(38, 3, new TFormula("='Inputs 2.0 Conv. default values'!I38"));

            fmt = xls.GetCellVisibleFormatDef(39, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));
            xls.SetCellValue(39, 2, "Compost mixing");

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));
            xls.SetCellValue(39, 3, new TFormula("='Inputs 2.0 Conv. default values'!I39"));

            fmt = xls.GetCellVisibleFormatDef(40, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));
            xls.SetCellValue(40, 2, "Others ");

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));
            xls.SetCellValue(40, 3, new TFormula("='Inputs 2.0 Conv. default values'!I40"));

            fmt = xls.GetCellVisibleFormatDef(41, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(41, 2, xls.AddFormat(fmt));
            xls.SetCellValue(41, 2, "Labor for the year corresponding to vegetative growth");

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));
            xls.SetCellValue(42, 2, "Weeding");

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));
            xls.SetCellValue(42, 3, new TFormula("='Inputs 2.0 Conv. default values'!I42"));

            fmt = xls.GetCellVisibleFormatDef(43, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(43, 2, xls.AddFormat(fmt));
            xls.SetCellValue(43, 2, "Organic fertilization");

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));
            xls.SetCellValue(43, 3, new TFormula("='Inputs 2.0 Conv. default values'!I43"));

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));
            xls.SetCellValue(44, 2, "Chemical fertilization");

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));
            xls.SetCellValue(44, 3, new TFormula("='Inputs 2.0 Conv. default values'!I44"));

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));
            xls.SetCellValue(45, 2, "Foliar spraying for fertilization and rust control");

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));
            xls.SetCellValue(45, 3, new TFormula("='Inputs 2.0 Conv. default values'!I45"));

            fmt = xls.GetCellVisibleFormatDef(46, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(46, 2, xls.AddFormat(fmt));
            xls.SetCellValue(46, 2, "Other:");

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));
            xls.SetCellValue(46, 3, new TFormula("='Inputs 2.0 Conv. default values'!I46"));

            fmt = xls.GetCellVisibleFormatDef(47, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));
            xls.SetCellValue(47, 2, "Labor for farm maintenance, harvesting and procesing");

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 10);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(47, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(48, 2, xls.AddFormat(fmt));
            xls.SetCellValue(48, 2, "Labor for maintenance when the coffee trees are young");

            fmt = xls.GetCellVisibleFormatDef(48, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(48, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(49, 2, xls.AddFormat(fmt));
            xls.SetCellValue(49, 2, "Manual weeding");

            fmt = xls.GetCellVisibleFormatDef(49, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(49, 3, xls.AddFormat(fmt));
            xls.SetCellValue(49, 3, new TFormula("='Inputs 2.0 Conv. default values'!I49"));

            fmt = xls.GetCellVisibleFormatDef(50, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(50, 2, xls.AddFormat(fmt));
            xls.SetCellValue(50, 2, "Chemical weeding");

            fmt = xls.GetCellVisibleFormatDef(50, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(50, 3, xls.AddFormat(fmt));
            xls.SetCellValue(50, 3, new TFormula("='Inputs 2.0 Conv. default values'!I50"));

            fmt = xls.GetCellVisibleFormatDef(51, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(51, 2, xls.AddFormat(fmt));
            xls.SetCellValue(51, 2, "Organic fertilizers for maintenance");

            fmt = xls.GetCellVisibleFormatDef(51, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(51, 3, xls.AddFormat(fmt));
            xls.SetCellValue(51, 3, new TFormula("='Inputs 2.0 Conv. default values'!I51"));

            fmt = xls.GetCellVisibleFormatDef(52, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(52, 2, xls.AddFormat(fmt));
            xls.SetCellValue(52, 2, "Chemical fertilizers for maintenance");

            fmt = xls.GetCellVisibleFormatDef(52, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(52, 3, xls.AddFormat(fmt));
            xls.SetCellValue(52, 3, new TFormula("='Inputs 2.0 Conv. default values'!I52"));

            fmt = xls.GetCellVisibleFormatDef(53, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(53, 2, xls.AddFormat(fmt));
            xls.SetCellValue(53, 2, "Foliar spraying and rust control");

            fmt = xls.GetCellVisibleFormatDef(53, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(53, 3, xls.AddFormat(fmt));
            xls.SetCellValue(53, 3, new TFormula("='Inputs 2.0 Conv. default values'!I53"));

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));
            xls.SetCellValue(54, 2, "Hedgerows construction");

            fmt = xls.GetCellVisibleFormatDef(54, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(54, 3, xls.AddFormat(fmt));
            xls.SetCellValue(54, 3, new TFormula("='Inputs 2.0 Conv. default values'!I54"));

            fmt = xls.GetCellVisibleFormatDef(55, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(55, 2, xls.AddFormat(fmt));
            xls.SetCellValue(55, 2, "Shade tree pruning (maintenance) ");

            fmt = xls.GetCellVisibleFormatDef(55, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(55, 3, xls.AddFormat(fmt));
            xls.SetCellValue(55, 3, new TFormula("='Inputs 2.0 Conv. default values'!I55"));

            fmt = xls.GetCellVisibleFormatDef(56, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(56, 2, xls.AddFormat(fmt));
            xls.SetCellValue(56, 2, "Pest control (broca: fumigation, trap setting, etc.)");

            fmt = xls.GetCellVisibleFormatDef(56, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(56, 3, xls.AddFormat(fmt));
            xls.SetCellValue(56, 3, new TFormula("='Inputs 2.0 Conv. default values'!I56"));

            fmt = xls.GetCellVisibleFormatDef(57, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(57, 2, xls.AddFormat(fmt));
            xls.SetCellValue(57, 2, "Coffee growing management (pruning - agobio)");

            fmt = xls.GetCellVisibleFormatDef(57, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(57, 3, xls.AddFormat(fmt));
            xls.SetCellValue(57, 3, new TFormula("='Inputs 2.0 Conv. default values'!I57"));

            fmt = xls.GetCellVisibleFormatDef(58, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(58, 2, xls.AddFormat(fmt));
            xls.SetCellValue(58, 2, "Others:");

            fmt = xls.GetCellVisibleFormatDef(58, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(58, 3, xls.AddFormat(fmt));
            xls.SetCellValue(58, 3, new TFormula("='Inputs 2.0 Conv. default values'!I58"));

            fmt = xls.GetCellVisibleFormatDef(59, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(59, 2, xls.AddFormat(fmt));
            xls.SetCellValue(59, 2, "Labor for harvest when the coffee trees are young");

            fmt = xls.GetCellVisibleFormatDef(59, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(59, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(60, 2, xls.AddFormat(fmt));
            xls.SetCellValue(60, 2, "Tolal number of days invested to collect coffee                                  "
            + "                                                                                 "
            + "                                                                           Remember"
            + " total number of days is equal to:  Number of people*Days * Number of times per year");

            fmt = xls.GetCellVisibleFormatDef(60, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(60, 3, xls.AddFormat(fmt));
            xls.SetCellValue(60, 3, new TFormula("='Inputs 2.0 Conv. default values'!I60"));

            fmt = xls.GetCellVisibleFormatDef(61, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(61, 2, xls.AddFormat(fmt));
            xls.SetCellValue(61, 2, "Additional days invested in other activities related with the harvest ");

            fmt = xls.GetCellVisibleFormatDef(61, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(61, 3, xls.AddFormat(fmt));
            xls.SetCellValue(61, 3, new TFormula("='Inputs 2.0 Conv. default values'!I61"));

            fmt = xls.GetCellVisibleFormatDef(62, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(62, 2, xls.AddFormat(fmt));
            xls.SetCellValue(62, 2, "Labor for procesing when the coffee trees are young");

            fmt = xls.GetCellVisibleFormatDef(62, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(62, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 10);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(62, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(63, 2, xls.AddFormat(fmt));
            xls.SetCellValue(63, 2, "Pulp separation and fermentation (work time)");

            fmt = xls.GetCellVisibleFormatDef(63, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(63, 3, xls.AddFormat(fmt));
            xls.SetCellValue(63, 3, new TFormula("='Inputs 2.0 Conv. default values'!I63"));

            fmt = xls.GetCellVisibleFormatDef(64, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(64, 2, xls.AddFormat(fmt));
            xls.SetCellValue(64, 2, "Washing");

            fmt = xls.GetCellVisibleFormatDef(64, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(64, 3, xls.AddFormat(fmt));
            xls.SetCellValue(64, 3, new TFormula("='Inputs 2.0 Conv. default values'!I64"));

            fmt = xls.GetCellVisibleFormatDef(65, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(65, 2, xls.AddFormat(fmt));
            xls.SetCellValue(65, 2, "Drying");

            fmt = xls.GetCellVisibleFormatDef(65, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(65, 3, xls.AddFormat(fmt));
            xls.SetCellValue(65, 3, new TFormula("='Inputs 2.0 Conv. default values'!I65"));

            fmt = xls.GetCellVisibleFormatDef(66, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(66, 2, xls.AddFormat(fmt));
            xls.SetCellValue(66, 2, "Screening / shaking");

            fmt = xls.GetCellVisibleFormatDef(66, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(66, 3, xls.AddFormat(fmt));
            xls.SetCellValue(66, 3, new TFormula("='Inputs 2.0 Conv. default values'!I66"));

            fmt = xls.GetCellVisibleFormatDef(67, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(67, 2, xls.AddFormat(fmt));
            xls.SetCellValue(67, 2, "Selection / picking");

            fmt = xls.GetCellVisibleFormatDef(67, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(67, 3, xls.AddFormat(fmt));
            xls.SetCellValue(67, 3, new TFormula("='Inputs 2.0 Conv. default values'!I67"));

            fmt = xls.GetCellVisibleFormatDef(68, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(68, 2, xls.AddFormat(fmt));
            xls.SetCellValue(68, 2, "Storage");

            fmt = xls.GetCellVisibleFormatDef(68, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(68, 3, xls.AddFormat(fmt));
            xls.SetCellValue(68, 3, new TFormula("='Inputs 2.0 Conv. default values'!I68"));

            fmt = xls.GetCellVisibleFormatDef(69, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(69, 2, xls.AddFormat(fmt));
            xls.SetCellValue(69, 2, "Management of coffee wastewater");

            fmt = xls.GetCellVisibleFormatDef(69, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(69, 3, xls.AddFormat(fmt));
            xls.SetCellValue(69, 3, new TFormula("='Inputs 2.0 Conv. default values'!I69"));

            fmt = xls.GetCellVisibleFormatDef(70, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(70, 2, xls.AddFormat(fmt));
            xls.SetCellValue(70, 2, "Pulp management");

            fmt = xls.GetCellVisibleFormatDef(70, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(70, 3, xls.AddFormat(fmt));
            xls.SetCellValue(70, 3, new TFormula("='Inputs 2.0 Conv. default values'!I70"));

            fmt = xls.GetCellVisibleFormatDef(71, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(71, 2, xls.AddFormat(fmt));
            xls.SetCellValue(71, 2, "Other activities of the processing/beneficio:");

            fmt = xls.GetCellVisibleFormatDef(71, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(71, 3, xls.AddFormat(fmt));
            xls.SetCellValue(71, 3, new TFormula("='Inputs 2.0 Conv. default values'!I71"));

            fmt = xls.GetCellVisibleFormatDef(72, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(72, 2, xls.AddFormat(fmt));
            xls.SetCellValue(72, 2, "Labor for maintenance when the coffee trees are mature");

            fmt = xls.GetCellVisibleFormatDef(72, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(72, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(73, 2, xls.AddFormat(fmt));
            xls.SetCellValue(73, 2, "Manual weeding");

            fmt = xls.GetCellVisibleFormatDef(73, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(73, 3, xls.AddFormat(fmt));
            xls.SetCellValue(73, 3, new TFormula("='Inputs 2.0 Conv. default values'!I73"));

            fmt = xls.GetCellVisibleFormatDef(74, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(74, 2, xls.AddFormat(fmt));
            xls.SetCellValue(74, 2, "Chemical weeding");

            fmt = xls.GetCellVisibleFormatDef(74, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(74, 3, xls.AddFormat(fmt));
            xls.SetCellValue(74, 3, new TFormula("='Inputs 2.0 Conv. default values'!I74"));

            fmt = xls.GetCellVisibleFormatDef(75, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(75, 2, xls.AddFormat(fmt));
            xls.SetCellValue(75, 2, "Organic fertilizers for maintenance");

            fmt = xls.GetCellVisibleFormatDef(75, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(75, 3, xls.AddFormat(fmt));
            xls.SetCellValue(75, 3, new TFormula("='Inputs 2.0 Conv. default values'!I75"));

            fmt = xls.GetCellVisibleFormatDef(76, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(76, 2, xls.AddFormat(fmt));
            xls.SetCellValue(76, 2, "Chemical fertilizers for maintenance");

            fmt = xls.GetCellVisibleFormatDef(76, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(76, 3, xls.AddFormat(fmt));
            xls.SetCellValue(76, 3, new TFormula("='Inputs 2.0 Conv. default values'!I76"));

            fmt = xls.GetCellVisibleFormatDef(77, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(77, 2, xls.AddFormat(fmt));
            xls.SetCellValue(77, 2, "Foliar spraying and rust control");

            fmt = xls.GetCellVisibleFormatDef(77, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(77, 3, xls.AddFormat(fmt));
            xls.SetCellValue(77, 3, new TFormula("='Inputs 2.0 Conv. default values'!I77"));

            fmt = xls.GetCellVisibleFormatDef(78, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(78, 2, xls.AddFormat(fmt));
            xls.SetCellValue(78, 2, "Hedgerows construction");

            fmt = xls.GetCellVisibleFormatDef(78, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(78, 3, xls.AddFormat(fmt));
            xls.SetCellValue(78, 3, new TFormula("='Inputs 2.0 Conv. default values'!I78"));

            fmt = xls.GetCellVisibleFormatDef(79, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(79, 2, xls.AddFormat(fmt));
            xls.SetCellValue(79, 2, "Shade tree pruning (maintenance) ");

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));
            xls.SetCellValue(79, 3, new TFormula("='Inputs 2.0 Conv. default values'!I79"));

            fmt = xls.GetCellVisibleFormatDef(80, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(80, 2, xls.AddFormat(fmt));
            xls.SetCellValue(80, 2, "Pest control (broca: fumigation, trap setting, etc.)");

            fmt = xls.GetCellVisibleFormatDef(80, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(80, 3, xls.AddFormat(fmt));
            xls.SetCellValue(80, 3, new TFormula("='Inputs 2.0 Conv. default values'!I80"));

            fmt = xls.GetCellVisibleFormatDef(81, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(81, 2, xls.AddFormat(fmt));
            xls.SetCellValue(81, 2, "Coffee growing management (pruning - agobio)");

            fmt = xls.GetCellVisibleFormatDef(81, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(81, 3, xls.AddFormat(fmt));
            xls.SetCellValue(81, 3, new TFormula("='Inputs 2.0 Conv. default values'!I81"));

            fmt = xls.GetCellVisibleFormatDef(82, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(82, 2, xls.AddFormat(fmt));
            xls.SetCellValue(82, 2, "Others:");

            fmt = xls.GetCellVisibleFormatDef(82, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(82, 3, xls.AddFormat(fmt));
            xls.SetCellValue(82, 3, new TFormula("='Inputs 2.0 Conv. default values'!I82"));

            fmt = xls.GetCellVisibleFormatDef(83, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(83, 2, xls.AddFormat(fmt));
            xls.SetCellValue(83, 2, "Labor for harvest when the coffee trees are mature");

            fmt = xls.GetCellVisibleFormatDef(83, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(83, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(84, 2, xls.AddFormat(fmt));
            xls.SetCellValue(84, 2, "Tolal number of days invested to collect coffee                                  "
            + "                                                                                 "
            + "                                                                           Remember"
            + " total number of days is equal to:  Number of people*Days * Number of times per year");

            fmt = xls.GetCellVisibleFormatDef(84, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(84, 3, xls.AddFormat(fmt));
            xls.SetCellValue(84, 3, new TFormula("='Inputs 2.0 Conv. default values'!I84"));

            fmt = xls.GetCellVisibleFormatDef(85, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(85, 2, xls.AddFormat(fmt));
            xls.SetCellValue(85, 2, "Additional days invested in other activities related with the harvest ");

            fmt = xls.GetCellVisibleFormatDef(85, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(85, 3, xls.AddFormat(fmt));
            xls.SetCellValue(85, 3, new TFormula("='Inputs 2.0 Conv. default values'!I85"));

            fmt = xls.GetCellVisibleFormatDef(86, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(86, 2, xls.AddFormat(fmt));
            xls.SetCellValue(86, 2, "Labor for procesing when the coffee trees are mature");

            fmt = xls.GetCellVisibleFormatDef(86, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(86, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(87, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(87, 2, xls.AddFormat(fmt));
            xls.SetCellValue(87, 2, "Pulp separation and fermentation (work time)");

            fmt = xls.GetCellVisibleFormatDef(87, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(87, 3, xls.AddFormat(fmt));
            xls.SetCellValue(87, 3, new TFormula("='Inputs 2.0 Conv. default values'!I87"));

            fmt = xls.GetCellVisibleFormatDef(88, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(88, 2, xls.AddFormat(fmt));
            xls.SetCellValue(88, 2, "Washing");

            fmt = xls.GetCellVisibleFormatDef(88, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(88, 3, xls.AddFormat(fmt));
            xls.SetCellValue(88, 3, new TFormula("='Inputs 2.0 Conv. default values'!I88"));

            fmt = xls.GetCellVisibleFormatDef(89, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(89, 2, xls.AddFormat(fmt));
            xls.SetCellValue(89, 2, "Drying");

            fmt = xls.GetCellVisibleFormatDef(89, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(89, 3, xls.AddFormat(fmt));
            xls.SetCellValue(89, 3, new TFormula("='Inputs 2.0 Conv. default values'!I89"));

            fmt = xls.GetCellVisibleFormatDef(90, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(90, 2, xls.AddFormat(fmt));
            xls.SetCellValue(90, 2, "Screening / shaking");

            fmt = xls.GetCellVisibleFormatDef(90, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(90, 3, xls.AddFormat(fmt));
            xls.SetCellValue(90, 3, new TFormula("='Inputs 2.0 Conv. default values'!I90"));

            fmt = xls.GetCellVisibleFormatDef(91, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(91, 2, xls.AddFormat(fmt));
            xls.SetCellValue(91, 2, "Selection / picking");

            fmt = xls.GetCellVisibleFormatDef(91, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(91, 3, xls.AddFormat(fmt));
            xls.SetCellValue(91, 3, new TFormula("='Inputs 2.0 Conv. default values'!I91"));

            fmt = xls.GetCellVisibleFormatDef(92, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(92, 2, xls.AddFormat(fmt));
            xls.SetCellValue(92, 2, "Storage");

            fmt = xls.GetCellVisibleFormatDef(92, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(92, 3, xls.AddFormat(fmt));
            xls.SetCellValue(92, 3, new TFormula("='Inputs 2.0 Conv. default values'!I92"));

            fmt = xls.GetCellVisibleFormatDef(93, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(93, 2, xls.AddFormat(fmt));
            xls.SetCellValue(93, 2, "Management of coffee wastewater");

            fmt = xls.GetCellVisibleFormatDef(93, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(93, 3, xls.AddFormat(fmt));
            xls.SetCellValue(93, 3, new TFormula("='Inputs 2.0 Conv. default values'!I93"));

            fmt = xls.GetCellVisibleFormatDef(94, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(94, 2, xls.AddFormat(fmt));
            xls.SetCellValue(94, 2, "Pulp management");

            fmt = xls.GetCellVisibleFormatDef(94, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(94, 3, xls.AddFormat(fmt));
            xls.SetCellValue(94, 3, new TFormula("='Inputs 2.0 Conv. default values'!I94"));

            fmt = xls.GetCellVisibleFormatDef(95, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(95, 2, xls.AddFormat(fmt));
            xls.SetCellValue(95, 2, "Other activities of the processing/beneficio:");

            fmt = xls.GetCellVisibleFormatDef(95, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(95, 3, xls.AddFormat(fmt));
            xls.SetCellValue(95, 3, new TFormula("='Inputs 2.0 Conv. default values'!I95"));

            fmt = xls.GetCellVisibleFormatDef(96, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(96, 2, xls.AddFormat(fmt));
            xls.SetCellValue(96, 2, "Labor for maintenance when the coffee trees are old");

            fmt = xls.GetCellVisibleFormatDef(96, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(96, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(96, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(97, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(97, 2, xls.AddFormat(fmt));
            xls.SetCellValue(97, 2, "Manual weeding");

            fmt = xls.GetCellVisibleFormatDef(97, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(97, 3, xls.AddFormat(fmt));
            xls.SetCellValue(97, 3, new TFormula("='Inputs 2.0 Conv. default values'!I97"));

            fmt = xls.GetCellVisibleFormatDef(98, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(98, 2, xls.AddFormat(fmt));
            xls.SetCellValue(98, 2, "Chemical weeding");

            fmt = xls.GetCellVisibleFormatDef(98, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(98, 3, xls.AddFormat(fmt));
            xls.SetCellValue(98, 3, new TFormula("='Inputs 2.0 Conv. default values'!I98"));

            fmt = xls.GetCellVisibleFormatDef(99, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(99, 2, xls.AddFormat(fmt));
            xls.SetCellValue(99, 2, "Organic fertilizers for maintenance");

            fmt = xls.GetCellVisibleFormatDef(99, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(99, 3, xls.AddFormat(fmt));
            xls.SetCellValue(99, 3, new TFormula("='Inputs 2.0 Conv. default values'!I99"));

            fmt = xls.GetCellVisibleFormatDef(100, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(100, 2, xls.AddFormat(fmt));
            xls.SetCellValue(100, 2, "Chemical fertilizers for maintenance");

            fmt = xls.GetCellVisibleFormatDef(100, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(100, 3, xls.AddFormat(fmt));
            xls.SetCellValue(100, 3, new TFormula("='Inputs 2.0 Conv. default values'!I100"));

            fmt = xls.GetCellVisibleFormatDef(101, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(101, 2, xls.AddFormat(fmt));
            xls.SetCellValue(101, 2, "Foliar spraying and rust control");

            fmt = xls.GetCellVisibleFormatDef(101, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(101, 3, xls.AddFormat(fmt));
            xls.SetCellValue(101, 3, new TFormula("='Inputs 2.0 Conv. default values'!I101"));

            fmt = xls.GetCellVisibleFormatDef(102, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(102, 2, xls.AddFormat(fmt));
            xls.SetCellValue(102, 2, "Hedgerows construction");

            fmt = xls.GetCellVisibleFormatDef(102, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(102, 3, xls.AddFormat(fmt));
            xls.SetCellValue(102, 3, new TFormula("='Inputs 2.0 Conv. default values'!I102"));

            fmt = xls.GetCellVisibleFormatDef(103, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(103, 2, xls.AddFormat(fmt));
            xls.SetCellValue(103, 2, "Shade tree pruning (maintenance) ");

            fmt = xls.GetCellVisibleFormatDef(103, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(103, 3, xls.AddFormat(fmt));
            xls.SetCellValue(103, 3, new TFormula("='Inputs 2.0 Conv. default values'!I103"));

            fmt = xls.GetCellVisibleFormatDef(104, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(104, 2, xls.AddFormat(fmt));
            xls.SetCellValue(104, 2, "Pest control (broca: fumigation, trap setting, etc.)");

            fmt = xls.GetCellVisibleFormatDef(104, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(104, 3, xls.AddFormat(fmt));
            xls.SetCellValue(104, 3, new TFormula("='Inputs 2.0 Conv. default values'!I104"));

            fmt = xls.GetCellVisibleFormatDef(105, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(105, 2, xls.AddFormat(fmt));
            xls.SetCellValue(105, 2, "Coffee growing management (pruning - agobio)");

            fmt = xls.GetCellVisibleFormatDef(105, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(105, 3, xls.AddFormat(fmt));
            xls.SetCellValue(105, 3, new TFormula("='Inputs 2.0 Conv. default values'!I105"));

            fmt = xls.GetCellVisibleFormatDef(106, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(106, 2, xls.AddFormat(fmt));
            xls.SetCellValue(106, 2, "Others:");

            fmt = xls.GetCellVisibleFormatDef(106, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(106, 3, xls.AddFormat(fmt));
            xls.SetCellValue(106, 3, new TFormula("='Inputs 2.0 Conv. default values'!I106"));

            fmt = xls.GetCellVisibleFormatDef(107, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(107, 2, xls.AddFormat(fmt));
            xls.SetCellValue(107, 2, "Labor for harvest when the coffee trees are old");

            fmt = xls.GetCellVisibleFormatDef(107, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(107, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(107, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(108, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(108, 2, xls.AddFormat(fmt));
            xls.SetCellValue(108, 2, "Tolal number of days invested to collect coffee                                  "
            + "                                                                                 "
            + "                                                                           Remember"
            + " total number of days is equal to:  Number of people*Days * Number of times per year");

            fmt = xls.GetCellVisibleFormatDef(108, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(108, 3, xls.AddFormat(fmt));
            xls.SetCellValue(108, 3, new TFormula("='Inputs 2.0 Conv. default values'!I108"));

            fmt = xls.GetCellVisibleFormatDef(109, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(109, 2, xls.AddFormat(fmt));
            xls.SetCellValue(109, 2, "Additional days invested in other activities related with the harvest ");

            fmt = xls.GetCellVisibleFormatDef(109, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(109, 3, xls.AddFormat(fmt));
            xls.SetCellValue(109, 3, new TFormula("='Inputs 2.0 Conv. default values'!I109"));

            fmt = xls.GetCellVisibleFormatDef(110, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(110, 2, xls.AddFormat(fmt));
            xls.SetCellValue(110, 2, "Labor for procesing when the coffee trees are old");

            fmt = xls.GetCellVisibleFormatDef(110, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(110, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(110, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(111, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(111, 2, xls.AddFormat(fmt));
            xls.SetCellValue(111, 2, "Pulp separation and fermentation (work time)");

            fmt = xls.GetCellVisibleFormatDef(111, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(111, 3, xls.AddFormat(fmt));
            xls.SetCellValue(111, 3, new TFormula("='Inputs 2.0 Conv. default values'!I111"));

            fmt = xls.GetCellVisibleFormatDef(112, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(112, 2, xls.AddFormat(fmt));
            xls.SetCellValue(112, 2, "Washing");

            fmt = xls.GetCellVisibleFormatDef(112, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(112, 3, xls.AddFormat(fmt));
            xls.SetCellValue(112, 3, new TFormula("='Inputs 2.0 Conv. default values'!I112"));

            fmt = xls.GetCellVisibleFormatDef(113, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(113, 2, xls.AddFormat(fmt));
            xls.SetCellValue(113, 2, "Drying");

            fmt = xls.GetCellVisibleFormatDef(113, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(113, 3, xls.AddFormat(fmt));
            xls.SetCellValue(113, 3, new TFormula("='Inputs 2.0 Conv. default values'!I113"));

            fmt = xls.GetCellVisibleFormatDef(114, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(114, 2, xls.AddFormat(fmt));
            xls.SetCellValue(114, 2, "Screening / shaking");

            fmt = xls.GetCellVisibleFormatDef(114, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(114, 3, xls.AddFormat(fmt));
            xls.SetCellValue(114, 3, new TFormula("='Inputs 2.0 Conv. default values'!I114"));

            fmt = xls.GetCellVisibleFormatDef(115, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(115, 2, xls.AddFormat(fmt));
            xls.SetCellValue(115, 2, "Selection / picking");

            fmt = xls.GetCellVisibleFormatDef(115, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(115, 3, xls.AddFormat(fmt));
            xls.SetCellValue(115, 3, new TFormula("='Inputs 2.0 Conv. default values'!I115"));

            fmt = xls.GetCellVisibleFormatDef(116, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(116, 2, xls.AddFormat(fmt));
            xls.SetCellValue(116, 2, "Storage");

            fmt = xls.GetCellVisibleFormatDef(116, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(116, 3, xls.AddFormat(fmt));
            xls.SetCellValue(116, 3, new TFormula("='Inputs 2.0 Conv. default values'!I116"));

            fmt = xls.GetCellVisibleFormatDef(117, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(117, 2, xls.AddFormat(fmt));
            xls.SetCellValue(117, 2, "Management of coffee wastewater");

            fmt = xls.GetCellVisibleFormatDef(117, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(117, 3, xls.AddFormat(fmt));
            xls.SetCellValue(117, 3, new TFormula("='Inputs 2.0 Conv. default values'!I117"));

            fmt = xls.GetCellVisibleFormatDef(118, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(118, 2, xls.AddFormat(fmt));
            xls.SetCellValue(118, 2, "Pulp management");

            fmt = xls.GetCellVisibleFormatDef(118, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(118, 3, xls.AddFormat(fmt));
            xls.SetCellValue(118, 3, new TFormula("='Inputs 2.0 Conv. default values'!I118"));

            fmt = xls.GetCellVisibleFormatDef(119, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(119, 2, xls.AddFormat(fmt));
            xls.SetCellValue(119, 2, "Other activities of the processing/beneficio:");

            fmt = xls.GetCellVisibleFormatDef(119, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(119, 3, xls.AddFormat(fmt));
            xls.SetCellValue(119, 3, new TFormula("='Inputs 2.0 Conv. default values'!I119"));

            fmt = xls.GetCellVisibleFormatDef(120, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 2, xls.AddFormat(fmt));
            xls.SetCellValue(120, 2, "Additional Income and remunertion");

            fmt = xls.GetCellVisibleFormatDef(120, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(120, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(120, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(121, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(121, 2, xls.AddFormat(fmt));
            xls.SetCellValue(121, 2, "Additional remuneration and indirect income");

            fmt = xls.GetCellVisibleFormatDef(121, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(121, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(122, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(122, 2, xls.AddFormat(fmt));
            xls.SetCellValue(122, 2, new TFormula("=+\"In addition to the daily payment or jornal, do you feed your workers? What is"
            + " the value estimated of this food in \"&'Gral Conf. Summary'!$H$33&\"?\""));

            fmt = xls.GetCellVisibleFormatDef(122, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(122, 3, xls.AddFormat(fmt));
            xls.SetCellValue(122, 3, new TFormula("='Inputs 2.0 Conv. default values'!I122"));

            fmt = xls.GetCellVisibleFormatDef(123, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(123, 2, xls.AddFormat(fmt));
            xls.SetCellValue(123, 2, "Value of additional transfers from the cooperative in money or goods (fertilizers"
            + " etc)");

            fmt = xls.GetCellVisibleFormatDef(123, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(123, 3, xls.AddFormat(fmt));
            xls.SetCellValue(123, 3, new TFormula("='Inputs 2.0 Conv. default values'!I123"));

            fmt = xls.GetCellVisibleFormatDef(124, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(124, 2, xls.AddFormat(fmt));
            xls.SetCellValue(124, 2, "How many days of trainning received in the cooperative per year?");

            fmt = xls.GetCellVisibleFormatDef(124, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(124, 3, xls.AddFormat(fmt));
            xls.SetCellValue(124, 3, new TFormula("='Inputs 2.0 Conv. default values'!I124"));

            fmt = xls.GetCellVisibleFormatDef(125, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(125, 2, xls.AddFormat(fmt));
            xls.SetCellValue(125, 2, "Credit");

            fmt = xls.GetCellVisibleFormatDef(125, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(125, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(126, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(126, 2, xls.AddFormat(fmt));
            xls.SetCellValue(126, 2, new TFormula("=+\"Did you receive any credit from the cooperative to invest in your farm or coffee"
            + " production activities? What was the amount of this credit in \"&'Gral Conf. Summary'!$H$33&\""
            + " ?\""));

            fmt = xls.GetCellVisibleFormatDef(126, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(126, 3, xls.AddFormat(fmt));
            xls.SetCellValue(126, 3, new TFormula("='Inputs 2.0 Conv. default values'!I126"));

            fmt = xls.GetCellVisibleFormatDef(127, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(127, 2, xls.AddFormat(fmt));
            xls.SetCellValue(127, 2, "Time of the credit in years");

            fmt = xls.GetCellVisibleFormatDef(127, 3);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(127, 3, xls.AddFormat(fmt));
            xls.SetCellValue(127, 3, new TFormula("='Inputs 2.0 Conv. default values'!I127"));

            fmt = xls.GetCellVisibleFormatDef(128, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(128, 2, xls.AddFormat(fmt));
            xls.SetCellValue(128, 2, "What is the annual interest rate of this loan?");

            fmt = xls.GetCellVisibleFormatDef(128, 3);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(128, 3, xls.AddFormat(fmt));
            xls.SetCellValue(128, 3, new TFormula("='Inputs 2.0 Conv. default values'!I128"));

            fmt = xls.GetCellVisibleFormatDef(129, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(129, 2, xls.AddFormat(fmt));
            xls.SetCellValue(129, 2, new TFormula("=+\"Did you receive any credit from an agent different than the  cooperative to invest"
            + " in your farm or coffee production activities? What was the amount of this credit"
            + " in \"&'Gral Conf. Summary'!$H$33&\" ?\""));

            fmt = xls.GetCellVisibleFormatDef(129, 3);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(129, 3, xls.AddFormat(fmt));
            xls.SetCellValue(129, 3, new TFormula("='Inputs 2.0 Conv. default values'!I129"));

            fmt = xls.GetCellVisibleFormatDef(130, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(130, 2, xls.AddFormat(fmt));
            xls.SetCellValue(130, 2, "Time of the credit in years");

            fmt = xls.GetCellVisibleFormatDef(130, 3);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(130, 3, xls.AddFormat(fmt));
            xls.SetCellValue(130, 3, new TFormula("='Inputs 2.0 Conv. default values'!I130"));

            fmt = xls.GetCellVisibleFormatDef(131, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(131, 2, xls.AddFormat(fmt));
            xls.SetCellValue(131, 2, "What is the annual interest rate of this loan?");

            fmt = xls.GetCellVisibleFormatDef(131, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(131, 3, xls.AddFormat(fmt));
            xls.SetCellValue(131, 3, new TFormula("='Inputs 2.0 Conv. default values'!I131"));

            fmt = xls.GetCellVisibleFormatDef(132, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(132, 2, xls.AddFormat(fmt));
            xls.SetCellValue(132, 2, "Cost of materials and inputs");

            fmt = xls.GetCellVisibleFormatDef(132, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(132, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(132, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(133, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(133, 2, xls.AddFormat(fmt));
            xls.SetCellValue(133, 2, new TFormula("=+\"Please, describe how much do you spent in  \"&'Gral Conf. Summary'!$H$33&\" in"
            + " the following inputs to establish and maintain ONE \"&'Gral Conf. Summary'!$I$23&\""
            + " of coffee\""));

            fmt = xls.GetCellVisibleFormatDef(133, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(133, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(133, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(133, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(134, 2, xls.AddFormat(fmt));
            xls.SetCellValue(134, 2, "Materials for germinator");

            fmt = xls.GetCellVisibleFormatDef(134, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(134, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(134, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(134, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(135, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(135, 2, xls.AddFormat(fmt));
            xls.SetCellValue(135, 2, "Seeds");

            fmt = xls.GetCellVisibleFormatDef(135, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(135, 3, xls.AddFormat(fmt));
            xls.SetCellValue(135, 3, new TFormula("='Inputs 2.0 Conv. default values'!I135"));

            fmt = xls.GetCellVisibleFormatDef(136, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(136, 2, xls.AddFormat(fmt));
            xls.SetCellValue(136, 2, "Germinator / Seedbed");

            fmt = xls.GetCellVisibleFormatDef(136, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(136, 3, xls.AddFormat(fmt));
            xls.SetCellValue(136, 3, new TFormula("='Inputs 2.0 Conv. default values'!I136"));

            fmt = xls.GetCellVisibleFormatDef(137, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(137, 2, xls.AddFormat(fmt));
            xls.SetCellValue(137, 2, "Sand substrate");

            fmt = xls.GetCellVisibleFormatDef(137, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(137, 3, xls.AddFormat(fmt));
            xls.SetCellValue(137, 3, new TFormula("='Inputs 2.0 Conv. default values'!I137"));

            fmt = xls.GetCellVisibleFormatDef(138, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(138, 2, xls.AddFormat(fmt));
            xls.SetCellValue(138, 2, "Calcium sulfide");

            fmt = xls.GetCellVisibleFormatDef(138, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(138, 3, xls.AddFormat(fmt));
            xls.SetCellValue(138, 3, new TFormula("='Inputs 2.0 Conv. default values'!I138"));

            fmt = xls.GetCellVisibleFormatDef(139, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(139, 2, xls.AddFormat(fmt));
            xls.SetCellValue(139, 2, "Lime");

            fmt = xls.GetCellVisibleFormatDef(139, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(139, 3, xls.AddFormat(fmt));
            xls.SetCellValue(139, 3, new TFormula("='Inputs 2.0 Conv. default values'!I139"));

            fmt = xls.GetCellVisibleFormatDef(140, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(140, 2, xls.AddFormat(fmt));
            xls.SetCellValue(140, 2, "Plastic");

            fmt = xls.GetCellVisibleFormatDef(140, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(140, 3, xls.AddFormat(fmt));
            xls.SetCellValue(140, 3, new TFormula("='Inputs 2.0 Conv. default values'!I140"));

            fmt = xls.GetCellVisibleFormatDef(141, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(141, 2, xls.AddFormat(fmt));
            xls.SetCellValue(141, 2, "Other material(s) for the germinator:");

            fmt = xls.GetCellVisibleFormatDef(141, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(141, 3, xls.AddFormat(fmt));
            xls.SetCellValue(141, 3, new TFormula("='Inputs 2.0 Conv. default values'!I141"));

            fmt = xls.GetCellVisibleFormatDef(142, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(142, 2, xls.AddFormat(fmt));
            xls.SetCellValue(142, 2, "Materials for nursery");

            fmt = xls.GetCellVisibleFormatDef(142, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(142, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(142, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            xls.SetCellFormat(142, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(143, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(143, 2, xls.AddFormat(fmt));
            xls.SetCellValue(143, 2, "Organic fertilizer (For example: Bocachi, others)");

            fmt = xls.GetCellVisibleFormatDef(143, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(143, 3, xls.AddFormat(fmt));
            xls.SetCellValue(143, 3, new TFormula("='Inputs 2.0 Conv. default values'!I143"));

            fmt = xls.GetCellVisibleFormatDef(144, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(144, 2, xls.AddFormat(fmt));
            xls.SetCellValue(144, 2, "Plastic bags");

            fmt = xls.GetCellVisibleFormatDef(144, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(144, 3, xls.AddFormat(fmt));
            xls.SetCellValue(144, 3, new TFormula("='Inputs 2.0 Conv. default values'!I144"));

            fmt = xls.GetCellVisibleFormatDef(145, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(145, 2, xls.AddFormat(fmt));
            xls.SetCellValue(145, 2, "Netting");

            fmt = xls.GetCellVisibleFormatDef(145, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(145, 3, xls.AddFormat(fmt));
            xls.SetCellValue(145, 3, new TFormula("='Inputs 2.0 Conv. default values'!I145"));

            fmt = xls.GetCellVisibleFormatDef(146, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(146, 2, xls.AddFormat(fmt));
            xls.SetCellValue(146, 2, "Studs");

            fmt = xls.GetCellVisibleFormatDef(146, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(146, 3, xls.AddFormat(fmt));
            xls.SetCellValue(146, 3, new TFormula("='Inputs 2.0 Conv. default values'!I146"));

            fmt = xls.GetCellVisibleFormatDef(147, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(147, 2, xls.AddFormat(fmt));
            xls.SetCellValue(147, 2, "Wire");

            fmt = xls.GetCellVisibleFormatDef(147, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(147, 3, xls.AddFormat(fmt));
            xls.SetCellValue(147, 3, new TFormula("='Inputs 2.0 Conv. default values'!I147"));

            fmt = xls.GetCellVisibleFormatDef(148, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(148, 2, xls.AddFormat(fmt));
            xls.SetCellValue(148, 2, "Ciclonics netting");

            fmt = xls.GetCellVisibleFormatDef(148, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(148, 3, xls.AddFormat(fmt));
            xls.SetCellValue(148, 3, new TFormula("='Inputs 2.0 Conv. default values'!I148"));

            fmt = xls.GetCellVisibleFormatDef(149, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(149, 2, xls.AddFormat(fmt));
            xls.SetCellValue(149, 2, "Staples");

            fmt = xls.GetCellVisibleFormatDef(149, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(149, 3, xls.AddFormat(fmt));
            xls.SetCellValue(149, 3, new TFormula("='Inputs 2.0 Conv. default values'!I149"));

            fmt = xls.GetCellVisibleFormatDef(150, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(150, 2, xls.AddFormat(fmt));
            xls.SetCellValue(150, 2, "Soil");

            fmt = xls.GetCellVisibleFormatDef(150, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(150, 3, xls.AddFormat(fmt));
            xls.SetCellValue(150, 3, new TFormula("='Inputs 2.0 Conv. default values'!I150"));

            fmt = xls.GetCellVisibleFormatDef(151, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(151, 2, xls.AddFormat(fmt));
            xls.SetCellValue(151, 2, "Liquid biofertilizers ");

            fmt = xls.GetCellVisibleFormatDef(151, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(151, 3, xls.AddFormat(fmt));
            xls.SetCellValue(151, 3, new TFormula("='Inputs 2.0 Conv. default values'!I151"));

            fmt = xls.GetCellVisibleFormatDef(152, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(152, 2, xls.AddFormat(fmt));
            xls.SetCellValue(152, 2, "Agrochemicals (for the nursery)");

            fmt = xls.GetCellVisibleFormatDef(152, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(152, 3, xls.AddFormat(fmt));
            xls.SetCellValue(152, 3, new TFormula("='Inputs 2.0 Conv. default values'!I152"));

            fmt = xls.GetCellVisibleFormatDef(153, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(153, 2, xls.AddFormat(fmt));
            xls.SetCellValue(153, 2, "Fungicide (for the nursery)");

            fmt = xls.GetCellVisibleFormatDef(153, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(153, 3, xls.AddFormat(fmt));
            xls.SetCellValue(153, 3, new TFormula("='Inputs 2.0 Conv. default values'!I153"));

            fmt = xls.GetCellVisibleFormatDef(154, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(154, 2, xls.AddFormat(fmt));
            xls.SetCellValue(154, 2, "Fosforic rock");

            fmt = xls.GetCellVisibleFormatDef(154, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(154, 3, xls.AddFormat(fmt));
            xls.SetCellValue(154, 3, new TFormula("='Inputs 2.0 Conv. default values'!I154"));

            fmt = xls.GetCellVisibleFormatDef(155, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(155, 2, xls.AddFormat(fmt));
            xls.SetCellValue(155, 2, "Other material(s) for the nursery:");

            fmt = xls.GetCellVisibleFormatDef(155, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(155, 3, xls.AddFormat(fmt));
            xls.SetCellValue(155, 3, new TFormula("='Inputs 2.0 Conv. default values'!I155"));

            fmt = xls.GetCellVisibleFormatDef(156, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(156, 2, xls.AddFormat(fmt));
            xls.SetCellValue(156, 2, "Fertilizers during the year of land prepararion and planting");

            fmt = xls.GetCellVisibleFormatDef(156, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(156, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(157, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(157, 2, xls.AddFormat(fmt));
            xls.SetCellValue(157, 2, "Organic fertilizers for the holes");

            fmt = xls.GetCellVisibleFormatDef(157, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(157, 3, xls.AddFormat(fmt));
            xls.SetCellValue(157, 3, new TFormula("='Inputs 2.0 Conv. default values'!I157"));

            fmt = xls.GetCellVisibleFormatDef(158, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(158, 2, xls.AddFormat(fmt));
            xls.SetCellValue(158, 2, "Chemical fertilizers for the holes");

            fmt = xls.GetCellVisibleFormatDef(158, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(158, 3, xls.AddFormat(fmt));
            xls.SetCellValue(158, 3, new TFormula("='Inputs 2.0 Conv. default values'!I158"));

            fmt = xls.GetCellVisibleFormatDef(159, 2);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(159, 2, xls.AddFormat(fmt));
            xls.SetCellValue(159, 2, "Fertilizers and foliadge during the year of vegetatitive growth");

            fmt = xls.GetCellVisibleFormatDef(159, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(159, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(160, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(160, 2, xls.AddFormat(fmt));
            xls.SetCellValue(160, 2, "Organic fertilizers");

            fmt = xls.GetCellVisibleFormatDef(160, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(160, 3, xls.AddFormat(fmt));
            xls.SetCellValue(160, 3, new TFormula("='Inputs 2.0 Conv. default values'!I160"));

            fmt = xls.GetCellVisibleFormatDef(161, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(161, 2, xls.AddFormat(fmt));
            xls.SetCellValue(161, 2, "Chemical fertilizers");

            fmt = xls.GetCellVisibleFormatDef(161, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(161, 3, xls.AddFormat(fmt));
            xls.SetCellValue(161, 3, new TFormula("='Inputs 2.0 Conv. default values'!I161"));

            fmt = xls.GetCellVisibleFormatDef(162, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(162, 2, xls.AddFormat(fmt));
            xls.SetCellValue(162, 2, "Fertilizers and foliadge during mantainance");

            fmt = xls.GetCellVisibleFormatDef(162, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(162, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(163, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(163, 2, xls.AddFormat(fmt));
            xls.SetCellValue(163, 2, "Other fertilizer for mantainace  not specified in basic inputs");

            fmt = xls.GetCellVisibleFormatDef(163, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(163, 3, xls.AddFormat(fmt));
            xls.SetCellValue(163, 3, new TFormula("='Inputs 2.0 Conv. default values'!I163"));

            fmt = xls.GetCellVisibleFormatDef(164, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(164, 2, xls.AddFormat(fmt));
            xls.SetCellValue(164, 2, "Organic foliar spraying");

            fmt = xls.GetCellVisibleFormatDef(164, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(164, 3, xls.AddFormat(fmt));
            xls.SetCellValue(164, 3, new TFormula("='Inputs 2.0 Conv. default values'!I164"));

            fmt = xls.GetCellVisibleFormatDef(165, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(165, 2, xls.AddFormat(fmt));
            xls.SetCellValue(165, 2, "Chemical foliar spraying");

            fmt = xls.GetCellVisibleFormatDef(165, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(165, 3, xls.AddFormat(fmt));
            xls.SetCellValue(165, 3, new TFormula("='Inputs 2.0 Conv. default values'!I165"));

            fmt = xls.GetCellVisibleFormatDef(166, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(166, 2, xls.AddFormat(fmt));
            xls.SetCellValue(166, 2, "Gas / fuel");

            fmt = xls.GetCellVisibleFormatDef(166, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(166, 3, xls.AddFormat(fmt));
            xls.SetCellValue(166, 3, new TFormula("='Inputs 2.0 Conv. default values'!I166"));

            fmt = xls.GetCellVisibleFormatDef(167, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(167, 2, xls.AddFormat(fmt));
            xls.SetCellValue(167, 2, "Other inputs for mantainance");

            fmt = xls.GetCellVisibleFormatDef(167, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(167, 3, xls.AddFormat(fmt));
            xls.SetCellValue(167, 3, new TFormula("='Inputs 2.0 Conv. default values'!I167"));

            fmt = xls.GetCellVisibleFormatDef(168, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(168, 2, xls.AddFormat(fmt));
            xls.SetCellValue(168, 2, "Equipment and Reusable material");

            fmt = xls.GetCellVisibleFormatDef(168, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(168, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(168, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(169, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(169, 2, xls.AddFormat(fmt));
            xls.SetCellValue(169, 2, new TFormula("=+\"Please, describe how much do you spent in  \"&'Gral Conf. Summary'!$H$33&\" in"
            + " the following equipment and reusable materials to establish and maintain ONE \"&'Gral"
            + " Conf. Summary'!$I$23&\" of coffee. In addition, provide the lifespam of these items"
            + " in years.\""));

            fmt = xls.GetCellVisibleFormatDef(169, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(169, 3, xls.AddFormat(fmt));
            xls.SetCellValue(169, 3, new TFormula("='Inputs 2.0 Conv. default values'!I169"));

            fmt = xls.GetCellVisibleFormatDef(169, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(169, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(170, 2, xls.AddFormat(fmt));
            xls.SetCellValue(170, 2, "General equipment");

            fmt = xls.GetCellVisibleFormatDef(170, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(170, 10);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(170, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(171, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(171, 2, xls.AddFormat(fmt));
            xls.SetCellValue(171, 2, "Manual sprayer or fumigation backpack");

            fmt = xls.GetCellVisibleFormatDef(171, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(171, 3, xls.AddFormat(fmt));
            xls.SetCellValue(171, 3, new TFormula("='Inputs 2.0 Conv. default values'!I171"));

            fmt = xls.GetCellVisibleFormatDef(172, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(172, 2, xls.AddFormat(fmt));
            xls.SetCellValue(172, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(172, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(172, 3, xls.AddFormat(fmt));
            xls.SetCellValue(172, 3, new TFormula("='Inputs 2.0 Conv. default values'!I172"));

            fmt = xls.GetCellVisibleFormatDef(173, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(173, 2, xls.AddFormat(fmt));
            xls.SetCellValue(173, 2, "Machetes (For example: cuta /cane machetes or others)");

            fmt = xls.GetCellVisibleFormatDef(173, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(173, 3, xls.AddFormat(fmt));
            xls.SetCellValue(173, 3, new TFormula("='Inputs 2.0 Conv. default values'!I173"));

            fmt = xls.GetCellVisibleFormatDef(174, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(174, 2, xls.AddFormat(fmt));
            xls.SetCellValue(174, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(174, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(174, 3, xls.AddFormat(fmt));
            xls.SetCellValue(174, 3, new TFormula("='Inputs 2.0 Conv. default values'!I174"));

            fmt = xls.GetCellVisibleFormatDef(175, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(175, 2, xls.AddFormat(fmt));
            xls.SetCellValue(175, 2, "Shovel");

            fmt = xls.GetCellVisibleFormatDef(175, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(175, 3, xls.AddFormat(fmt));
            xls.SetCellValue(175, 3, new TFormula("='Inputs 2.0 Conv. default values'!I175"));

            fmt = xls.GetCellVisibleFormatDef(176, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(176, 2, xls.AddFormat(fmt));
            xls.SetCellValue(176, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(176, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(176, 3, xls.AddFormat(fmt));
            xls.SetCellValue(176, 3, new TFormula("='Inputs 2.0 Conv. default values'!I176"));

            fmt = xls.GetCellVisibleFormatDef(177, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(177, 2, xls.AddFormat(fmt));
            xls.SetCellValue(177, 2, "Hoe");

            fmt = xls.GetCellVisibleFormatDef(177, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(177, 3, xls.AddFormat(fmt));
            xls.SetCellValue(177, 3, new TFormula("='Inputs 2.0 Conv. default values'!I177"));

            fmt = xls.GetCellVisibleFormatDef(178, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(178, 2, xls.AddFormat(fmt));
            xls.SetCellValue(178, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(178, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(178, 3, xls.AddFormat(fmt));
            xls.SetCellValue(178, 3, new TFormula("='Inputs 2.0 Conv. default values'!I178"));

            fmt = xls.GetCellVisibleFormatDef(179, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(179, 2, xls.AddFormat(fmt));
            xls.SetCellValue(179, 2, "Wheelbarrow");

            fmt = xls.GetCellVisibleFormatDef(179, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(179, 3, xls.AddFormat(fmt));
            xls.SetCellValue(179, 3, new TFormula("='Inputs 2.0 Conv. default values'!I179"));

            fmt = xls.GetCellVisibleFormatDef(180, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(180, 2, xls.AddFormat(fmt));
            xls.SetCellValue(180, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(180, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(180, 3, xls.AddFormat(fmt));
            xls.SetCellValue(180, 3, new TFormula("='Inputs 2.0 Conv. default values'!I180"));

            fmt = xls.GetCellVisibleFormatDef(181, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(181, 2, xls.AddFormat(fmt));
            xls.SetCellValue(181, 2, "Lime / file");

            fmt = xls.GetCellVisibleFormatDef(181, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(181, 3, xls.AddFormat(fmt));
            xls.SetCellValue(181, 3, new TFormula("='Inputs 2.0 Conv. default values'!I181"));

            fmt = xls.GetCellVisibleFormatDef(182, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(182, 2, xls.AddFormat(fmt));
            xls.SetCellValue(182, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(182, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(182, 3, xls.AddFormat(fmt));
            xls.SetCellValue(182, 3, new TFormula("='Inputs 2.0 Conv. default values'!I182"));

            fmt = xls.GetCellVisibleFormatDef(183, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(183, 2, xls.AddFormat(fmt));
            xls.SetCellValue(183, 2, "Auger / drilling device");

            fmt = xls.GetCellVisibleFormatDef(183, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(183, 3, xls.AddFormat(fmt));
            xls.SetCellValue(183, 3, new TFormula("='Inputs 2.0 Conv. default values'!I183"));

            fmt = xls.GetCellVisibleFormatDef(184, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(184, 2, xls.AddFormat(fmt));
            xls.SetCellValue(184, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(184, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(184, 3, xls.AddFormat(fmt));
            xls.SetCellValue(184, 3, new TFormula("='Inputs 2.0 Conv. default values'!I184"));

            fmt = xls.GetCellVisibleFormatDef(185, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(185, 2, xls.AddFormat(fmt));
            xls.SetCellValue(185, 2, "Metal bar / Barretón");

            fmt = xls.GetCellVisibleFormatDef(185, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(185, 3, xls.AddFormat(fmt));
            xls.SetCellValue(185, 3, new TFormula("='Inputs 2.0 Conv. default values'!I185"));

            fmt = xls.GetCellVisibleFormatDef(186, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(186, 2, xls.AddFormat(fmt));
            xls.SetCellValue(186, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(186, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(186, 3, xls.AddFormat(fmt));
            xls.SetCellValue(186, 3, new TFormula("='Inputs 2.0 Conv. default values'!I186"));

            fmt = xls.GetCellVisibleFormatDef(187, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(187, 2, xls.AddFormat(fmt));
            xls.SetCellValue(187, 2, "Hose");

            fmt = xls.GetCellVisibleFormatDef(187, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(187, 3, xls.AddFormat(fmt));
            xls.SetCellValue(187, 3, new TFormula("='Inputs 2.0 Conv. default values'!I187"));

            fmt = xls.GetCellVisibleFormatDef(188, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(188, 2, xls.AddFormat(fmt));
            xls.SetCellValue(188, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(188, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(188, 3, xls.AddFormat(fmt));
            xls.SetCellValue(188, 3, new TFormula("='Inputs 2.0 Conv. default values'!I188"));

            fmt = xls.GetCellVisibleFormatDef(189, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(189, 2, xls.AddFormat(fmt));
            xls.SetCellValue(189, 2, "Irrigation system (sprinklers)");

            fmt = xls.GetCellVisibleFormatDef(189, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(189, 3, xls.AddFormat(fmt));
            xls.SetCellValue(189, 3, new TFormula("='Inputs 2.0 Conv. default values'!I189"));

            fmt = xls.GetCellVisibleFormatDef(190, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(190, 2, xls.AddFormat(fmt));
            xls.SetCellValue(190, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(190, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(190, 3, xls.AddFormat(fmt));
            xls.SetCellValue(190, 3, new TFormula("='Inputs 2.0 Conv. default values'!I190"));

            fmt = xls.GetCellVisibleFormatDef(191, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(191, 2, xls.AddFormat(fmt));
            xls.SetCellValue(191, 2, "Chainsaw");

            fmt = xls.GetCellVisibleFormatDef(191, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(191, 3, xls.AddFormat(fmt));
            xls.SetCellValue(191, 3, new TFormula("='Inputs 2.0 Conv. default values'!I191"));

            fmt = xls.GetCellVisibleFormatDef(192, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(192, 2, xls.AddFormat(fmt));
            xls.SetCellValue(192, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(192, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(192, 3, xls.AddFormat(fmt));
            xls.SetCellValue(192, 3, new TFormula("='Inputs 2.0 Conv. default values'!I192"));

            fmt = xls.GetCellVisibleFormatDef(193, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(193, 2, xls.AddFormat(fmt));
            xls.SetCellValue(193, 2, "Handsaw");

            fmt = xls.GetCellVisibleFormatDef(193, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(193, 3, xls.AddFormat(fmt));
            xls.SetCellValue(193, 3, new TFormula("='Inputs 2.0 Conv. default values'!I193"));

            fmt = xls.GetCellVisibleFormatDef(194, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(194, 2, xls.AddFormat(fmt));
            xls.SetCellValue(194, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(194, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(194, 3, xls.AddFormat(fmt));
            xls.SetCellValue(194, 3, new TFormula("='Inputs 2.0 Conv. default values'!I194"));

            fmt = xls.GetCellVisibleFormatDef(195, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(195, 2, xls.AddFormat(fmt));
            xls.SetCellValue(195, 2, "Motor pump");

            fmt = xls.GetCellVisibleFormatDef(195, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(195, 3, xls.AddFormat(fmt));
            xls.SetCellValue(195, 3, new TFormula("='Inputs 2.0 Conv. default values'!I195"));

            fmt = xls.GetCellVisibleFormatDef(196, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(196, 2, xls.AddFormat(fmt));
            xls.SetCellValue(196, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(196, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(196, 3, xls.AddFormat(fmt));
            xls.SetCellValue(196, 3, new TFormula("='Inputs 2.0 Conv. default values'!I196"));

            fmt = xls.GetCellVisibleFormatDef(197, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(197, 2, xls.AddFormat(fmt));
            xls.SetCellValue(197, 2, "Prunning scissors");

            fmt = xls.GetCellVisibleFormatDef(197, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(197, 3, xls.AddFormat(fmt));
            xls.SetCellValue(197, 3, new TFormula("='Inputs 2.0 Conv. default values'!I197"));

            fmt = xls.GetCellVisibleFormatDef(198, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(198, 2, xls.AddFormat(fmt));
            xls.SetCellValue(198, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(198, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(198, 3, xls.AddFormat(fmt));
            xls.SetCellValue(198, 3, new TFormula("='Inputs 2.0 Conv. default values'!I198"));

            fmt = xls.GetCellVisibleFormatDef(199, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(199, 2, xls.AddFormat(fmt));
            xls.SetCellValue(199, 2, "Axe");

            fmt = xls.GetCellVisibleFormatDef(199, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(199, 3, xls.AddFormat(fmt));
            xls.SetCellValue(199, 3, new TFormula("='Inputs 2.0 Conv. default values'!I199"));

            fmt = xls.GetCellVisibleFormatDef(200, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(200, 2, xls.AddFormat(fmt));
            xls.SetCellValue(200, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(200, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(200, 3, xls.AddFormat(fmt));
            xls.SetCellValue(200, 3, new TFormula("='Inputs 2.0 Conv. default values'!I200"));

            fmt = xls.GetCellVisibleFormatDef(201, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xD8, 0xD8, 0xD8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(201, 2, xls.AddFormat(fmt));
            xls.SetCellValue(201, 2, "Equipment and materials for the harvest and other activities");

            fmt = xls.GetCellVisibleFormatDef(201, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 4);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 5);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 6);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 7);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 8);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(201, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(201, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(202, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(202, 2, xls.AddFormat(fmt));
            xls.SetCellValue(202, 2, "Scale or balance");

            fmt = xls.GetCellVisibleFormatDef(202, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(202, 3, xls.AddFormat(fmt));
            xls.SetCellValue(202, 3, new TFormula("='Inputs 2.0 Conv. default values'!I202"));

            fmt = xls.GetCellVisibleFormatDef(203, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(203, 2, xls.AddFormat(fmt));
            xls.SetCellValue(203, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(203, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(203, 3, xls.AddFormat(fmt));
            xls.SetCellValue(203, 3, new TFormula("='Inputs 2.0 Conv. default values'!I203"));

            fmt = xls.GetCellVisibleFormatDef(204, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(204, 2, xls.AddFormat(fmt));
            xls.SetCellValue(204, 2, "Vehicle or car for labor");

            fmt = xls.GetCellVisibleFormatDef(204, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(204, 3, xls.AddFormat(fmt));
            xls.SetCellValue(204, 3, new TFormula("='Inputs 2.0 Conv. default values'!I204"));

            fmt = xls.GetCellVisibleFormatDef(205, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(205, 2, xls.AddFormat(fmt));
            xls.SetCellValue(205, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(205, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(205, 3, xls.AddFormat(fmt));
            xls.SetCellValue(205, 3, new TFormula("='Inputs 2.0 Conv. default values'!I205"));

            fmt = xls.GetCellVisibleFormatDef(206, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(206, 2, xls.AddFormat(fmt));
            xls.SetCellValue(206, 2, "Work animal");

            fmt = xls.GetCellVisibleFormatDef(206, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(206, 3, xls.AddFormat(fmt));
            xls.SetCellValue(206, 3, new TFormula("='Inputs 2.0 Conv. default values'!I206"));

            fmt = xls.GetCellVisibleFormatDef(207, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(207, 2, xls.AddFormat(fmt));
            xls.SetCellValue(207, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(207, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(207, 3, xls.AddFormat(fmt));
            xls.SetCellValue(207, 3, new TFormula("='Inputs 2.0 Conv. default values'!I207"));

            fmt = xls.GetCellVisibleFormatDef(208, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(208, 2, xls.AddFormat(fmt));
            xls.SetCellValue(208, 2, "Motorcycle for labor");

            fmt = xls.GetCellVisibleFormatDef(208, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(208, 3, xls.AddFormat(fmt));
            xls.SetCellValue(208, 3, new TFormula("='Inputs 2.0 Conv. default values'!I208"));

            fmt = xls.GetCellVisibleFormatDef(209, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(209, 2, xls.AddFormat(fmt));
            xls.SetCellValue(209, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(209, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(209, 3, xls.AddFormat(fmt));
            xls.SetCellValue(209, 3, new TFormula("='Inputs 2.0 Conv. default values'!I209"));

            fmt = xls.GetCellVisibleFormatDef(210, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(210, 2, xls.AddFormat(fmt));
            xls.SetCellValue(210, 2, "Bags for collecting / sacks");

            fmt = xls.GetCellVisibleFormatDef(210, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(210, 3, xls.AddFormat(fmt));
            xls.SetCellValue(210, 3, new TFormula("='Inputs 2.0 Conv. default values'!I210"));

            fmt = xls.GetCellVisibleFormatDef(211, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(211, 2, xls.AddFormat(fmt));
            xls.SetCellValue(211, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(211, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(211, 3, xls.AddFormat(fmt));
            xls.SetCellValue(211, 3, new TFormula("='Inputs 2.0 Conv. default values'!I211"));

            fmt = xls.GetCellVisibleFormatDef(212, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(212, 2, xls.AddFormat(fmt));
            xls.SetCellValue(212, 2, "Sack for dry parchment ");

            fmt = xls.GetCellVisibleFormatDef(212, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(212, 3, xls.AddFormat(fmt));
            xls.SetCellValue(212, 3, new TFormula("='Inputs 2.0 Conv. default values'!I212"));

            fmt = xls.GetCellVisibleFormatDef(213, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(213, 2, xls.AddFormat(fmt));
            xls.SetCellValue(213, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(213, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(213, 3, xls.AddFormat(fmt));
            xls.SetCellValue(213, 3, new TFormula("='Inputs 2.0 Conv. default values'!I213"));

            fmt = xls.GetCellVisibleFormatDef(214, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(214, 2, xls.AddFormat(fmt));
            xls.SetCellValue(214, 2, "Straw / Raffia");

            fmt = xls.GetCellVisibleFormatDef(214, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(214, 3, xls.AddFormat(fmt));
            xls.SetCellValue(214, 3, new TFormula("='Inputs 2.0 Conv. default values'!I214"));

            fmt = xls.GetCellVisibleFormatDef(215, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(215, 2, xls.AddFormat(fmt));
            xls.SetCellValue(215, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(215, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(215, 3, xls.AddFormat(fmt));
            xls.SetCellValue(215, 3, new TFormula("='Inputs 2.0 Conv. default values'!I215"));

            fmt = xls.GetCellVisibleFormatDef(216, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(216, 2, xls.AddFormat(fmt));
            xls.SetCellValue(216, 2, "Baskets");

            fmt = xls.GetCellVisibleFormatDef(216, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(216, 3, xls.AddFormat(fmt));
            xls.SetCellValue(216, 3, new TFormula("='Inputs 2.0 Conv. default values'!I216"));

            fmt = xls.GetCellVisibleFormatDef(217, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(217, 2, xls.AddFormat(fmt));
            xls.SetCellValue(217, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(217, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(217, 3, xls.AddFormat(fmt));
            xls.SetCellValue(217, 3, new TFormula("='Inputs 2.0 Conv. default values'!I217"));

            fmt = xls.GetCellVisibleFormatDef(218, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(218, 2, xls.AddFormat(fmt));
            xls.SetCellValue(218, 2, "Boxes");

            fmt = xls.GetCellVisibleFormatDef(218, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(218, 3, xls.AddFormat(fmt));
            xls.SetCellValue(218, 3, new TFormula("='Inputs 2.0 Conv. default values'!I218"));

            fmt = xls.GetCellVisibleFormatDef(219, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(219, 2, xls.AddFormat(fmt));
            xls.SetCellValue(219, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(219, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(219, 3, xls.AddFormat(fmt));
            xls.SetCellValue(219, 3, new TFormula("='Inputs 2.0 Conv. default values'!I219"));

            fmt = xls.GetCellVisibleFormatDef(220, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(220, 2, xls.AddFormat(fmt));
            xls.SetCellValue(220, 2, "Other(s):");

            fmt = xls.GetCellVisibleFormatDef(220, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(220, 3, xls.AddFormat(fmt));
            xls.SetCellValue(220, 3, new TFormula("='Inputs 2.0 Conv. default values'!I220"));

            fmt = xls.GetCellVisibleFormatDef(221, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(221, 2, xls.AddFormat(fmt));
            xls.SetCellValue(221, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(221, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(221, 3, xls.AddFormat(fmt));
            xls.SetCellValue(221, 3, new TFormula("='Inputs 2.0 Conv. default values'!I221"));

            fmt = xls.GetCellVisibleFormatDef(222, 1);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xD8, 0xD8, 0xD8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(222, 2, xls.AddFormat(fmt));
            xls.SetCellValue(222, 2, "Equipment and Materials for processing");

            fmt = xls.GetCellVisibleFormatDef(222, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 4);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 5);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 6);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 7);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 8);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 9);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(222, 10);
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(222, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(223, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(223, 2, xls.AddFormat(fmt));
            xls.SetCellValue(223, 2, "Pulper machine");

            fmt = xls.GetCellVisibleFormatDef(223, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(223, 3, xls.AddFormat(fmt));
            xls.SetCellValue(223, 3, new TFormula("='Inputs 2.0 Conv. default values'!I223"));

            fmt = xls.GetCellVisibleFormatDef(224, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(224, 2, xls.AddFormat(fmt));
            xls.SetCellValue(224, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(224, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(224, 3, xls.AddFormat(fmt));
            xls.SetCellValue(224, 3, new TFormula("='Inputs 2.0 Conv. default values'!I224"));

            fmt = xls.GetCellVisibleFormatDef(225, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(225, 2, xls.AddFormat(fmt));
            xls.SetCellValue(225, 2, "Tolca");

            fmt = xls.GetCellVisibleFormatDef(225, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(225, 3, xls.AddFormat(fmt));
            xls.SetCellValue(225, 3, new TFormula("='Inputs 2.0 Conv. default values'!I225"));

            fmt = xls.GetCellVisibleFormatDef(226, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(226, 2, xls.AddFormat(fmt));
            xls.SetCellValue(226, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(226, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(226, 3, xls.AddFormat(fmt));
            xls.SetCellValue(226, 3, new TFormula("='Inputs 2.0 Conv. default values'!I226"));

            fmt = xls.GetCellVisibleFormatDef(227, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(227, 2, xls.AddFormat(fmt));
            xls.SetCellValue(227, 2, "Engine");

            fmt = xls.GetCellVisibleFormatDef(227, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(227, 3, xls.AddFormat(fmt));
            xls.SetCellValue(227, 3, new TFormula("='Inputs 2.0 Conv. default values'!I227"));

            fmt = xls.GetCellVisibleFormatDef(228, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(228, 2, xls.AddFormat(fmt));
            xls.SetCellValue(228, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(228, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(228, 3, xls.AddFormat(fmt));
            xls.SetCellValue(228, 3, new TFormula("='Inputs 2.0 Conv. default values'!I228"));

            fmt = xls.GetCellVisibleFormatDef(229, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(229, 2, xls.AddFormat(fmt));
            xls.SetCellValue(229, 2, "Tanks for fermentation");

            fmt = xls.GetCellVisibleFormatDef(229, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(229, 3, xls.AddFormat(fmt));
            xls.SetCellValue(229, 3, new TFormula("='Inputs 2.0 Conv. default values'!I229"));

            fmt = xls.GetCellVisibleFormatDef(230, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(230, 2, xls.AddFormat(fmt));
            xls.SetCellValue(230, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(230, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(230, 3, xls.AddFormat(fmt));
            xls.SetCellValue(230, 3, new TFormula("='Inputs 2.0 Conv. default values'!I230"));

            fmt = xls.GetCellVisibleFormatDef(231, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(231, 2, xls.AddFormat(fmt));
            xls.SetCellValue(231, 2, "Water channel for coffee washing");

            fmt = xls.GetCellVisibleFormatDef(231, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(231, 3, xls.AddFormat(fmt));
            xls.SetCellValue(231, 3, new TFormula("='Inputs 2.0 Conv. default values'!I231"));

            fmt = xls.GetCellVisibleFormatDef(232, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(232, 2, xls.AddFormat(fmt));
            xls.SetCellValue(232, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(232, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(232, 3, xls.AddFormat(fmt));
            xls.SetCellValue(232, 3, new TFormula("='Inputs 2.0 Conv. default values'!I232"));

            fmt = xls.GetCellVisibleFormatDef(233, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(233, 2, xls.AddFormat(fmt));
            xls.SetCellValue(233, 2, "PVC pipes");

            fmt = xls.GetCellVisibleFormatDef(233, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(233, 3, xls.AddFormat(fmt));
            xls.SetCellValue(233, 3, new TFormula("='Inputs 2.0 Conv. default values'!I233"));

            fmt = xls.GetCellVisibleFormatDef(234, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(234, 2, xls.AddFormat(fmt));
            xls.SetCellValue(234, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(234, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(234, 3, xls.AddFormat(fmt));
            xls.SetCellValue(234, 3, new TFormula("='Inputs 2.0 Conv. default values'!I234"));

            fmt = xls.GetCellVisibleFormatDef(235, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(235, 2, xls.AddFormat(fmt));
            xls.SetCellValue(235, 2, "Water filtering system (organic farm)");

            fmt = xls.GetCellVisibleFormatDef(235, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(235, 3, xls.AddFormat(fmt));
            xls.SetCellValue(235, 3, new TFormula("='Inputs 2.0 Conv. default values'!I235"));

            fmt = xls.GetCellVisibleFormatDef(236, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(236, 2, xls.AddFormat(fmt));
            xls.SetCellValue(236, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(236, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(236, 3, xls.AddFormat(fmt));
            xls.SetCellValue(236, 3, new TFormula("='Inputs 2.0 Conv. default values'!I236"));

            fmt = xls.GetCellVisibleFormatDef(237, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(237, 2, xls.AddFormat(fmt));
            xls.SetCellValue(237, 2, "Sieve or screening machine");

            fmt = xls.GetCellVisibleFormatDef(237, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(237, 3, xls.AddFormat(fmt));
            xls.SetCellValue(237, 3, new TFormula("='Inputs 2.0 Conv. default values'!I237"));

            fmt = xls.GetCellVisibleFormatDef(238, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(238, 2, xls.AddFormat(fmt));
            xls.SetCellValue(238, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(238, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(238, 3, xls.AddFormat(fmt));
            xls.SetCellValue(238, 3, new TFormula("='Inputs 2.0 Conv. default values'!I238"));

            fmt = xls.GetCellVisibleFormatDef(239, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(239, 2, xls.AddFormat(fmt));
            xls.SetCellValue(239, 2, "Desmucilaginador Machine to remove mucilage");

            fmt = xls.GetCellVisibleFormatDef(239, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(239, 3, xls.AddFormat(fmt));
            xls.SetCellValue(239, 3, new TFormula("='Inputs 2.0 Conv. default values'!I239"));

            fmt = xls.GetCellVisibleFormatDef(240, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(240, 2, xls.AddFormat(fmt));
            xls.SetCellValue(240, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(240, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(240, 3, xls.AddFormat(fmt));
            xls.SetCellValue(240, 3, new TFormula("='Inputs 2.0 Conv. default values'!I240"));

            fmt = xls.GetCellVisibleFormatDef(241, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(241, 2, xls.AddFormat(fmt));
            xls.SetCellValue(241, 2, "Motor pump");

            fmt = xls.GetCellVisibleFormatDef(241, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(241, 3, xls.AddFormat(fmt));
            xls.SetCellValue(241, 3, new TFormula("='Inputs 2.0 Conv. default values'!I241"));

            fmt = xls.GetCellVisibleFormatDef(242, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(242, 2, xls.AddFormat(fmt));
            xls.SetCellValue(242, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(242, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(242, 3, xls.AddFormat(fmt));
            xls.SetCellValue(242, 3, new TFormula("='Inputs 2.0 Conv. default values'!I242"));

            fmt = xls.GetCellVisibleFormatDef(243, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(243, 2, xls.AddFormat(fmt));
            xls.SetCellValue(243, 2, "Other input(s) for the wet processing:");

            fmt = xls.GetCellVisibleFormatDef(243, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(243, 3, xls.AddFormat(fmt));
            xls.SetCellValue(243, 3, new TFormula("='Inputs 2.0 Conv. default values'!I243"));

            fmt = xls.GetCellVisibleFormatDef(244, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(244, 2, xls.AddFormat(fmt));
            xls.SetCellValue(244, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(244, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(244, 3, xls.AddFormat(fmt));
            xls.SetCellValue(244, 3, new TFormula("='Inputs 2.0 Conv. default values'!I244"));

            fmt = xls.GetCellVisibleFormatDef(245, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(245, 2, xls.AddFormat(fmt));
            xls.SetCellValue(245, 2, "Concrete yard / patio");

            fmt = xls.GetCellVisibleFormatDef(245, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(245, 3, xls.AddFormat(fmt));
            xls.SetCellValue(245, 3, new TFormula("='Inputs 2.0 Conv. default values'!I245"));

            fmt = xls.GetCellVisibleFormatDef(246, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(246, 2, xls.AddFormat(fmt));
            xls.SetCellValue(246, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(246, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(246, 3, xls.AddFormat(fmt));
            xls.SetCellValue(246, 3, new TFormula("='Inputs 2.0 Conv. default values'!I246"));

            fmt = xls.GetCellVisibleFormatDef(247, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(247, 2, xls.AddFormat(fmt));
            xls.SetCellValue(247, 2, "Plastic");

            fmt = xls.GetCellVisibleFormatDef(247, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(247, 3, xls.AddFormat(fmt));
            xls.SetCellValue(247, 3, new TFormula("='Inputs 2.0 Conv. default values'!I247"));

            fmt = xls.GetCellVisibleFormatDef(248, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(248, 2, xls.AddFormat(fmt));
            xls.SetCellValue(248, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(248, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(248, 3, xls.AddFormat(fmt));
            xls.SetCellValue(248, 3, new TFormula("='Inputs 2.0 Conv. default values'!I248"));

            fmt = xls.GetCellVisibleFormatDef(249, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(249, 2, xls.AddFormat(fmt));
            xls.SetCellValue(249, 2, "Rake");

            fmt = xls.GetCellVisibleFormatDef(249, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(249, 3, xls.AddFormat(fmt));
            xls.SetCellValue(249, 3, new TFormula("='Inputs 2.0 Conv. default values'!I249"));

            fmt = xls.GetCellVisibleFormatDef(250, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(250, 2, xls.AddFormat(fmt));
            xls.SetCellValue(250, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(250, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(250, 3, xls.AddFormat(fmt));
            xls.SetCellValue(250, 3, new TFormula("='Inputs 2.0 Conv. default values'!I250"));

            fmt = xls.GetCellVisibleFormatDef(251, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(251, 2, xls.AddFormat(fmt));
            xls.SetCellValue(251, 2, "Broom");

            fmt = xls.GetCellVisibleFormatDef(251, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(251, 3, xls.AddFormat(fmt));
            xls.SetCellValue(251, 3, new TFormula("='Inputs 2.0 Conv. default values'!I251"));

            fmt = xls.GetCellVisibleFormatDef(252, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(252, 2, xls.AddFormat(fmt));
            xls.SetCellValue(252, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(252, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(252, 3, xls.AddFormat(fmt));
            xls.SetCellValue(252, 3, new TFormula("='Inputs 2.0 Conv. default values'!I252"));

            fmt = xls.GetCellVisibleFormatDef(253, 2);
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(253, 2, xls.AddFormat(fmt));
            xls.SetCellValue(253, 2, "Storage room");

            fmt = xls.GetCellVisibleFormatDef(253, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(253, 3, xls.AddFormat(fmt));
            xls.SetCellValue(253, 3, new TFormula("='Inputs 2.0 Conv. default values'!I253"));

            fmt = xls.GetCellVisibleFormatDef(254, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(254, 2, xls.AddFormat(fmt));
            xls.SetCellValue(254, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(254, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(254, 3, xls.AddFormat(fmt));
            xls.SetCellValue(254, 3, new TFormula("='Inputs 2.0 Conv. default values'!I254"));

            fmt = xls.GetCellVisibleFormatDef(255, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(255, 2, xls.AddFormat(fmt));
            xls.SetCellValue(255, 2, "Other input(s) for the dry processing:");

            fmt = xls.GetCellVisibleFormatDef(255, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(255, 3, xls.AddFormat(fmt));
            xls.SetCellValue(255, 3, new TFormula("='Inputs 2.0 Conv. default values'!I255"));

            fmt = xls.GetCellVisibleFormatDef(256, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.right;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(256, 2, xls.AddFormat(fmt));
            xls.SetCellValue(256, 2, "Lifespam ");

            fmt = xls.GetCellVisibleFormatDef(256, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(256, 3, xls.AddFormat(fmt));
            xls.SetCellValue(256, 3, new TFormula("='Inputs 2.0 Conv. default values'!I256"));

            fmt = xls.GetCellVisibleFormatDef(257, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(257, 2, xls.AddFormat(fmt));
            xls.SetCellValue(257, 2, "Administrative costs, taxes and land");

            fmt = xls.GetCellVisibleFormatDef(257, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(257, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(257, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(258, 2, xls.AddFormat(fmt));
            xls.SetCellValue(258, 2, "Cooperative membership expenses");

            fmt = xls.GetCellVisibleFormatDef(258, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 4);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 5);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 6);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 7);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 8);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 9);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(258, 10);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(258, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(259, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(259, 2, xls.AddFormat(fmt));
            xls.SetCellValue(259, 2, "Application fee to entrance the cooperative");

            fmt = xls.GetCellVisibleFormatDef(259, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(259, 3, xls.AddFormat(fmt));
            xls.SetCellValue(259, 3, new TFormula("='Inputs 2.0 Conv. default values'!I259"));

            fmt = xls.GetCellVisibleFormatDef(260, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 2, xls.AddFormat(fmt));
            xls.SetCellValue(260, 2, " Annual membership to the cooperative");

            fmt = xls.GetCellVisibleFormatDef(260, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 3, xls.AddFormat(fmt));
            xls.SetCellValue(260, 3, new TFormula("='Inputs 2.0 Conv. default values'!I260"));

            fmt = xls.GetCellVisibleFormatDef(260, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(260, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(260, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 2, xls.AddFormat(fmt));
            xls.SetCellValue(261, 2, "Life insurance");

            fmt = xls.GetCellVisibleFormatDef(261, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 3, xls.AddFormat(fmt));
            xls.SetCellValue(261, 3, new TFormula("='Inputs 2.0 Conv. default values'!I261"));

            fmt = xls.GetCellVisibleFormatDef(261, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(261, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(261, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 2, xls.AddFormat(fmt));
            xls.SetCellValue(262, 2, "FLO Certificatoin");

            fmt = xls.GetCellVisibleFormatDef(262, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 3, xls.AddFormat(fmt));
            xls.SetCellValue(262, 3, new TFormula("='Inputs 2.0 Conv. default values'!I262"));

            fmt = xls.GetCellVisibleFormatDef(262, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(262, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(262, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 2, xls.AddFormat(fmt));
            xls.SetCellValue(263, 2, "Organic Certification");

            fmt = xls.GetCellVisibleFormatDef(263, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 3, xls.AddFormat(fmt));
            xls.SetCellValue(263, 3, new TFormula("='Inputs 2.0 Conv. default values'!I263"));

            fmt = xls.GetCellVisibleFormatDef(263, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(263, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(263, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(264, 2, xls.AddFormat(fmt));
            xls.SetCellValue(264, 2, "Land");

            fmt = xls.GetCellVisibleFormatDef(264, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 4);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 5);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 6);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 7);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 8);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(264, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(264, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(265, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(265, 2, xls.AddFormat(fmt));
            xls.SetCellValue(265, 2, new TFormula("=+\"Value of your land in  \"&'Gral Conf. Summary'!$H$33&\" per  \"&'Gral Conf. Summary'!$I$23&\""
            + " (without crop)\""));

            fmt = xls.GetCellVisibleFormatDef(265, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(265, 3, xls.AddFormat(fmt));
            xls.SetCellValue(265, 3, new TFormula("='Inputs 2.0 Conv. default values'!I265"));

            fmt = xls.GetCellVisibleFormatDef(266, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Family = 0;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(266, 2, xls.AddFormat(fmt));
            xls.SetCellValue(266, 2, new TFormula("=+\"Property tax  in  \"&'Gral Conf. Summary'!$H$33&\" \""));

            fmt = xls.GetCellVisibleFormatDef(266, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.right;
            xls.SetCellFormat(266, 3, xls.AddFormat(fmt));
            xls.SetCellValue(266, 3, new TFormula("='Inputs 2.0 Conv. default values'!I266"));

            fmt = xls.GetCellVisibleFormatDef(267, 1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 2, xls.AddFormat(fmt));
            xls.SetCellValue(267, 2, "Administrative costs and unexpected events");

            fmt = xls.GetCellVisibleFormatDef(267, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 4);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 5);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 6);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 7);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 8);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 9);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(267, 10);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(267, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(268, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(268, 2, xls.AddFormat(fmt));
            xls.SetCellValue(268, 2, "How many days per year can you invest in supervising activities as cleaning, weeding,"
            + " management, pruning, maintenance, harvest, etc");

            fmt = xls.GetCellVisibleFormatDef(268, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(268, 3, xls.AddFormat(fmt));
            xls.SetCellValue(268, 3, new TFormula("='Inputs 2.0 Conv. default values'!I268"));

            fmt = xls.GetCellVisibleFormatDef(269, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(269, 2, xls.AddFormat(fmt));
            xls.SetCellValue(269, 2, "How many days per year can you invest in administrative matters related to your farm,"
            + " for example, keeping records, paying bills, paying hired workers, going to the bank,"
            + " doing paperwork at the cooperative, meetings (not training sessions)");

            fmt = xls.GetCellVisibleFormatDef(269, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(269, 3, xls.AddFormat(fmt));
            xls.SetCellValue(269, 3, new TFormula("='Inputs 2.0 Conv. default values'!I269"));

            fmt = xls.GetCellVisibleFormatDef(270, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(270, 2, xls.AddFormat(fmt));
            xls.SetCellValue(270, 2, "How many days per year can you invest in providing training sessions for the workers"
            + " you hire for different farm related activities? ");

            fmt = xls.GetCellVisibleFormatDef(270, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(270, 3, xls.AddFormat(fmt));
            xls.SetCellValue(270, 3, new TFormula("='Inputs 2.0 Conv. default values'!I270"));

            fmt = xls.GetCellVisibleFormatDef(271, 2);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(271, 2, xls.AddFormat(fmt));
            xls.SetCellValue(271, 2, new TFormula("=+\"How much in  \"&'Gral Conf. Summary'!$H$33&\" can you invest in extraordinary"
            + " events such as medical assistance for work accidents to you employees?\""));

            fmt = xls.GetCellVisibleFormatDef(271, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(271, 3, xls.AddFormat(fmt));
            xls.SetCellValue(271, 3, new TFormula("='Inputs 2.0 Conv. default values'!I271"));

            fmt = xls.GetCellVisibleFormatDef(272, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 2);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(272, 2, xls.AddFormat(fmt));
            xls.SetCellValue(272, 2, "Transportation");

            fmt = xls.GetCellVisibleFormatDef(272, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 5);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 7);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(272, 10);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(272, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(273, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(273, 2, xls.AddFormat(fmt));
            xls.SetCellValue(273, 2, new TFormula("=+\"Please, describe how much do you spent in  \"&'Gral Conf. Summary'!$H$33&\" in"
            + " the following transportation activities related to the coffee produced in  ONE \"&'Gral"
            + " Conf. Summary'!$I$23&\" of coffee\""));

            fmt = xls.GetCellVisibleFormatDef(273, 3);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(273, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 1);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xD8, 0xD8, 0xD8);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(274, 2, xls.AddFormat(fmt));
            xls.SetCellValue(274, 2, "Transportation activities realted to the germinator");

            fmt = xls.GetCellVisibleFormatDef(274, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 4);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 5);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 6);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 7);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 8);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 9);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(274, 10);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(274, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(275, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(275, 2, xls.AddFormat(fmt));
            xls.SetCellValue(275, 2, "Seed purchase trip");

            fmt = xls.GetCellVisibleFormatDef(275, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(275, 3, xls.AddFormat(fmt));
            xls.SetCellValue(275, 3, new TFormula("='Inputs 2.0 Conv. default values'!I275"));

            fmt = xls.GetCellVisibleFormatDef(276, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(276, 2, xls.AddFormat(fmt));
            xls.SetCellValue(276, 2, "Wood transportation");

            fmt = xls.GetCellVisibleFormatDef(276, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(276, 3, xls.AddFormat(fmt));
            xls.SetCellValue(276, 3, new TFormula("='Inputs 2.0 Conv. default values'!I276"));

            fmt = xls.GetCellVisibleFormatDef(277, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(277, 2, xls.AddFormat(fmt));
            xls.SetCellValue(277, 2, "Transportation of sand for the germinator");

            fmt = xls.GetCellVisibleFormatDef(277, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(277, 3, xls.AddFormat(fmt));
            xls.SetCellValue(277, 3, new TFormula("='Inputs 2.0 Conv. default values'!I277"));

            fmt = xls.GetCellVisibleFormatDef(278, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(278, 2, xls.AddFormat(fmt));
            xls.SetCellValue(278, 2, "Other(s):");

            fmt = xls.GetCellVisibleFormatDef(278, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(278, 3, xls.AddFormat(fmt));
            xls.SetCellValue(278, 3, new TFormula("='Inputs 2.0 Conv. default values'!I278"));

            fmt = xls.GetCellVisibleFormatDef(279, 1);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xD8, 0xD8, 0xD8);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(279, 2, xls.AddFormat(fmt));
            xls.SetCellValue(279, 2, "Transportation activities realted to the nursery");

            fmt = xls.GetCellVisibleFormatDef(279, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 4);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 5);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 6);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 7);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 8);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 9);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(279, 10);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(279, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(280, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(280, 2, xls.AddFormat(fmt));
            xls.SetCellValue(280, 2, "Soil transportation");

            fmt = xls.GetCellVisibleFormatDef(280, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(280, 3, xls.AddFormat(fmt));
            xls.SetCellValue(280, 3, new TFormula("='Inputs 2.0 Conv. default values'!I280"));

            fmt = xls.GetCellVisibleFormatDef(281, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(281, 2, xls.AddFormat(fmt));
            xls.SetCellValue(281, 2, "Sacks and other material shopping for the nursery");

            fmt = xls.GetCellVisibleFormatDef(281, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(281, 3, xls.AddFormat(fmt));
            xls.SetCellValue(281, 3, new TFormula("='Inputs 2.0 Conv. default values'!I281"));

            fmt = xls.GetCellVisibleFormatDef(282, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(282, 2, xls.AddFormat(fmt));
            xls.SetCellValue(282, 2, "Other(s):");

            fmt = xls.GetCellVisibleFormatDef(282, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(282, 3, xls.AddFormat(fmt));
            xls.SetCellValue(282, 3, new TFormula("='Inputs 2.0 Conv. default values'!I282"));

            fmt = xls.GetCellVisibleFormatDef(283, 1);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xD8, 0xD8, 0xD8);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(283, 2, xls.AddFormat(fmt));
            xls.SetCellValue(283, 2, "Transportation activities realted to the land preparation and planting");

            fmt = xls.GetCellVisibleFormatDef(283, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 4);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 5);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 6);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 7);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 8);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 9);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(283, 10);
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(283, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(284, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(284, 2, xls.AddFormat(fmt));
            xls.SetCellValue(284, 2, "Wood transportation");

            fmt = xls.GetCellVisibleFormatDef(284, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(284, 3, xls.AddFormat(fmt));
            xls.SetCellValue(284, 3, new TFormula("='Inputs 2.0 Conv. default values'!I284"));

            fmt = xls.GetCellVisibleFormatDef(285, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(285, 2, xls.AddFormat(fmt));
            xls.SetCellValue(285, 2, "Compost transportation");

            fmt = xls.GetCellVisibleFormatDef(285, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(285, 3, xls.AddFormat(fmt));
            xls.SetCellValue(285, 3, new TFormula("='Inputs 2.0 Conv. default values'!I285"));

            fmt = xls.GetCellVisibleFormatDef(286, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(286, 2, xls.AddFormat(fmt));
            xls.SetCellValue(286, 2, "Plant transportation from the nursery to the land");

            fmt = xls.GetCellVisibleFormatDef(286, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(286, 3, xls.AddFormat(fmt));
            xls.SetCellValue(286, 3, new TFormula("='Inputs 2.0 Conv. default values'!I286"));

            fmt = xls.GetCellVisibleFormatDef(287, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(287, 2, xls.AddFormat(fmt));
            xls.SetCellValue(287, 2, "Other(s):");

            fmt = xls.GetCellVisibleFormatDef(287, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(287, 3, xls.AddFormat(fmt));
            xls.SetCellValue(287, 3, new TFormula("='Inputs 2.0 Conv. default values'!I287"));

            fmt = xls.GetCellVisibleFormatDef(288, 1);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold | TFlxFontStyles.Italic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0xD8, 0xD8, 0xD8);
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(288, 2, xls.AddFormat(fmt));
            xls.SetCellValue(288, 2, "Other transportation expenses, annual sums");

            fmt = xls.GetCellVisibleFormatDef(288, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 4);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 5);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 6);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 7);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 8);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 9);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(288, 10);
            fmt.HAlignment = THFlxAlignment.left;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(288, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(289, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(289, 2, xls.AddFormat(fmt));
            xls.SetCellValue(289, 2, "Equipment and tools transportation");

            fmt = xls.GetCellVisibleFormatDef(289, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(289, 3, xls.AddFormat(fmt));
            xls.SetCellValue(289, 3, new TFormula("='Inputs 2.0 Conv. default values'!I289"));

            fmt = xls.GetCellVisibleFormatDef(290, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(290, 2, xls.AddFormat(fmt));
            xls.SetCellValue(290, 2, "Labor / workforce transportation (not included in the daily wage)");

            fmt = xls.GetCellVisibleFormatDef(290, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(290, 3, xls.AddFormat(fmt));
            xls.SetCellValue(290, 3, new TFormula("='Inputs 2.0 Conv. default values'!I290"));

            fmt = xls.GetCellVisibleFormatDef(291, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(291, 2, xls.AddFormat(fmt));
            xls.SetCellValue(291, 2, "Coffee transportation to the collection center or cooperative");

            fmt = xls.GetCellVisibleFormatDef(291, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(291, 3, xls.AddFormat(fmt));
            xls.SetCellValue(291, 3, new TFormula("='Inputs 2.0 Conv. default values'!I291"));

            fmt = xls.GetCellVisibleFormatDef(292, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            fmt.WrapText = true;
            xls.SetCellFormat(292, 2, xls.AddFormat(fmt));
            xls.SetCellValue(292, 2, "Transportation for supervising activities (weeding, management, pruning, maintenance"
            + " work)");

            fmt = xls.GetCellVisibleFormatDef(292, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(292, 3, xls.AddFormat(fmt));
            xls.SetCellValue(292, 3, new TFormula("='Inputs 2.0 Conv. default values'!I292"));

            fmt = xls.GetCellVisibleFormatDef(293, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Borders.Left.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Left.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.VAlignment = TVFlxAlignment.center;
            xls.SetCellFormat(293, 2, xls.AddFormat(fmt));
            xls.SetCellValue(293, 2, "Other(s) transportation not considered:");

            fmt = xls.GetCellVisibleFormatDef(293, 3);
            fmt.Borders.Right.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Right.Color = TExcelColor.Automatic;
            fmt.Borders.Bottom.Style = TFlxBorderStyle.Medium;
            fmt.Borders.Bottom.Color = TExcelColor.Automatic;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Format = "0.00";
            xls.SetCellFormat(293, 3, xls.AddFormat(fmt));
            xls.SetCellValue(293, 3, new TFormula("='Inputs 2.0 Conv. default values'!I293"));

            //Cell selection and scroll position.
            xls.SelectCell(25, 6, false);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");
            //xls.Recalc();
            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-08T03:31:31Z");

        }

    }
}
