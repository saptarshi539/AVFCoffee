using FlexCel.Core;

namespace CoffeeInfrastructure.Flexcel
{
    public class Gral_Conf
    {
        public void Gral_Conf_Summary(ExcelFile xls)
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

            xls.ActiveSheet = 21;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Gral Conf. Summary";

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
            xls.DefaultColWidth = 2784;

            xls.SetColWidth(1, 1, 2784);    //(10.13 + 0.75) * 256

            xls.SetColWidth(2, 2, 3552);    //(13.13 + 0.75) * 256

            xls.SetColWidth(3, 3, 5216);    //(19.63 + 0.75) * 256

            xls.SetColWidth(4, 4, 4128);    //(15.38 + 0.75) * 256

            xls.SetColWidth(5, 11, 2784);    //(10.13 + 0.75) * 256

            xls.SetColWidth(12, 12, 4576);    //(17.13 + 0.75) * 256

            xls.SetColWidth(13, 13, 4064);    //(15.13 + 0.75) * 256

            xls.SetColWidth(14, 16384, 2784);    //(10.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(3, 375);    //18.75 * 20
            xls.SetRowHeight(10, 360);    //18.00 * 20
            xls.SetRowHeight(12, 375);    //18.75 * 20

            //Merged Cells
            xls.MergeCells(3, 2, 3, 14);
            xls.MergeCells(10, 8, 10, 9);

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(3, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 2, xls.AddFormat(fmt));
            xls.SetCellValue(3, 2, "CONFIGURATION");

            fmt = xls.GetCellVisibleFormatDef(3, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 4);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 5);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 6);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 7);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 8);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 9);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 10);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 11);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 12);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 13);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 14);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x80, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(3, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
            xls.SetCellValue(5, 2, "Metric and meassure system definition");

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 12);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 13);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(5, 14);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background2, -0.0999786370433668);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 14, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, "Please answer the following questions and select one option in each question");
            xls.SetCellValue(8, 11, "To Mexican metrics");
            xls.SetCellValue(10, 2, "The most commom meassure for coffee parchment production is:");

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));
            xls.SetCellValue(10, 8, "Meassure user");

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 12);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(10, 12, xls.AddFormat(fmt));
            xls.SetCellValue(10, 12, "Quintales");

            fmt = xls.GetCellVisibleFormatDef(10, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 16);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(10, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));
            xls.SetCellValue(11, 4, new TFormula("=Metrics!D11"));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));
            xls.SetCellValue(11, 5, "kilograms");

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));
            xls.SetCellValue(11, 6, "kilogram");

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(11, 11, xls.AddFormat(fmt));
            xls.SetCellValue(11, 11, "kilograms");

            fmt = xls.GetCellVisibleFormatDef(11, 12);
            fmt.Format = "0.00";
            xls.SetCellFormat(11, 12, xls.AddFormat(fmt));
            xls.SetCellValue(11, 12, new TFormula("=1/Conversiones!$D$14"));
            xls.SetCellValue(11, 13, "quintal/kg");

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));
            xls.SetCellValue(12, 4, new TFormula("=Metrics!D12"));

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));
            xls.SetCellValue(12, 5, "pounds");

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));
            xls.SetCellValue(12, 6, "pound");

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.Font.Size20 = 280;
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Family = 0;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));
            xls.SetCellValue(12, 11, "kilogram");

            fmt = xls.GetCellVisibleFormatDef(12, 12);
            fmt.Format = "0.00";
            xls.SetCellFormat(12, 12, xls.AddFormat(fmt));
            xls.SetCellValue(12, 12, new TFormula("=L11"));

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));
            xls.SetCellValue(13, 4, new TFormula("=Metrics!D13"));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));
            xls.SetCellValue(13, 5, "quintales");

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));
            xls.SetCellValue(13, 6, "quintal");

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(13, 11, xls.AddFormat(fmt));
            xls.SetCellValue(13, 11, "pounds");
            xls.SetCellValue(13, 12, new TFormula("=1/Conversiones!$C$14"));
            xls.SetCellValue(13, 13, "quintal/lb");

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));
            xls.SetCellValue(14, 4, new TFormula("=Metrics!D14"));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));
            xls.SetCellValue(14, 5, "arrobas");

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));
            xls.SetCellValue(14, 6, "arroba");

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));
            xls.SetCellValue(14, 11, "pound");
            xls.SetCellValue(14, 12, new TFormula("=L13"));

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));
            xls.SetCellValue(15, 4, new TFormula("=Metrics!D16"));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));
            xls.SetCellValue(15, 5, "cargas");

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));
            xls.SetCellValue(15, 6, "carga");

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));
            xls.SetCellValue(15, 8, new TFormula("=VLOOKUP(1,$D$11:$E$15,2,FALSE)"));
            xls.SetCellValue(15, 9, new TFormula("=VLOOKUP(1,$D$11:$F$15,3,FALSE)"));

            fmt = xls.GetCellVisibleFormatDef(15, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(15, 11, xls.AddFormat(fmt));
            xls.SetCellValue(15, 11, "quintales");
            xls.SetCellValue(15, 12, 1);
            xls.SetCellValue(15, 13, "quintal/quintal");
            xls.SetCellValue(15, 15, new TFormula("=H15"));
            xls.SetCellValue(15, 16, new TFormula("=VLOOKUP(O15,K11:L19,2,FALSE)"));

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));
            xls.SetCellValue(16, 11, "quintal");
            xls.SetCellValue(16, 12, new TFormula("=L15"));

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, "The most commom meassure of length is:");

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(17, 11, xls.AddFormat(fmt));
            xls.SetCellValue(17, 11, "arrobas");
            xls.SetCellValue(17, 12, new TFormula("=Conversiones!$D$17/Conversiones!$D$14"));
            xls.SetCellValue(17, 13, "quintal/arroba");

            fmt = xls.GetCellVisibleFormatDef(17, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(17, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));
            xls.SetCellValue(18, 4, new TFormula("=Metrics!D18"));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));
            xls.SetCellValue(18, 5, "meters");

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));
            xls.SetCellValue(18, 11, "arroba");
            xls.SetCellValue(18, 12, new TFormula("=L17"));

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));
            xls.SetCellValue(19, 4, new TFormula("=Metrics!D19"));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));
            xls.SetCellValue(19, 5, "feet");

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));
            xls.SetCellValue(19, 8, new TFormula("=VLOOKUP(1,$D$18:$E$19,2,FALSE)"));

            fmt = xls.GetCellVisibleFormatDef(19, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(19, 11, xls.AddFormat(fmt));
            xls.SetCellValue(19, 11, "cargas");
            xls.SetCellValue(19, 12, new TFormula("=Conversiones!D20/Conversiones!D14"));
            xls.SetCellValue(19, 13, "quintal/carga");

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));
            xls.SetCellValue(20, 11, "carga");
            xls.SetCellValue(20, 12, new TFormula("=L19"));

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, "The most commom meassure for the farm area is:");

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 12);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(21, 12, xls.AddFormat(fmt));
            xls.SetCellValue(21, 12, "Metros");

            fmt = xls.GetCellVisibleFormatDef(21, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(21, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));
            xls.SetCellValue(22, 4, new TFormula("=Metrics!D22"));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));
            xls.SetCellValue(22, 5, "hectares");

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));
            xls.SetCellValue(22, 6, "hectare");

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(22, 11, xls.AddFormat(fmt));
            xls.SetCellValue(22, 11, "meters");
            xls.SetCellValue(22, 12, 1);
            xls.SetCellValue(22, 13, "m/m");

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));
            xls.SetCellValue(23, 4, new TFormula("=Metrics!D23"));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));
            xls.SetCellValue(23, 5, "manzanas");

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));
            xls.SetCellValue(23, 6, "manzana");

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));
            xls.SetCellValue(23, 8, new TFormula("=VLOOKUP(1,$D$22:$E$23,2,FALSE)"));
            xls.SetCellValue(23, 9, new TFormula("=VLOOKUP(1,$D$22:$F$23,3,FALSE)"));
            xls.SetCellValue(23, 11, "meter");
            xls.SetCellValue(23, 12, new TFormula("=L22"));

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(24, 11, xls.AddFormat(fmt));
            xls.SetCellValue(24, 11, "feet");
            xls.SetCellValue(24, 12, 0.3048);
            xls.SetCellValue(24, 13, "m/ft");

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));
            xls.SetCellValue(25, 2, "The most commom meassure for inputs weight (Ex: fertilizers is):");

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));
            xls.SetCellValue(25, 11, "foot");
            xls.SetCellValue(25, 12, new TFormula("=L24"));

            fmt = xls.GetCellVisibleFormatDef(25, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(25, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));
            xls.SetCellValue(26, 4, new TFormula("=Metrics!D26"));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));
            xls.SetCellValue(26, 5, "kilograms");

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 12);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(26, 12, xls.AddFormat(fmt));
            xls.SetCellValue(26, 12, "Hectarea");

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));
            xls.SetCellValue(27, 4, new TFormula("=Metrics!D27"));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));
            xls.SetCellValue(27, 5, "pounds");

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));
            xls.SetCellValue(27, 8, new TFormula("=VLOOKUP(1,$D$26:$E$27,2,FALSE)"));

            fmt = xls.GetCellVisibleFormatDef(27, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(27, 11, xls.AddFormat(fmt));
            xls.SetCellValue(27, 11, new TFormula("=E22"));
            xls.SetCellValue(27, 12, 1);
            xls.SetCellValue(27, 13, "ht/ht");
            xls.SetCellValue(28, 11, "hectare");
            xls.SetCellValue(28, 12, new TFormula("=L27"));

            fmt = xls.GetCellVisibleFormatDef(29, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, "The most commom meassure for capacity:");

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(29, 11, xls.AddFormat(fmt));
            xls.SetCellValue(29, 11, new TFormula("=E23"));
            xls.SetCellValue(29, 12, new TFormula("=1/Conversiones!$C$7"));
            xls.SetCellValue(29, 13, "ht/manz");

            fmt = xls.GetCellVisibleFormatDef(29, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(29, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));
            xls.SetCellValue(30, 4, new TFormula("=Metrics!D30"));

            fmt = xls.GetCellVisibleFormatDef(30, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(30, 5, xls.AddFormat(fmt));
            xls.SetCellValue(30, 5, "liters");

            fmt = xls.GetCellVisibleFormatDef(30, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(30, 6, xls.AddFormat(fmt));
            xls.SetCellValue(30, 11, "manzana");
            xls.SetCellValue(30, 12, new TFormula("=L29"));

            fmt = xls.GetCellVisibleFormatDef(31, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));
            xls.SetCellValue(31, 4, new TFormula("=Metrics!D31"));

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));
            xls.SetCellValue(31, 5, "gallons");

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));
            xls.SetCellValue(31, 8, new TFormula("=VLOOKUP(1,$D$30:$E$31,2,FALSE)"));

            fmt = xls.GetCellVisibleFormatDef(31, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(31, 11, xls.AddFormat(fmt));
            xls.SetCellValue(31, 11, "liters");
            xls.SetCellValue(32, 11, "liter");

            fmt = xls.GetCellVisibleFormatDef(33, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, "Country currency");

            fmt = xls.GetCellVisibleFormatDef(33, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));
            xls.SetCellValue(33, 6, "Local / USD");
            xls.SetCellValue(33, 8, new TFormula("=VLOOKUP(1,$D$36:$E$49,2,FALSE)"));

            fmt = xls.GetCellVisibleFormatDef(33, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(33, 11, xls.AddFormat(fmt));
            xls.SetCellValue(33, 11, "gallons");

            fmt = xls.GetCellVisibleFormatDef(33, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(33, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));
            xls.SetCellValue(34, 11, "gallon");

            fmt = xls.GetCellVisibleFormatDef(34, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(34, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Font.Color = TExcelColor.Automatic;
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 12);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(35, 12, xls.AddFormat(fmt));
            xls.SetCellValue(35, 12, "Mexican pesos");

            fmt = xls.GetCellVisibleFormatDef(35, 14);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            xls.SetCellFormat(35, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));
            xls.SetCellValue(36, 3, "Bolivian boliviano");

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));
            xls.SetCellValue(36, 4, new TFormula("=Metrics!D34"));
            xls.SetCellValue(36, 5, "bolivianos");

            fmt = xls.GetCellVisibleFormatDef(36, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(36, 6, xls.AddFormat(fmt));
            xls.SetCellValue(36, 6, 6.91);
            xls.SetCellValue(36, 11, "bolivianos");
            xls.SetCellValue(36, 12, new TFormula("=(1/F36)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));
            xls.SetCellValue(37, 3, "Brazilian real");

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));
            xls.SetCellValue(37, 4, new TFormula("=Metrics!D35"));
            xls.SetCellValue(37, 5, "reales");

            fmt = xls.GetCellVisibleFormatDef(37, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(37, 6, xls.AddFormat(fmt));
            xls.SetCellValue(37, 6, 3.86);
            xls.SetCellValue(37, 11, "reales");
            xls.SetCellValue(37, 12, new TFormula("=(1/F37)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(38, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));
            xls.SetCellValue(38, 3, "Colombian peso");

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));
            xls.SetCellValue(38, 4, new TFormula("=Metrics!D36"));
            xls.SetCellValue(38, 5, "pesos");

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));
            xls.SetCellValue(38, 6, new TFormula("=Conversiones!E24"));

            fmt = xls.GetCellVisibleFormatDef(38, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            xls.SetCellFormat(38, 11, xls.AddFormat(fmt));
            xls.SetCellValue(38, 11, "colpesos");
            xls.SetCellValue(38, 12, new TFormula("=(1/F38)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(39, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));
            xls.SetCellValue(39, 3, "Costa rican colon");

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));
            xls.SetCellValue(39, 4, new TFormula("=Metrics!D37"));
            xls.SetCellValue(39, 5, "colones");

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));
            xls.SetCellValue(39, 6, 567.36);
            xls.SetCellValue(39, 11, "colones");
            xls.SetCellValue(39, 12, new TFormula("=(1/F39)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(40, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));
            xls.SetCellValue(40, 3, "Cuban peso");

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));
            xls.SetCellValue(40, 4, new TFormula("=Metrics!D38"));
            xls.SetCellValue(40, 5, "pesos");

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));
            xls.SetCellValue(40, 6, 1);

            fmt = xls.GetCellVisibleFormatDef(40, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            xls.SetCellFormat(40, 11, xls.AddFormat(fmt));
            xls.SetCellValue(40, 11, "cubpesos");
            xls.SetCellValue(40, 12, new TFormula("=(1/F40)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(41, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(41, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));
            xls.SetCellValue(41, 3, "Guatemalan quetzal");

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));
            xls.SetCellValue(41, 4, new TFormula("=Metrics!D39"));
            xls.SetCellValue(41, 5, "quetzales");

            fmt = xls.GetCellVisibleFormatDef(41, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(41, 6, xls.AddFormat(fmt));
            xls.SetCellValue(41, 6, 7.49);
            xls.SetCellValue(41, 11, "quetzales");
            xls.SetCellValue(41, 12, new TFormula("=(1/F41)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));
            xls.SetCellValue(42, 3, "Haitian gourde");

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));
            xls.SetCellValue(42, 4, new TFormula("=Metrics!D40"));
            xls.SetCellValue(42, 5, "gourdes");

            fmt = xls.GetCellVisibleFormatDef(42, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(42, 6, xls.AddFormat(fmt));
            xls.SetCellValue(42, 6, 66.76);
            xls.SetCellValue(42, 11, "gourdes");
            xls.SetCellValue(42, 12, new TFormula("=(1/F42)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(43, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(43, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));
            xls.SetCellValue(43, 3, "Honduran lempira");

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));
            xls.SetCellValue(43, 4, new TFormula("=Metrics!D41"));
            xls.SetCellValue(43, 5, "lempiras");

            fmt = xls.GetCellVisibleFormatDef(43, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(43, 6, xls.AddFormat(fmt));
            xls.SetCellValue(43, 6, new TFormula("=Conversiones!C24"));
            xls.SetCellValue(43, 11, "lempiras");
            xls.SetCellValue(43, 12, new TFormula("=(1/F43)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));
            xls.SetCellValue(44, 3, "Jamaican dollar");

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));
            xls.SetCellValue(44, 4, new TFormula("=Metrics!D42"));
            xls.SetCellValue(44, 5, "j. dollars");

            fmt = xls.GetCellVisibleFormatDef(44, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(44, 6, xls.AddFormat(fmt));
            xls.SetCellValue(44, 6, 134.95);
            xls.SetCellValue(44, 11, "j. dollars");
            xls.SetCellValue(44, 12, new TFormula("=(1/F44)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));
            xls.SetCellValue(45, 3, "Mexican peso");

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));
            xls.SetCellValue(45, 4, new TFormula("=Metrics!D43"));
            xls.SetCellValue(45, 5, "pesos");

            fmt = xls.GetCellVisibleFormatDef(45, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(45, 6, xls.AddFormat(fmt));
            xls.SetCellValue(45, 6, new TFormula("=Conversiones!F24"));

            fmt = xls.GetCellVisibleFormatDef(45, 11);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x66, 0x00);
            xls.SetCellFormat(45, 11, xls.AddFormat(fmt));
            xls.SetCellValue(45, 11, "mpesos");
            xls.SetCellValue(45, 12, new TFormula("=(1/F45)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(46, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(46, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));
            xls.SetCellValue(46, 3, "Nicaraguan cordoba");

            fmt = xls.GetCellVisibleFormatDef(46, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(46, 4, xls.AddFormat(fmt));
            xls.SetCellValue(46, 4, new TFormula("=Metrics!D44"));
            xls.SetCellValue(46, 5, "cordobas");

            fmt = xls.GetCellVisibleFormatDef(46, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(46, 6, xls.AddFormat(fmt));
            xls.SetCellValue(46, 6, 31.95);
            xls.SetCellValue(46, 11, "cordobas");
            xls.SetCellValue(46, 12, new TFormula("=(1/F46)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));
            xls.SetCellValue(47, 3, "Peruvian sol");

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));
            xls.SetCellValue(47, 4, new TFormula("=Metrics!D45"));
            xls.SetCellValue(47, 5, "soles");

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));
            xls.SetCellValue(47, 6, new TFormula("=Conversiones!D24"));
            xls.SetCellValue(47, 11, "soles");
            xls.SetCellValue(47, 12, new TFormula("=(1/F47)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(48, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(48, 3, xls.AddFormat(fmt));
            xls.SetCellValue(48, 3, "US dollar");

            fmt = xls.GetCellVisibleFormatDef(48, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(48, 4, xls.AddFormat(fmt));
            xls.SetCellValue(48, 4, new TFormula("=Metrics!D46"));
            xls.SetCellValue(48, 5, "dollars");

            fmt = xls.GetCellVisibleFormatDef(48, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(48, 6, xls.AddFormat(fmt));
            xls.SetCellValue(48, 6, 1);
            xls.SetCellValue(48, 11, "dollars");
            xls.SetCellValue(48, 12, new TFormula("=(1/F48)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(49, 3);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(49, 3, xls.AddFormat(fmt));
            xls.SetCellValue(49, 3, "Venezuelan bolivar");

            fmt = xls.GetCellVisibleFormatDef(49, 4);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(49, 4, xls.AddFormat(fmt));
            xls.SetCellValue(49, 4, new TFormula("=Metrics!D47"));
            xls.SetCellValue(49, 5, "bolivares");

            fmt = xls.GetCellVisibleFormatDef(49, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(49, 6, xls.AddFormat(fmt));
            xls.SetCellValue(49, 6, 207101);
            xls.SetCellValue(49, 11, "bolivares");
            xls.SetCellValue(49, 12, new TFormula("=(1/F49)*$F$45"));

            fmt = xls.GetCellVisibleFormatDef(50, 11);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xFF, 0xFF, 0x00);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(50, 11, xls.AddFormat(fmt));
            xls.SetCellValue(50, 11, "pesos");
            xls.SetCellValue(50, 12, new TFormula("=IF(D38=1,L38,IF(D40=1,L40,IF(D45=1,L45,0)))"));

            fmt = xls.GetCellVisibleFormatDef(52, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            xls.SetCellFormat(52, 6, xls.AddFormat(fmt));
            xls.SetCellValue(52, 6, "Fieldwork");

            fmt = xls.GetCellVisibleFormatDef(53, 6);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Family = 0;
            xls.SetCellFormat(53, 6, xls.AddFormat(fmt));
            xls.SetCellValue(53, 6, "Google exchange rate");

            //Cell selection and scroll position.
            xls.SelectCell(22, 4, false);
            xls.ScrollWindow(7, 1);

            //Standard Document Properties - Most are only for xlsx files. In xls files FlexCel will only change the Creation Date and Modified Date.
            xls.DocumentProperties.SetStandardProperty(TPropertyId.Author, "Mary Kate");

            //You will normally not set LastSavedBy, since this is a new file.
            //If you don't set it, FlexCel will use the creator instead.
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.LastSavedBy, "SAPTARSHI MALLICK");


            //You will normally not set CreateDateTime, since this is a new file and FlexCel will automatically use the current datetime.
            //But if you are editing a file and want to preserve the original creation date, you need to either set PreserveCreationDate to true:
            //    xls.DocumentProperties.PreserveCreationDate = true;
            //Or you can hardcode a creating date by setting it in UTC time, ISO8601 format:
            //    xls.DocumentProperties.SetStandardProperty(TPropertyId.CreateTimeDate, "2015-01-08T03:31:31Z");


            xls.Recalc(true);
            
            var s = xls.GetCellValue(23, 9);
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
