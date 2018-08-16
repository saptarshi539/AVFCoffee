using FlexCel.Core;

namespace CoffeeInfrastructure.Flexcel
{
    public class Budget_M_Obra
    {
        public void BudgetMObra(ExcelFile xls)
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

            xls.ActiveSheet = 30;    //Set the sheet we are working in.

            //Global Workbook Options
            xls.OptionsAutoCompressPictures = false;
            xls.OptionsMultithreadRecalc = 0;

            //Sheet Options
            xls.SheetName = "Budget_M Obra";
            xls.SheetZoom = 69;

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
            xls.PrintToFit = true;
            xls.PrintScale = 48;
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
            xls.DefaultColWidth = 2784;

            xls.SetColWidth(1, 1, 12416);    //(47.75 + 0.75) * 256

            xls.SetColWidth(2, 2, 3328);    //(12.25 + 0.75) * 256

            xls.SetColWidth(3, 3, 2784);    //(10.13 + 0.75) * 256

            xls.SetColWidth(4, 4, 2816);    //(10.25 + 0.75) * 256

            xls.SetColWidth(5, 7, 3040);    //(11.13 + 0.75) * 256

            xls.SetColWidth(8, 10, 2816);    //(10.25 + 0.75) * 256

            xls.SetColWidth(11, 11, 14976);    //(57.75 + 0.75) * 256

            xls.SetColWidth(12, 16384, 2784);    //(10.13 + 0.75) * 256
            xls.DefaultRowHeight = 315;

            xls.SetRowHeight(1, 375);    //18.75 * 20
            xls.SetRowHeight(2, 375);    //18.75 * 20
            xls.SetRowHeight(84, 1260);    //63.00 * 20
            xls.SetRowHeight(86, 945);    //47.25 * 20
            xls.SetRowHeight(88, 945);    //47.25 * 20
            xls.SetRowHeight(90, 1260);    //63.00 * 20

            //Set the cell values
            TFlxFormat fmt;
            fmt = xls.GetCellVisibleFormatDef(1, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(1, 1, xls.AddFormat(fmt));
            xls.SetCellValue(1, 1, "En días por favor describa cuanto tiempo se invierte en las siguientes actividades"
            + " para una hectarea de café:");

            fmt = xls.GetCellVisibleFormatDef(1, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(1, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(1, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 4);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(1, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 5);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(1, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(1, 6);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(1, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 1);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 1, xls.AddFormat(fmt));
            xls.SetCellValue(2, 1, "Tenga en cuenta que la intensidad de días puede cambiar de año a año");

            fmt = xls.GetCellVisibleFormatDef(2, 2);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 3);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 4);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 5);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(2, 6);
            fmt.Font.Size20 = 280;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(2, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(3, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(3, 1, xls.AddFormat(fmt));
            xls.SetCellValue(3, 1, "Mano de Obra");

            fmt = xls.GetCellVisibleFormatDef(4, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(4, 1, xls.AddFormat(fmt));
            xls.SetCellValue(4, 1, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(5, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(5, 1, xls.AddFormat(fmt));
            xls.SetCellValue(5, 1, "Mano de obra para el germinador");

            fmt = xls.GetCellVisibleFormatDef(5, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 2, xls.AddFormat(fmt));
            xls.SetCellValue(5, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(5, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 3, xls.AddFormat(fmt));
            xls.SetCellValue(5, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(5, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 4, xls.AddFormat(fmt));
            xls.SetCellValue(5, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(5, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 5, xls.AddFormat(fmt));
            xls.SetCellValue(5, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(5, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 6, xls.AddFormat(fmt));
            xls.SetCellValue(5, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(5, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 7, xls.AddFormat(fmt));
            xls.SetCellValue(5, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(5, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 8, xls.AddFormat(fmt));
            xls.SetCellValue(5, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(5, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 9, xls.AddFormat(fmt));
            xls.SetCellValue(5, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(5, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 10, xls.AddFormat(fmt));
            xls.SetCellValue(5, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(5, 11);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(5, 11, xls.AddFormat(fmt));
            xls.SetCellValue(5, 11, "Notes");

            fmt = xls.GetCellVisibleFormatDef(6, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 1, xls.AddFormat(fmt));
            xls.SetCellValue(6, 1, "Recolección de semillas");

            fmt = xls.GetCellVisibleFormatDef(6, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 2, xls.AddFormat(fmt));
            xls.SetCellValue(6, 2, new TFormula("='Inputs TOT advanced'!F7"));

            fmt = xls.GetCellVisibleFormatDef(6, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(6, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(6, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(6, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 1, xls.AddFormat(fmt));
            xls.SetCellValue(7, 1, "Selección de semillas");

            fmt = xls.GetCellVisibleFormatDef(7, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 2, xls.AddFormat(fmt));
            xls.SetCellValue(7, 2, new TFormula("='Inputs TOT advanced'!F8"));

            fmt = xls.GetCellVisibleFormatDef(7, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(7, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(7, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(7, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 1, xls.AddFormat(fmt));
            xls.SetCellValue(8, 1, "Construcción Semillero");

            fmt = xls.GetCellVisibleFormatDef(8, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 2, xls.AddFormat(fmt));
            xls.SetCellValue(8, 2, new TFormula("='Inputs TOT advanced'!F9"));

            fmt = xls.GetCellVisibleFormatDef(8, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(8, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(8, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(8, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 1, xls.AddFormat(fmt));
            xls.SetCellValue(9, 1, "Sostenimiento semillero - Riego");

            fmt = xls.GetCellVisibleFormatDef(9, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 2, xls.AddFormat(fmt));
            xls.SetCellValue(9, 2, new TFormula("='Inputs TOT advanced'!F10"));

            fmt = xls.GetCellVisibleFormatDef(9, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(9, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(9, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(9, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 1, xls.AddFormat(fmt));
            xls.SetCellValue(10, 1, "Otros");

            fmt = xls.GetCellVisibleFormatDef(10, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 2, xls.AddFormat(fmt));
            xls.SetCellValue(10, 2, new TFormula("='Inputs TOT advanced'!F11"));

            fmt = xls.GetCellVisibleFormatDef(10, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(10, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(10, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(10, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(11, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(11, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(11, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(12, 1, xls.AddFormat(fmt));
            xls.SetCellValue(12, 1, "Mano de obra para el vivero");

            fmt = xls.GetCellVisibleFormatDef(12, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 2, xls.AddFormat(fmt));
            xls.SetCellValue(12, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(12, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 3, xls.AddFormat(fmt));
            xls.SetCellValue(12, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(12, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 4, xls.AddFormat(fmt));
            xls.SetCellValue(12, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(12, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 5, xls.AddFormat(fmt));
            xls.SetCellValue(12, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(12, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 6, xls.AddFormat(fmt));
            xls.SetCellValue(12, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(12, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 7, xls.AddFormat(fmt));
            xls.SetCellValue(12, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(12, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 8, xls.AddFormat(fmt));
            xls.SetCellValue(12, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(12, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 9, xls.AddFormat(fmt));
            xls.SetCellValue(12, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(12, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(12, 10, xls.AddFormat(fmt));
            xls.SetCellValue(12, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(12, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(12, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(12, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 1, xls.AddFormat(fmt));
            xls.SetCellValue(13, 1, "Construcción del vivero");

            fmt = xls.GetCellVisibleFormatDef(13, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 2, xls.AddFormat(fmt));
            xls.SetCellValue(13, 2, new TFormula("='Inputs TOT advanced'!F13"));

            fmt = xls.GetCellVisibleFormatDef(13, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(13, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(13, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(13, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 1, xls.AddFormat(fmt));
            xls.SetCellValue(14, 1, "Jalada y arrancada de la tierra para el vivero");

            fmt = xls.GetCellVisibleFormatDef(14, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 2, xls.AddFormat(fmt));
            xls.SetCellValue(14, 2, new TFormula("='Inputs TOT advanced'!F14"));

            fmt = xls.GetCellVisibleFormatDef(14, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(14, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(14, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(14, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(15, 1, xls.AddFormat(fmt));
            xls.SetCellValue(15, 1, "Limpia del vivero");

            fmt = xls.GetCellVisibleFormatDef(15, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 2, xls.AddFormat(fmt));
            xls.SetCellValue(15, 2, new TFormula("='Inputs TOT advanced'!F15"));

            fmt = xls.GetCellVisibleFormatDef(15, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(15, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(15, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(15, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 1, xls.AddFormat(fmt));
            xls.SetCellValue(16, 1, "Preparacion de tierra con abono organico para llenado");

            fmt = xls.GetCellVisibleFormatDef(16, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 2, xls.AddFormat(fmt));
            xls.SetCellValue(16, 2, new TFormula("='Inputs TOT advanced'!F16"));

            fmt = xls.GetCellVisibleFormatDef(16, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(16, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(16, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(16, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 1, xls.AddFormat(fmt));
            xls.SetCellValue(17, 1, "Llenada y encerrada de bolsas");

            fmt = xls.GetCellVisibleFormatDef(17, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 2, xls.AddFormat(fmt));
            xls.SetCellValue(17, 2, new TFormula("='Inputs TOT advanced'!F17"));

            fmt = xls.GetCellVisibleFormatDef(17, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(17, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(17, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(17, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 1, xls.AddFormat(fmt));
            xls.SetCellValue(18, 1, "Siembra de maripositas");

            fmt = xls.GetCellVisibleFormatDef(18, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 2, xls.AddFormat(fmt));
            xls.SetCellValue(18, 2, new TFormula("='Inputs TOT advanced'!F18"));

            fmt = xls.GetCellVisibleFormatDef(18, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(18, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(18, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(18, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 1, xls.AddFormat(fmt));
            xls.SetCellValue(19, 1, "Riego");

            fmt = xls.GetCellVisibleFormatDef(19, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 2, xls.AddFormat(fmt));
            xls.SetCellValue(19, 2, new TFormula("='Inputs TOT advanced'!F19"));

            fmt = xls.GetCellVisibleFormatDef(19, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(19, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(19, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(19, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 1, xls.AddFormat(fmt));
            xls.SetCellValue(20, 1, "Aplicación de foliares");

            fmt = xls.GetCellVisibleFormatDef(20, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 2, xls.AddFormat(fmt));
            xls.SetCellValue(20, 2, new TFormula("='Inputs TOT advanced'!F20"));

            fmt = xls.GetCellVisibleFormatDef(20, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(20, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(20, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(20, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 1, xls.AddFormat(fmt));
            xls.SetCellValue(21, 1, "Resiembras");

            fmt = xls.GetCellVisibleFormatDef(21, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 2, xls.AddFormat(fmt));
            xls.SetCellValue(21, 2, new TFormula("='Inputs TOT advanced'!F21"));

            fmt = xls.GetCellVisibleFormatDef(21, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(21, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(21, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(21, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 1, xls.AddFormat(fmt));
            xls.SetCellValue(22, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(22, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 2, xls.AddFormat(fmt));
            xls.SetCellValue(22, 2, new TFormula("='Inputs TOT advanced'!F22"));

            fmt = xls.GetCellVisibleFormatDef(22, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(22, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(22, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Accent6, 0.399975585192419);
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(22, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(23, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(23, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(23, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(24, 1, xls.AddFormat(fmt));
            xls.SetCellValue(24, 1, "Mano de obra preparación terreno para renovacion");

            fmt = xls.GetCellVisibleFormatDef(24, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 2, xls.AddFormat(fmt));
            xls.SetCellValue(24, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(24, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 3, xls.AddFormat(fmt));
            xls.SetCellValue(24, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(24, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 4, xls.AddFormat(fmt));
            xls.SetCellValue(24, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(24, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 5, xls.AddFormat(fmt));
            xls.SetCellValue(24, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(24, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 6, xls.AddFormat(fmt));
            xls.SetCellValue(24, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(24, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 7, xls.AddFormat(fmt));
            xls.SetCellValue(24, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(24, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 8, xls.AddFormat(fmt));
            xls.SetCellValue(24, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(24, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 9, xls.AddFormat(fmt));
            xls.SetCellValue(24, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(24, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(24, 10, xls.AddFormat(fmt));
            xls.SetCellValue(24, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(24, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 13);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 14);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 15);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 16);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 17);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 18);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(24, 19);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(24, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 1, xls.AddFormat(fmt));
            xls.SetCellValue(25, 1, "Limpia del terreno");

            fmt = xls.GetCellVisibleFormatDef(25, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 2, xls.AddFormat(fmt));
            xls.SetCellValue(25, 2, new TFormula("='Inputs TOT advanced'!F24"));

            fmt = xls.GetCellVisibleFormatDef(25, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(25, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(25, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(25, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 1, xls.AddFormat(fmt));
            xls.SetCellValue(26, 1, "Corte de arboles de café viejos u otros maderables");

            fmt = xls.GetCellVisibleFormatDef(26, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 2, xls.AddFormat(fmt));
            xls.SetCellValue(26, 2, new TFormula("='Inputs TOT advanced'!F25"));

            fmt = xls.GetCellVisibleFormatDef(26, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(26, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(26, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(26, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 1, xls.AddFormat(fmt));
            xls.SetCellValue(27, 1, "Recolección y acopio de madera de café");

            fmt = xls.GetCellVisibleFormatDef(27, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 2, xls.AddFormat(fmt));
            xls.SetCellValue(27, 2, new TFormula("='Inputs TOT advanced'!F26"));

            fmt = xls.GetCellVisibleFormatDef(27, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(27, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(27, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(27, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 1, xls.AddFormat(fmt));
            xls.SetCellValue(28, 1, "Pique de la madera y/o elaboración de estacas");

            fmt = xls.GetCellVisibleFormatDef(28, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 2, xls.AddFormat(fmt));
            xls.SetCellValue(28, 2, new TFormula("='Inputs TOT advanced'!F27"));

            fmt = xls.GetCellVisibleFormatDef(28, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(28, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(28, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(28, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 1, xls.AddFormat(fmt));
            xls.SetCellValue(29, 1, "Trazado Café");

            fmt = xls.GetCellVisibleFormatDef(29, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 2, xls.AddFormat(fmt));
            xls.SetCellValue(29, 2, new TFormula("='Inputs TOT advanced'!F28"));

            fmt = xls.GetCellVisibleFormatDef(29, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(29, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(29, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(29, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 1, xls.AddFormat(fmt));
            xls.SetCellValue(30, 1, "Ahoyado para la siembra");

            fmt = xls.GetCellVisibleFormatDef(30, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 2, xls.AddFormat(fmt));
            xls.SetCellValue(30, 2, new TFormula("='Inputs TOT advanced'!F29"));

            fmt = xls.GetCellVisibleFormatDef(30, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(30, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(30, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(30, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 1, xls.AddFormat(fmt));
            xls.SetCellValue(31, 1, "Llevada de las plantas del vivero (en la finca) al terreno ");

            fmt = xls.GetCellVisibleFormatDef(31, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 2, xls.AddFormat(fmt));
            xls.SetCellValue(31, 2, new TFormula("='Inputs TOT advanced'!F30"));

            fmt = xls.GetCellVisibleFormatDef(31, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(31, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(31, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(31, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 1, xls.AddFormat(fmt));
            xls.SetCellValue(32, 1, "Siembra de plantones (o plantulas)");

            fmt = xls.GetCellVisibleFormatDef(32, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 2, xls.AddFormat(fmt));
            xls.SetCellValue(32, 2, new TFormula("='Inputs TOT advanced'!F31"));

            fmt = xls.GetCellVisibleFormatDef(32, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(32, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(32, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(32, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 1, xls.AddFormat(fmt));
            xls.SetCellValue(33, 1, "Adecuación de los arboles de sombrio");

            fmt = xls.GetCellVisibleFormatDef(33, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 2, xls.AddFormat(fmt));
            xls.SetCellValue(33, 2, new TFormula("='Inputs TOT advanced'!F32"));

            fmt = xls.GetCellVisibleFormatDef(33, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(33, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(33, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(33, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(34, 1, xls.AddFormat(fmt));
            xls.SetCellValue(34, 1, "Preparación de abonos orgánicos");

            fmt = xls.GetCellVisibleFormatDef(34, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(34, 2, xls.AddFormat(fmt));
            xls.SetCellValue(34, 2, new TFormula("='Inputs TOT advanced'!F33"));

            fmt = xls.GetCellVisibleFormatDef(34, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            fmt.WrapText = true;
            xls.SetCellFormat(34, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(34, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(34, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 1, xls.AddFormat(fmt));
            xls.SetCellValue(35, 1, "Otros");

            fmt = xls.GetCellVisibleFormatDef(35, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 2, xls.AddFormat(fmt));
            xls.SetCellValue(35, 2, new TFormula("='Inputs TOT advanced'!F34"));

            fmt = xls.GetCellVisibleFormatDef(35, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(35, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(35, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(35, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(36, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(36, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(36, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(37, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(37, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(37, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(38, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(38, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(38, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(39, 1, xls.AddFormat(fmt));
            xls.SetCellValue(39, 1, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(39, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(39, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(39, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(39, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(40, 1, xls.AddFormat(fmt));
            xls.SetCellValue(40, 1, "Mano de obra para la plantilla o levante ");

            fmt = xls.GetCellVisibleFormatDef(40, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 2, xls.AddFormat(fmt));
            xls.SetCellValue(40, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(40, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 3, xls.AddFormat(fmt));
            xls.SetCellValue(40, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(40, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 4, xls.AddFormat(fmt));
            xls.SetCellValue(40, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(40, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 5, xls.AddFormat(fmt));
            xls.SetCellValue(40, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(40, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 6, xls.AddFormat(fmt));
            xls.SetCellValue(40, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(40, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 7, xls.AddFormat(fmt));
            xls.SetCellValue(40, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(40, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 8, xls.AddFormat(fmt));
            xls.SetCellValue(40, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(40, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 9, xls.AddFormat(fmt));
            xls.SetCellValue(40, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(40, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(40, 10, xls.AddFormat(fmt));
            xls.SetCellValue(40, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(40, 11);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 12);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 13);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 14);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 15);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 16);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 17);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 18);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(40, 19);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            xls.SetCellFormat(40, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 1, xls.AddFormat(fmt));
            xls.SetCellValue(41, 1, "Desyerbe periodico ");

            fmt = xls.GetCellVisibleFormatDef(41, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 3, xls.AddFormat(fmt));
            xls.SetCellValue(41, 3, new TFormula("='Inputs TOT advanced'!F36"));

            fmt = xls.GetCellVisibleFormatDef(41, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(41, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(41, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(41, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 1, xls.AddFormat(fmt));
            xls.SetCellValue(42, 1, "Aplicación de abonos orgánicos para levante");

            fmt = xls.GetCellVisibleFormatDef(42, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 3, xls.AddFormat(fmt));
            xls.SetCellValue(42, 3, new TFormula("='Inputs TOT advanced'!F37"));

            fmt = xls.GetCellVisibleFormatDef(42, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(42, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(42, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(42, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 1, xls.AddFormat(fmt));
            xls.SetCellValue(43, 1, "Aplicación de abonos químicos para levante");

            fmt = xls.GetCellVisibleFormatDef(43, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 3, xls.AddFormat(fmt));
            xls.SetCellValue(43, 3, new TFormula("='Inputs TOT advanced'!F38"));

            fmt = xls.GetCellVisibleFormatDef(43, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(43, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(43, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(43, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 1, xls.AddFormat(fmt));
            xls.SetCellValue(44, 1, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(44, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 3, xls.AddFormat(fmt));
            xls.SetCellValue(44, 3, new TFormula("='Inputs TOT advanced'!F39"));

            fmt = xls.GetCellVisibleFormatDef(44, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(44, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(44, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(44, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 1, xls.AddFormat(fmt));
            xls.SetCellValue(45, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(45, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 3, xls.AddFormat(fmt));
            xls.SetCellValue(45, 3, new TFormula("='Inputs TOT advanced'!F40"));

            fmt = xls.GetCellVisibleFormatDef(45, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(45, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(45, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(45, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(46, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(46, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(46, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(47, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(47, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(47, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(48, 1, xls.AddFormat(fmt));
            xls.SetCellValue(48, 1, "Año 2-8");

            fmt = xls.GetCellVisibleFormatDef(48, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(48, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(48, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(48, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(49, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(49, 1, xls.AddFormat(fmt));
            xls.SetCellValue(49, 1, "Valor mano de obra para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(49, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 2, xls.AddFormat(fmt));
            xls.SetCellValue(49, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(49, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 3, xls.AddFormat(fmt));
            xls.SetCellValue(49, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(49, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 4, xls.AddFormat(fmt));
            xls.SetCellValue(49, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(49, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 5, xls.AddFormat(fmt));
            xls.SetCellValue(49, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(49, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 6, xls.AddFormat(fmt));
            xls.SetCellValue(49, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(49, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 7, xls.AddFormat(fmt));
            xls.SetCellValue(49, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(49, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 8, xls.AddFormat(fmt));
            xls.SetCellValue(49, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(49, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 9, xls.AddFormat(fmt));
            xls.SetCellValue(49, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(49, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(49, 10, xls.AddFormat(fmt));
            xls.SetCellValue(49, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(50, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(50, 1, xls.AddFormat(fmt));
            xls.SetCellValue(50, 1, "Desyerbe para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(50, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(50, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(50, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(50, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(50, 4, xls.AddFormat(fmt));
            xls.SetCellValue(50, 4, new TFormula("='Inputs TOT advanced'!F63"));

            fmt = xls.GetCellVisibleFormatDef(50, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(50, 5, xls.AddFormat(fmt));
            xls.SetCellValue(50, 5, new TFormula("='Inputs TOT advanced'!F63"));

            fmt = xls.GetCellVisibleFormatDef(50, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(50, 6, xls.AddFormat(fmt));
            xls.SetCellValue(50, 6, new TFormula("='Inputs TOT advanced'!F90"));

            fmt = xls.GetCellVisibleFormatDef(50, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(50, 7, xls.AddFormat(fmt));
            xls.SetCellValue(50, 7, new TFormula("='Inputs TOT advanced'!F90"));

            fmt = xls.GetCellVisibleFormatDef(50, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(50, 8, xls.AddFormat(fmt));
            xls.SetCellValue(50, 8, new TFormula("='Inputs TOT advanced'!F90"));

            fmt = xls.GetCellVisibleFormatDef(50, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(50, 9, xls.AddFormat(fmt));
            xls.SetCellValue(50, 9, new TFormula("='Inputs TOT advanced'!F117"));

            fmt = xls.GetCellVisibleFormatDef(50, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(50, 10, xls.AddFormat(fmt));
            xls.SetCellValue(50, 10, new TFormula("='Inputs TOT advanced'!F117"));

            fmt = xls.GetCellVisibleFormatDef(51, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(51, 1, xls.AddFormat(fmt));
            xls.SetCellValue(51, 1, "Desyerbe quimico para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(51, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(51, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(51, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(51, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(51, 4, xls.AddFormat(fmt));
            xls.SetCellValue(51, 4, new TFormula("='Inputs TOT advanced'!F64"));

            fmt = xls.GetCellVisibleFormatDef(51, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(51, 5, xls.AddFormat(fmt));
            xls.SetCellValue(51, 5, new TFormula("='Inputs TOT advanced'!F64"));

            fmt = xls.GetCellVisibleFormatDef(51, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(51, 6, xls.AddFormat(fmt));
            xls.SetCellValue(51, 6, new TFormula("='Inputs TOT advanced'!F91"));

            fmt = xls.GetCellVisibleFormatDef(51, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(51, 7, xls.AddFormat(fmt));
            xls.SetCellValue(51, 7, new TFormula("='Inputs TOT advanced'!F91"));

            fmt = xls.GetCellVisibleFormatDef(51, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(51, 8, xls.AddFormat(fmt));
            xls.SetCellValue(51, 8, new TFormula("='Inputs TOT advanced'!F91"));

            fmt = xls.GetCellVisibleFormatDef(51, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(51, 9, xls.AddFormat(fmt));
            xls.SetCellValue(51, 9, new TFormula("='Inputs TOT advanced'!F118"));

            fmt = xls.GetCellVisibleFormatDef(51, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(51, 10, xls.AddFormat(fmt));
            xls.SetCellValue(51, 10, new TFormula("='Inputs TOT advanced'!F118"));

            fmt = xls.GetCellVisibleFormatDef(52, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(52, 1, xls.AddFormat(fmt));
            xls.SetCellValue(52, 1, "Aplicación de abonos orgánicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(52, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(52, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(52, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(52, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(52, 4, xls.AddFormat(fmt));
            xls.SetCellValue(52, 4, new TFormula("='Inputs TOT advanced'!F65"));

            fmt = xls.GetCellVisibleFormatDef(52, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(52, 5, xls.AddFormat(fmt));
            xls.SetCellValue(52, 5, new TFormula("='Inputs TOT advanced'!F65"));

            fmt = xls.GetCellVisibleFormatDef(52, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(52, 6, xls.AddFormat(fmt));
            xls.SetCellValue(52, 6, new TFormula("='Inputs TOT advanced'!F92"));

            fmt = xls.GetCellVisibleFormatDef(52, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(52, 7, xls.AddFormat(fmt));
            xls.SetCellValue(52, 7, new TFormula("='Inputs TOT advanced'!F92"));

            fmt = xls.GetCellVisibleFormatDef(52, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(52, 8, xls.AddFormat(fmt));
            xls.SetCellValue(52, 8, new TFormula("='Inputs TOT advanced'!F92"));

            fmt = xls.GetCellVisibleFormatDef(52, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(52, 9, xls.AddFormat(fmt));
            xls.SetCellValue(52, 9, new TFormula("='Inputs TOT advanced'!F119"));

            fmt = xls.GetCellVisibleFormatDef(52, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(52, 10, xls.AddFormat(fmt));
            xls.SetCellValue(52, 10, new TFormula("='Inputs TOT advanced'!F119"));

            fmt = xls.GetCellVisibleFormatDef(53, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(53, 1, xls.AddFormat(fmt));
            xls.SetCellValue(53, 1, "Aplicación de abonos químicos para mantenimiento");

            fmt = xls.GetCellVisibleFormatDef(53, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(53, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(53, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(53, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(53, 4, xls.AddFormat(fmt));
            xls.SetCellValue(53, 4, new TFormula("='Inputs TOT advanced'!F66"));

            fmt = xls.GetCellVisibleFormatDef(53, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(53, 5, xls.AddFormat(fmt));
            xls.SetCellValue(53, 5, new TFormula("='Inputs TOT advanced'!F66"));

            fmt = xls.GetCellVisibleFormatDef(53, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(53, 6, xls.AddFormat(fmt));
            xls.SetCellValue(53, 6, new TFormula("='Inputs TOT advanced'!F93"));

            fmt = xls.GetCellVisibleFormatDef(53, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(53, 7, xls.AddFormat(fmt));
            xls.SetCellValue(53, 7, new TFormula("='Inputs TOT advanced'!F93"));

            fmt = xls.GetCellVisibleFormatDef(53, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(53, 8, xls.AddFormat(fmt));
            xls.SetCellValue(53, 8, new TFormula("='Inputs TOT advanced'!F93"));

            fmt = xls.GetCellVisibleFormatDef(53, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(53, 9, xls.AddFormat(fmt));
            xls.SetCellValue(53, 9, new TFormula("='Inputs TOT advanced'!F120"));

            fmt = xls.GetCellVisibleFormatDef(53, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(53, 10, xls.AddFormat(fmt));
            xls.SetCellValue(53, 10, new TFormula("='Inputs TOT advanced'!F120"));

            fmt = xls.GetCellVisibleFormatDef(54, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(54, 1, xls.AddFormat(fmt));
            xls.SetCellValue(54, 1, "Aplicación de foliares para fertilización y control roya");

            fmt = xls.GetCellVisibleFormatDef(54, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(54, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(54, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(54, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(54, 4, xls.AddFormat(fmt));
            xls.SetCellValue(54, 4, new TFormula("='Inputs TOT advanced'!F67"));

            fmt = xls.GetCellVisibleFormatDef(54, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(54, 5, xls.AddFormat(fmt));
            xls.SetCellValue(54, 5, new TFormula("='Inputs TOT advanced'!F67"));

            fmt = xls.GetCellVisibleFormatDef(54, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(54, 6, xls.AddFormat(fmt));
            xls.SetCellValue(54, 6, new TFormula("='Inputs TOT advanced'!F94"));

            fmt = xls.GetCellVisibleFormatDef(54, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(54, 7, xls.AddFormat(fmt));
            xls.SetCellValue(54, 7, new TFormula("='Inputs TOT advanced'!F94"));

            fmt = xls.GetCellVisibleFormatDef(54, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(54, 8, xls.AddFormat(fmt));
            xls.SetCellValue(54, 8, new TFormula("='Inputs TOT advanced'!F94"));

            fmt = xls.GetCellVisibleFormatDef(54, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(54, 9, xls.AddFormat(fmt));
            xls.SetCellValue(54, 9, new TFormula("='Inputs TOT advanced'!F121"));

            fmt = xls.GetCellVisibleFormatDef(54, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(54, 10, xls.AddFormat(fmt));
            xls.SetCellValue(54, 10, new TFormula("='Inputs TOT advanced'!F121"));

            fmt = xls.GetCellVisibleFormatDef(55, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(55, 1, xls.AddFormat(fmt));
            xls.SetCellValue(55, 1, "Construcción de barreras vivas (rompe-vientos)");

            fmt = xls.GetCellVisibleFormatDef(55, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(55, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(55, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(55, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(55, 4, xls.AddFormat(fmt));
            xls.SetCellValue(55, 4, new TFormula("='Inputs TOT advanced'!F68"));

            fmt = xls.GetCellVisibleFormatDef(55, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(55, 5, xls.AddFormat(fmt));
            xls.SetCellValue(55, 5, new TFormula("='Inputs TOT advanced'!F68"));

            fmt = xls.GetCellVisibleFormatDef(55, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(55, 6, xls.AddFormat(fmt));
            xls.SetCellValue(55, 6, new TFormula("='Inputs TOT advanced'!F95"));

            fmt = xls.GetCellVisibleFormatDef(55, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(55, 7, xls.AddFormat(fmt));
            xls.SetCellValue(55, 7, new TFormula("='Inputs TOT advanced'!F95"));

            fmt = xls.GetCellVisibleFormatDef(55, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(55, 8, xls.AddFormat(fmt));
            xls.SetCellValue(55, 8, new TFormula("='Inputs TOT advanced'!F95"));

            fmt = xls.GetCellVisibleFormatDef(55, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(55, 9, xls.AddFormat(fmt));
            xls.SetCellValue(55, 9, new TFormula("='Inputs TOT advanced'!F122"));

            fmt = xls.GetCellVisibleFormatDef(55, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(55, 10, xls.AddFormat(fmt));
            xls.SetCellValue(55, 10, new TFormula("='Inputs TOT advanced'!F122"));

            fmt = xls.GetCellVisibleFormatDef(56, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(56, 1, xls.AddFormat(fmt));
            xls.SetCellValue(56, 1, "Podas de árboles de sombra (sostenimiento)");

            fmt = xls.GetCellVisibleFormatDef(56, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(56, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(56, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(56, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(56, 4, xls.AddFormat(fmt));
            xls.SetCellValue(56, 4, new TFormula("='Inputs TOT advanced'!F69"));

            fmt = xls.GetCellVisibleFormatDef(56, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(56, 5, xls.AddFormat(fmt));
            xls.SetCellValue(56, 5, new TFormula("='Inputs TOT advanced'!F69"));

            fmt = xls.GetCellVisibleFormatDef(56, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(56, 6, xls.AddFormat(fmt));
            xls.SetCellValue(56, 6, new TFormula("='Inputs TOT advanced'!F96"));

            fmt = xls.GetCellVisibleFormatDef(56, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(56, 7, xls.AddFormat(fmt));
            xls.SetCellValue(56, 7, new TFormula("='Inputs TOT advanced'!F96"));

            fmt = xls.GetCellVisibleFormatDef(56, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(56, 8, xls.AddFormat(fmt));
            xls.SetCellValue(56, 8, new TFormula("='Inputs TOT advanced'!F96"));

            fmt = xls.GetCellVisibleFormatDef(56, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(56, 9, xls.AddFormat(fmt));
            xls.SetCellValue(56, 9, new TFormula("='Inputs TOT advanced'!F123"));

            fmt = xls.GetCellVisibleFormatDef(56, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(56, 10, xls.AddFormat(fmt));
            xls.SetCellValue(56, 10, new TFormula("='Inputs TOT advanced'!F123"));

            fmt = xls.GetCellVisibleFormatDef(57, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(57, 1, xls.AddFormat(fmt));
            xls.SetCellValue(57, 1, "Control de Broca (re-re, repela, fumigaciones)");

            fmt = xls.GetCellVisibleFormatDef(57, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(57, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(57, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(57, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(57, 4, xls.AddFormat(fmt));
            xls.SetCellValue(57, 4, new TFormula("='Inputs TOT advanced'!F70"));

            fmt = xls.GetCellVisibleFormatDef(57, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(57, 5, xls.AddFormat(fmt));
            xls.SetCellValue(57, 5, new TFormula("='Inputs TOT advanced'!F70"));

            fmt = xls.GetCellVisibleFormatDef(57, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(57, 6, xls.AddFormat(fmt));
            xls.SetCellValue(57, 6, new TFormula("='Inputs TOT advanced'!F97"));

            fmt = xls.GetCellVisibleFormatDef(57, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(57, 7, xls.AddFormat(fmt));
            xls.SetCellValue(57, 7, new TFormula("='Inputs TOT advanced'!F97"));

            fmt = xls.GetCellVisibleFormatDef(57, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(57, 8, xls.AddFormat(fmt));
            xls.SetCellValue(57, 8, new TFormula("='Inputs TOT advanced'!F97"));

            fmt = xls.GetCellVisibleFormatDef(57, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(57, 9, xls.AddFormat(fmt));
            xls.SetCellValue(57, 9, new TFormula("='Inputs TOT advanced'!F124"));

            fmt = xls.GetCellVisibleFormatDef(57, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(57, 10, xls.AddFormat(fmt));
            xls.SetCellValue(57, 10, new TFormula("='Inputs TOT advanced'!F124"));

            fmt = xls.GetCellVisibleFormatDef(58, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(58, 1, xls.AddFormat(fmt));
            xls.SetCellValue(58, 1, "Manejo de tejido (desrrame o podas del café)");

            fmt = xls.GetCellVisibleFormatDef(58, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(58, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(58, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(58, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(58, 4, xls.AddFormat(fmt));
            xls.SetCellValue(58, 4, new TFormula("='Inputs TOT advanced'!F71"));

            fmt = xls.GetCellVisibleFormatDef(58, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(58, 5, xls.AddFormat(fmt));
            xls.SetCellValue(58, 5, new TFormula("='Inputs TOT advanced'!F71"));

            fmt = xls.GetCellVisibleFormatDef(58, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(58, 6, xls.AddFormat(fmt));
            xls.SetCellValue(58, 6, new TFormula("='Inputs TOT advanced'!F98"));

            fmt = xls.GetCellVisibleFormatDef(58, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(58, 7, xls.AddFormat(fmt));
            xls.SetCellValue(58, 7, new TFormula("='Inputs TOT advanced'!F98"));

            fmt = xls.GetCellVisibleFormatDef(58, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(58, 8, xls.AddFormat(fmt));
            xls.SetCellValue(58, 8, new TFormula("='Inputs TOT advanced'!F98"));

            fmt = xls.GetCellVisibleFormatDef(58, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(58, 9, xls.AddFormat(fmt));
            xls.SetCellValue(58, 9, new TFormula("='Inputs TOT advanced'!F125"));

            fmt = xls.GetCellVisibleFormatDef(58, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(58, 10, xls.AddFormat(fmt));
            xls.SetCellValue(58, 10, new TFormula("='Inputs TOT advanced'!F125"));

            fmt = xls.GetCellVisibleFormatDef(59, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(59, 1, xls.AddFormat(fmt));
            xls.SetCellValue(59, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(59, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(59, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(59, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(59, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(59, 4, xls.AddFormat(fmt));
            xls.SetCellValue(59, 4, new TFormula("='Inputs TOT advanced'!F72"));

            fmt = xls.GetCellVisibleFormatDef(59, 5);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(59, 5, xls.AddFormat(fmt));
            xls.SetCellValue(59, 5, new TFormula("='Inputs TOT advanced'!F72"));

            fmt = xls.GetCellVisibleFormatDef(59, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(59, 6, xls.AddFormat(fmt));
            xls.SetCellValue(59, 6, new TFormula("='Inputs TOT advanced'!F99"));

            fmt = xls.GetCellVisibleFormatDef(59, 7);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(59, 7, xls.AddFormat(fmt));
            xls.SetCellValue(59, 7, new TFormula("='Inputs TOT advanced'!F99"));

            fmt = xls.GetCellVisibleFormatDef(59, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(59, 8, xls.AddFormat(fmt));
            xls.SetCellValue(59, 8, new TFormula("='Inputs TOT advanced'!F99"));

            fmt = xls.GetCellVisibleFormatDef(59, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(59, 9, xls.AddFormat(fmt));
            xls.SetCellValue(59, 9, new TFormula("='Inputs TOT advanced'!F126"));

            fmt = xls.GetCellVisibleFormatDef(59, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(59, 10, xls.AddFormat(fmt));
            xls.SetCellValue(59, 10, new TFormula("='Inputs TOT advanced'!F126"));

            fmt = xls.GetCellVisibleFormatDef(60, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.HAlignment = THFlxAlignment.left;
            xls.SetCellFormat(60, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 2);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 3);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(60, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(60, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(60, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(60, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(60, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(60, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(60, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(60, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(60, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(61, 1, xls.AddFormat(fmt));
            xls.SetCellValue(61, 1, "Valor mano de obra cosecha");

            fmt = xls.GetCellVisibleFormatDef(61, 2);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 2, xls.AddFormat(fmt));
            xls.SetCellValue(61, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(61, 3);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 3, xls.AddFormat(fmt));
            xls.SetCellValue(61, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(61, 4);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 4, xls.AddFormat(fmt));
            xls.SetCellValue(61, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(61, 5);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 5, xls.AddFormat(fmt));
            xls.SetCellValue(61, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(61, 6);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 6, xls.AddFormat(fmt));
            xls.SetCellValue(61, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(61, 7);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 7, xls.AddFormat(fmt));
            xls.SetCellValue(61, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(61, 8);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 8, xls.AddFormat(fmt));
            xls.SetCellValue(61, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(61, 9);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 9, xls.AddFormat(fmt));
            xls.SetCellValue(61, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(61, 10);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(61, 10, xls.AddFormat(fmt));
            xls.SetCellValue(61, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(61, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(61, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(61, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(62, 1, xls.AddFormat(fmt));
            xls.SetCellValue(62, 1, "Recoleccion de café");

            fmt = xls.GetCellVisibleFormatDef(62, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 4, xls.AddFormat(fmt));
            xls.SetCellValue(62, 4, new TFormula("='Inputs TOT advanced'!F75"));

            fmt = xls.GetCellVisibleFormatDef(62, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 5, xls.AddFormat(fmt));
            xls.SetCellValue(62, 5, new TFormula("='Inputs TOT advanced'!F75"));

            fmt = xls.GetCellVisibleFormatDef(62, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 6, xls.AddFormat(fmt));
            xls.SetCellValue(62, 6, new TFormula("='Inputs TOT advanced'!F102"));

            fmt = xls.GetCellVisibleFormatDef(62, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 7, xls.AddFormat(fmt));
            xls.SetCellValue(62, 7, new TFormula("='Inputs TOT advanced'!F102"));

            fmt = xls.GetCellVisibleFormatDef(62, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 8, xls.AddFormat(fmt));
            xls.SetCellValue(62, 8, new TFormula("='Inputs TOT advanced'!F102"));

            fmt = xls.GetCellVisibleFormatDef(62, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 9, xls.AddFormat(fmt));
            xls.SetCellValue(62, 9, new TFormula("='Inputs TOT advanced'!F129"));

            fmt = xls.GetCellVisibleFormatDef(62, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(62, 10, xls.AddFormat(fmt));
            xls.SetCellValue(62, 10, new TFormula("='Inputs TOT advanced'!F129"));

            fmt = xls.GetCellVisibleFormatDef(62, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(62, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(62, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(63, 1, xls.AddFormat(fmt));
            xls.SetCellValue(63, 1, "Otros:");

            fmt = xls.GetCellVisibleFormatDef(63, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 4);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 4, xls.AddFormat(fmt));
            xls.SetCellValue(63, 4, new TFormula("='Inputs TOT advanced'!F76"));

            fmt = xls.GetCellVisibleFormatDef(63, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 5, xls.AddFormat(fmt));
            xls.SetCellValue(63, 5, new TFormula("='Inputs TOT advanced'!F76"));

            fmt = xls.GetCellVisibleFormatDef(63, 6);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 6, xls.AddFormat(fmt));
            xls.SetCellValue(63, 6, new TFormula("='Inputs TOT advanced'!F103"));

            fmt = xls.GetCellVisibleFormatDef(63, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 7, xls.AddFormat(fmt));
            xls.SetCellValue(63, 7, new TFormula("='Inputs TOT advanced'!F103"));

            fmt = xls.GetCellVisibleFormatDef(63, 8);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 8, xls.AddFormat(fmt));
            xls.SetCellValue(63, 8, new TFormula("='Inputs TOT advanced'!F103"));

            fmt = xls.GetCellVisibleFormatDef(63, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 9, xls.AddFormat(fmt));
            xls.SetCellValue(63, 9, new TFormula("='Inputs TOT advanced'!F130"));

            fmt = xls.GetCellVisibleFormatDef(63, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(63, 10, xls.AddFormat(fmt));
            xls.SetCellValue(63, 10, new TFormula("='Inputs TOT advanced'!F130"));

            fmt = xls.GetCellVisibleFormatDef(63, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(63, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(63, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Locked = false;
            fmt.WrapText = true;
            xls.SetCellFormat(64, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(64, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(64, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(64, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 1);
            fmt.Font.Color = TExcelColor.FromTheme(TThemeColor.Background1);
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(65, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(65, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(65, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(65, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(66, 1, xls.AddFormat(fmt));
            xls.SetCellValue(66, 1, "Valor mano de obra para beneficio");

            fmt = xls.GetCellVisibleFormatDef(66, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(66, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(66, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(66, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(67, 1, xls.AddFormat(fmt));
            xls.SetCellValue(67, 1, "Beneficio humedo ");

            fmt = xls.GetCellVisibleFormatDef(67, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(67, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(67, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(67, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 1, xls.AddFormat(fmt));
            xls.SetCellValue(68, 1, "Despulpado y Fermentado");

            fmt = xls.GetCellVisibleFormatDef(68, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 4, xls.AddFormat(fmt));
            xls.SetCellValue(68, 4, new TFormula("='Inputs TOT advanced'!F80"));

            fmt = xls.GetCellVisibleFormatDef(68, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 5, xls.AddFormat(fmt));
            xls.SetCellValue(68, 5, new TFormula("='Inputs TOT advanced'!F80"));

            fmt = xls.GetCellVisibleFormatDef(68, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 6, xls.AddFormat(fmt));
            xls.SetCellValue(68, 6, new TFormula("='Inputs TOT advanced'!F107"));

            fmt = xls.GetCellVisibleFormatDef(68, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 7, xls.AddFormat(fmt));
            xls.SetCellValue(68, 7, new TFormula("='Inputs TOT advanced'!F107"));

            fmt = xls.GetCellVisibleFormatDef(68, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 8, xls.AddFormat(fmt));
            xls.SetCellValue(68, 8, new TFormula("='Inputs TOT advanced'!F107"));

            fmt = xls.GetCellVisibleFormatDef(68, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 9, xls.AddFormat(fmt));
            xls.SetCellValue(68, 9, new TFormula("='Inputs TOT advanced'!F134"));

            fmt = xls.GetCellVisibleFormatDef(68, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(68, 10, xls.AddFormat(fmt));
            xls.SetCellValue(68, 10, new TFormula("='Inputs TOT advanced'!F134"));

            fmt = xls.GetCellVisibleFormatDef(68, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(68, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(68, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 1, xls.AddFormat(fmt));
            xls.SetCellValue(69, 1, "Lavado");

            fmt = xls.GetCellVisibleFormatDef(69, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 4, xls.AddFormat(fmt));
            xls.SetCellValue(69, 4, new TFormula("='Inputs TOT advanced'!F81"));

            fmt = xls.GetCellVisibleFormatDef(69, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 5, xls.AddFormat(fmt));
            xls.SetCellValue(69, 5, new TFormula("='Inputs TOT advanced'!F81"));

            fmt = xls.GetCellVisibleFormatDef(69, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 6, xls.AddFormat(fmt));
            xls.SetCellValue(69, 6, new TFormula("='Inputs TOT advanced'!F108"));

            fmt = xls.GetCellVisibleFormatDef(69, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 7, xls.AddFormat(fmt));
            xls.SetCellValue(69, 7, new TFormula("='Inputs TOT advanced'!F108"));

            fmt = xls.GetCellVisibleFormatDef(69, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 8, xls.AddFormat(fmt));
            xls.SetCellValue(69, 8, new TFormula("='Inputs TOT advanced'!F108"));

            fmt = xls.GetCellVisibleFormatDef(69, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 9, xls.AddFormat(fmt));
            xls.SetCellValue(69, 9, new TFormula("='Inputs TOT advanced'!F135"));

            fmt = xls.GetCellVisibleFormatDef(69, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(69, 10, xls.AddFormat(fmt));
            xls.SetCellValue(69, 10, new TFormula("='Inputs TOT advanced'!F135"));

            fmt = xls.GetCellVisibleFormatDef(69, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(69, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(69, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xCC, 0xFF, 0xCC);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(70, 1, xls.AddFormat(fmt));
            xls.SetCellValue(70, 1, "Beneficio seco");

            fmt = xls.GetCellVisibleFormatDef(70, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(70, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(70, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(70, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(70, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 1, xls.AddFormat(fmt));
            xls.SetCellValue(71, 1, "Secado");

            fmt = xls.GetCellVisibleFormatDef(71, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 4, xls.AddFormat(fmt));
            xls.SetCellValue(71, 4, new TFormula("='Inputs TOT advanced'!F82"));

            fmt = xls.GetCellVisibleFormatDef(71, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 5, xls.AddFormat(fmt));
            xls.SetCellValue(71, 5, new TFormula("='Inputs TOT advanced'!F82"));

            fmt = xls.GetCellVisibleFormatDef(71, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 6, xls.AddFormat(fmt));
            xls.SetCellValue(71, 6, new TFormula("='Inputs TOT advanced'!F109"));

            fmt = xls.GetCellVisibleFormatDef(71, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 7, xls.AddFormat(fmt));
            xls.SetCellValue(71, 7, new TFormula("='Inputs TOT advanced'!F109"));

            fmt = xls.GetCellVisibleFormatDef(71, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 8, xls.AddFormat(fmt));
            xls.SetCellValue(71, 8, new TFormula("='Inputs TOT advanced'!F109"));

            fmt = xls.GetCellVisibleFormatDef(71, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 9, xls.AddFormat(fmt));
            xls.SetCellValue(71, 9, new TFormula("='Inputs TOT advanced'!F136"));

            fmt = xls.GetCellVisibleFormatDef(71, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(71, 10, xls.AddFormat(fmt));
            xls.SetCellValue(71, 10, new TFormula("='Inputs TOT advanced'!F136"));

            fmt = xls.GetCellVisibleFormatDef(71, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(71, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(71, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 1, xls.AddFormat(fmt));
            xls.SetCellValue(72, 1, "Zarandeo");

            fmt = xls.GetCellVisibleFormatDef(72, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 4, xls.AddFormat(fmt));
            xls.SetCellValue(72, 4, new TFormula("='Inputs TOT advanced'!F83"));

            fmt = xls.GetCellVisibleFormatDef(72, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 5, xls.AddFormat(fmt));
            xls.SetCellValue(72, 5, new TFormula("='Inputs TOT advanced'!F83"));

            fmt = xls.GetCellVisibleFormatDef(72, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 6, xls.AddFormat(fmt));
            xls.SetCellValue(72, 6, new TFormula("='Inputs TOT advanced'!F110"));

            fmt = xls.GetCellVisibleFormatDef(72, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 7, xls.AddFormat(fmt));
            xls.SetCellValue(72, 7, new TFormula("='Inputs TOT advanced'!F110"));

            fmt = xls.GetCellVisibleFormatDef(72, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 8, xls.AddFormat(fmt));
            xls.SetCellValue(72, 8, new TFormula("='Inputs TOT advanced'!F110"));

            fmt = xls.GetCellVisibleFormatDef(72, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 9, xls.AddFormat(fmt));
            xls.SetCellValue(72, 9, new TFormula("='Inputs TOT advanced'!F137"));

            fmt = xls.GetCellVisibleFormatDef(72, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(72, 10, xls.AddFormat(fmt));
            xls.SetCellValue(72, 10, new TFormula("='Inputs TOT advanced'!F137"));

            fmt = xls.GetCellVisibleFormatDef(72, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(72, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(72, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 1, xls.AddFormat(fmt));
            xls.SetCellValue(73, 1, "Escojo Selección");

            fmt = xls.GetCellVisibleFormatDef(73, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 4, xls.AddFormat(fmt));
            xls.SetCellValue(73, 4, new TFormula("='Inputs TOT advanced'!F84"));

            fmt = xls.GetCellVisibleFormatDef(73, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 5, xls.AddFormat(fmt));
            xls.SetCellValue(73, 5, new TFormula("='Inputs TOT advanced'!F84"));

            fmt = xls.GetCellVisibleFormatDef(73, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 6, xls.AddFormat(fmt));
            xls.SetCellValue(73, 6, new TFormula("='Inputs TOT advanced'!F111"));

            fmt = xls.GetCellVisibleFormatDef(73, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 7, xls.AddFormat(fmt));
            xls.SetCellValue(73, 7, new TFormula("='Inputs TOT advanced'!F111"));

            fmt = xls.GetCellVisibleFormatDef(73, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 8, xls.AddFormat(fmt));
            xls.SetCellValue(73, 8, new TFormula("='Inputs TOT advanced'!F111"));

            fmt = xls.GetCellVisibleFormatDef(73, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 9, xls.AddFormat(fmt));
            xls.SetCellValue(73, 9, new TFormula("='Inputs TOT advanced'!F138"));

            fmt = xls.GetCellVisibleFormatDef(73, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(73, 10, xls.AddFormat(fmt));
            xls.SetCellValue(73, 10, new TFormula("='Inputs TOT advanced'!F138"));

            fmt = xls.GetCellVisibleFormatDef(73, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(73, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(73, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(74, 1, xls.AddFormat(fmt));
            xls.SetCellValue(74, 1, "Almacenamiento");

            fmt = xls.GetCellVisibleFormatDef(74, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 4, xls.AddFormat(fmt));
            xls.SetCellValue(74, 4, new TFormula("='Inputs TOT advanced'!F85"));

            fmt = xls.GetCellVisibleFormatDef(74, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 5, xls.AddFormat(fmt));
            xls.SetCellValue(74, 5, new TFormula("='Inputs TOT advanced'!F85"));

            fmt = xls.GetCellVisibleFormatDef(74, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 6, xls.AddFormat(fmt));
            xls.SetCellValue(74, 6, new TFormula("='Inputs TOT advanced'!F112"));

            fmt = xls.GetCellVisibleFormatDef(74, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 7, xls.AddFormat(fmt));
            xls.SetCellValue(74, 7, new TFormula("='Inputs TOT advanced'!F112"));

            fmt = xls.GetCellVisibleFormatDef(74, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 8, xls.AddFormat(fmt));
            xls.SetCellValue(74, 8, new TFormula("='Inputs TOT advanced'!F112"));

            fmt = xls.GetCellVisibleFormatDef(74, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 9, xls.AddFormat(fmt));
            xls.SetCellValue(74, 9, new TFormula("='Inputs TOT advanced'!F139"));

            fmt = xls.GetCellVisibleFormatDef(74, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(74, 10, xls.AddFormat(fmt));
            xls.SetCellValue(74, 10, new TFormula("='Inputs TOT advanced'!F139"));

            fmt = xls.GetCellVisibleFormatDef(74, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(74, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(74, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(75, 1, xls.AddFormat(fmt));
            xls.SetCellValue(75, 1, "Aguas Miel");

            fmt = xls.GetCellVisibleFormatDef(75, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 4, xls.AddFormat(fmt));
            xls.SetCellValue(75, 4, new TFormula("='Inputs TOT advanced'!F86"));

            fmt = xls.GetCellVisibleFormatDef(75, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 5, xls.AddFormat(fmt));
            xls.SetCellValue(75, 5, new TFormula("='Inputs TOT advanced'!F86"));

            fmt = xls.GetCellVisibleFormatDef(75, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 6, xls.AddFormat(fmt));
            xls.SetCellValue(75, 6, new TFormula("='Inputs TOT advanced'!F113"));

            fmt = xls.GetCellVisibleFormatDef(75, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 7, xls.AddFormat(fmt));
            xls.SetCellValue(75, 7, new TFormula("='Inputs TOT advanced'!F113"));

            fmt = xls.GetCellVisibleFormatDef(75, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 8, xls.AddFormat(fmt));
            xls.SetCellValue(75, 8, new TFormula("='Inputs TOT advanced'!F113"));

            fmt = xls.GetCellVisibleFormatDef(75, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 9, xls.AddFormat(fmt));
            xls.SetCellValue(75, 9, new TFormula("='Inputs TOT advanced'!F140"));

            fmt = xls.GetCellVisibleFormatDef(75, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(75, 10, xls.AddFormat(fmt));
            xls.SetCellValue(75, 10, new TFormula("='Inputs TOT advanced'!F140"));

            fmt = xls.GetCellVisibleFormatDef(75, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(75, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(75, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(76, 1, xls.AddFormat(fmt));
            xls.SetCellValue(76, 1, "Manejo de Pulpa");

            fmt = xls.GetCellVisibleFormatDef(76, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 4, xls.AddFormat(fmt));
            xls.SetCellValue(76, 4, new TFormula("='Inputs TOT advanced'!F87"));

            fmt = xls.GetCellVisibleFormatDef(76, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 5, xls.AddFormat(fmt));
            xls.SetCellValue(76, 5, new TFormula("='Inputs TOT advanced'!F87"));

            fmt = xls.GetCellVisibleFormatDef(76, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 6, xls.AddFormat(fmt));
            xls.SetCellValue(76, 6, new TFormula("='Inputs TOT advanced'!F114"));

            fmt = xls.GetCellVisibleFormatDef(76, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 7, xls.AddFormat(fmt));
            xls.SetCellValue(76, 7, new TFormula("='Inputs TOT advanced'!F114"));

            fmt = xls.GetCellVisibleFormatDef(76, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 8, xls.AddFormat(fmt));
            xls.SetCellValue(76, 8, new TFormula("='Inputs TOT advanced'!F114"));

            fmt = xls.GetCellVisibleFormatDef(76, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 9, xls.AddFormat(fmt));
            xls.SetCellValue(76, 9, new TFormula("='Inputs TOT advanced'!F141"));

            fmt = xls.GetCellVisibleFormatDef(76, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(76, 10, xls.AddFormat(fmt));
            xls.SetCellValue(76, 10, new TFormula("='Inputs TOT advanced'!F141"));

            fmt = xls.GetCellVisibleFormatDef(76, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(76, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(76, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 1);
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Style = TFlxFontStyles.Italic;
            xls.SetCellFormat(77, 1, xls.AddFormat(fmt));
            xls.SetCellValue(77, 1, "Otros");

            fmt = xls.GetCellVisibleFormatDef(77, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 4, xls.AddFormat(fmt));
            xls.SetCellValue(77, 4, new TFormula("='Inputs TOT advanced'!F88"));

            fmt = xls.GetCellVisibleFormatDef(77, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 5, xls.AddFormat(fmt));
            xls.SetCellValue(77, 5, new TFormula("='Inputs TOT advanced'!F88"));

            fmt = xls.GetCellVisibleFormatDef(77, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 6, xls.AddFormat(fmt));
            xls.SetCellValue(77, 6, new TFormula("='Inputs TOT advanced'!F115"));

            fmt = xls.GetCellVisibleFormatDef(77, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 7, xls.AddFormat(fmt));
            xls.SetCellValue(77, 7, new TFormula("='Inputs TOT advanced'!F115"));

            fmt = xls.GetCellVisibleFormatDef(77, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 8, xls.AddFormat(fmt));
            xls.SetCellValue(77, 8, new TFormula("='Inputs TOT advanced'!F115"));

            fmt = xls.GetCellVisibleFormatDef(77, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 9, xls.AddFormat(fmt));
            xls.SetCellValue(77, 9, new TFormula("='Inputs TOT advanced'!F142"));

            fmt = xls.GetCellVisibleFormatDef(77, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(77, 10, xls.AddFormat(fmt));
            xls.SetCellValue(77, 10, new TFormula("='Inputs TOT advanced'!F142"));

            fmt = xls.GetCellVisibleFormatDef(77, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(77, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(77, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(78, 1, xls.AddFormat(fmt));
            xls.SetCellValue(78, 1, "Valor mano de obra para cuestiones administrativas");

            fmt = xls.GetCellVisibleFormatDef(78, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(78, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 11);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 11, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 12);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 12, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 13);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 13, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 14);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 14, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 15);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 15, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 16);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 16, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 17);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 17, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 18);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 18, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(78, 19);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            xls.SetCellFormat(78, 19, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Background1, -0.249977111117893);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(79, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 3);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 9);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(79, 10);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(79, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 3);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(80, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(80, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(81, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(81, 1, xls.AddFormat(fmt));
            xls.SetCellValue(81, 1, "Administración de su finca");

            fmt = xls.GetCellVisibleFormatDef(81, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 2, xls.AddFormat(fmt));
            xls.SetCellValue(81, 2, "Año 0");

            fmt = xls.GetCellVisibleFormatDef(81, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 3, xls.AddFormat(fmt));
            xls.SetCellValue(81, 3, "Año 1");

            fmt = xls.GetCellVisibleFormatDef(81, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 4, xls.AddFormat(fmt));
            xls.SetCellValue(81, 4, "Año 2");

            fmt = xls.GetCellVisibleFormatDef(81, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 5, xls.AddFormat(fmt));
            xls.SetCellValue(81, 5, "Año 3");

            fmt = xls.GetCellVisibleFormatDef(81, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 6, xls.AddFormat(fmt));
            xls.SetCellValue(81, 6, "Año 4");

            fmt = xls.GetCellVisibleFormatDef(81, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 7, xls.AddFormat(fmt));
            xls.SetCellValue(81, 7, "Año 5");

            fmt = xls.GetCellVisibleFormatDef(81, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 8, xls.AddFormat(fmt));
            xls.SetCellValue(81, 8, "Año 6");

            fmt = xls.GetCellVisibleFormatDef(81, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 9, xls.AddFormat(fmt));
            xls.SetCellValue(81, 9, "Año 7");

            fmt = xls.GetCellVisibleFormatDef(81, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(81, 10, xls.AddFormat(fmt));
            xls.SetCellValue(81, 10, "Año 8");

            fmt = xls.GetCellVisibleFormatDef(82, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(82, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(82, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(82, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 1);
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.799981688894314);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            xls.SetCellFormat(83, 1, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(83, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TUIColor.FromArgb(0xBF, 0xBF, 0xBF);
            fmt.FillPattern.BgColor = TUIColor.FromArgb(0x00, 0x00, 0x00);
            fmt.Format = "#,##0.00";
            fmt.HAlignment = THFlxAlignment.center;
            xls.SetCellFormat(83, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
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
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(84, 1, xls.AddFormat(fmt));
            xls.SetCellValue(84, 1, "Cuanto tiempo puede gastar Ud. Supervisando (no trabajando) actividades como limpias,"
            + " manejos, podas, obras conservación, cosecha etc");

            fmt = xls.GetCellVisibleFormatDef(84, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 3);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(84, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(84, 10, xls.AddFormat(fmt));
            xls.SetCellValue(85, 1, "Dias al año");

            fmt = xls.GetCellVisibleFormatDef(85, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 2, xls.AddFormat(fmt));
            xls.SetCellValue(85, 2, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 3, xls.AddFormat(fmt));
            xls.SetCellValue(85, 3, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 4, xls.AddFormat(fmt));
            xls.SetCellValue(85, 4, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 5, xls.AddFormat(fmt));
            xls.SetCellValue(85, 5, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 6, xls.AddFormat(fmt));
            xls.SetCellValue(85, 6, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 7, xls.AddFormat(fmt));
            xls.SetCellValue(85, 7, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 8, xls.AddFormat(fmt));
            xls.SetCellValue(85, 8, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 9, xls.AddFormat(fmt));
            xls.SetCellValue(85, 9, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(85, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(85, 10, xls.AddFormat(fmt));
            xls.SetCellValue(85, 10, new TFormula("='Inputs TOT advanced'!$F$412"));

            fmt = xls.GetCellVisibleFormatDef(86, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
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
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(86, 1, xls.AddFormat(fmt));
            xls.SetCellValue(86, 1, "Cuantos dias al mes gasta UD en cuestiones administrativas de su finca como llevar"
            + " las cuentas, pagar servicios etc.?");

            fmt = xls.GetCellVisibleFormatDef(86, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 3);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(86, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(86, 10, xls.AddFormat(fmt));
            xls.SetCellValue(87, 1, "Dias al año");

            fmt = xls.GetCellVisibleFormatDef(87, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 2, xls.AddFormat(fmt));
            xls.SetCellValue(87, 2, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 3, xls.AddFormat(fmt));
            xls.SetCellValue(87, 3, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 4, xls.AddFormat(fmt));
            xls.SetCellValue(87, 4, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 5, xls.AddFormat(fmt));
            xls.SetCellValue(87, 5, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 6, xls.AddFormat(fmt));
            xls.SetCellValue(87, 6, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 7, xls.AddFormat(fmt));
            xls.SetCellValue(87, 7, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 8, xls.AddFormat(fmt));
            xls.SetCellValue(87, 8, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 9, xls.AddFormat(fmt));
            xls.SetCellValue(87, 9, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(87, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(87, 10, xls.AddFormat(fmt));
            xls.SetCellValue(87, 10, new TFormula("='Inputs TOT advanced'!$F$413"));

            fmt = xls.GetCellVisibleFormatDef(88, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TExcelColor.Automatic;
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
            fmt.HAlignment = THFlxAlignment.left;
            fmt.WrapText = true;
            xls.SetCellFormat(88, 1, xls.AddFormat(fmt));
            xls.SetCellValue(88, 1, "Cuanto tiempo puede gastar Ud. al año en capacitar a la gente que contrata para las"
            + " diversas labores de la finca");

            fmt = xls.GetCellVisibleFormatDef(88, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 3);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(88, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(88, 10, xls.AddFormat(fmt));
            xls.SetCellValue(89, 1, "Dias al año");

            fmt = xls.GetCellVisibleFormatDef(89, 2);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 2, xls.AddFormat(fmt));
            xls.SetCellValue(89, 2, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 3, xls.AddFormat(fmt));
            xls.SetCellValue(89, 3, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 4, xls.AddFormat(fmt));
            xls.SetCellValue(89, 4, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 5, xls.AddFormat(fmt));
            xls.SetCellValue(89, 5, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 6, xls.AddFormat(fmt));
            xls.SetCellValue(89, 6, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 7, xls.AddFormat(fmt));
            xls.SetCellValue(89, 7, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 8, xls.AddFormat(fmt));
            xls.SetCellValue(89, 8, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 9, xls.AddFormat(fmt));
            xls.SetCellValue(89, 9, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(89, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(89, 10, xls.AddFormat(fmt));
            xls.SetCellValue(89, 10, new TFormula("='Inputs TOT advanced'!$F$414"));

            fmt = xls.GetCellVisibleFormatDef(90, 1);
            fmt.Font.Name = "Arial";
            fmt.Font.Style = TFlxFontStyles.Bold;
            fmt.Font.Scheme = TFontScheme.None;
            fmt.WrapText = true;
            xls.SetCellFormat(90, 1, xls.AddFormat(fmt));
            xls.SetCellValue(90, 1, "Cuanto puede gastar Ud. En costos extraordinarios tales como cubrir asistencias médicas"
            + " por accidentes de trabajo de sus trabajadores");

            fmt = xls.GetCellVisibleFormatDef(90, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 3);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(90, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(90, 10, xls.AddFormat(fmt));
            xls.SetCellValue(91, 1, "En Moneda Local");

            fmt = xls.GetCellVisibleFormatDef(91, 2);
            fmt.Font.Name = "Arial";
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.Font.Scheme = TFontScheme.None;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 2, xls.AddFormat(fmt));
            xls.SetCellValue(91, 2, new TFormula("='Inputs TOT advanced'!$F$415"));

            fmt = xls.GetCellVisibleFormatDef(91, 3);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 3, xls.AddFormat(fmt));
            xls.SetCellValue(91, 3, new TFormula("=B91"));

            fmt = xls.GetCellVisibleFormatDef(91, 4);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 4, xls.AddFormat(fmt));
            xls.SetCellValue(91, 4, new TFormula("=C91"));

            fmt = xls.GetCellVisibleFormatDef(91, 5);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 5, xls.AddFormat(fmt));
            xls.SetCellValue(91, 5, new TFormula("=D91"));

            fmt = xls.GetCellVisibleFormatDef(91, 6);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 6, xls.AddFormat(fmt));
            xls.SetCellValue(91, 6, new TFormula("=E91"));

            fmt = xls.GetCellVisibleFormatDef(91, 7);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 7, xls.AddFormat(fmt));
            xls.SetCellValue(91, 7, new TFormula("=F91"));

            fmt = xls.GetCellVisibleFormatDef(91, 8);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 8, xls.AddFormat(fmt));
            xls.SetCellValue(91, 8, new TFormula("=G91"));

            fmt = xls.GetCellVisibleFormatDef(91, 9);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 9, xls.AddFormat(fmt));
            xls.SetCellValue(91, 9, new TFormula("=H91"));

            fmt = xls.GetCellVisibleFormatDef(91, 10);
            fmt.Font.Color = TUIColor.FromArgb(0x00, 0x00, 0xFF);
            fmt.FillPattern.Pattern = TFlxPatternStyle.Solid;
            fmt.FillPattern.FgColor = TExcelColor.FromTheme(TThemeColor.Accent3, 0.599993896298105);
            fmt.FillPattern.BgColor = TExcelColor.Automatic;
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(91, 10, xls.AddFormat(fmt));
            xls.SetCellValue(91, 10, new TFormula("=I91"));

            fmt = xls.GetCellVisibleFormatDef(92, 1);
            fmt.Font.Color = TUIColor.FromArgb(0xFF, 0x00, 0x00);
            xls.SetCellFormat(92, 1, xls.AddFormat(fmt));
            xls.SetCellValue(92, 1, "Atarlo a presupuesto C50");

            fmt = xls.GetCellVisibleFormatDef(92, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 3);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(92, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(92, 10, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 2);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 2, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 3);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 3, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 4);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 4, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 5);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 5, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 6);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 6, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 7);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 7, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 8);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 8, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 9);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 9, xls.AddFormat(fmt));

            fmt = xls.GetCellVisibleFormatDef(93, 10);
            fmt.Format = "#,##0.00";
            xls.SetCellFormat(93, 10, xls.AddFormat(fmt));

            //Cell selection and scroll position.
            xls.SelectCell(1, 20, false);
            xls.ScrollWindow(1, 9);

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

        }


    }
}
